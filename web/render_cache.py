"""In-memory cache of the expensive part of turning a report (plus render
options) into a slide preview: the `.pptx` render and its LibreOffice-
converted `.pdf`. That conversion is where essentially all of a preview's
~10s cost lives -- extracting one more page from an already-converted PDF
via `pdftoppm` is well under a second -- so caching the PDF is what makes
paging back through a multi-slide render (or re-previewing after an
unrelated edit reuses the same content) fast on the *second* look, without
touching how long the *first* one takes.

Keyed by a hash of everything that determines the rendered content (see
`make_key` -- called with the same fields `/api/render`/`/api/preview`
already take), so there's no separate invalidation logic to get wrong: any
edit that changes the output changes the key, and the old entry for the
previous content just ages out on its own (see eviction below) rather than
needing to be explicitly cleared.

`/api/render` and `/api/preview` share one cache, but only `/api/preview`
ever *populates* it (`get_or_compute`) -- an entry always carries both the
`.pptx` and its converted `.pdf` together, and populating one from
`/api/render` alone would mean either doing the PDF conversion someone
only asked to download (adding ~10s to a request that's normally
near-instant) or caching a partial entry with no PDF, which then needs its
own "fill in the missing half later" dance. Simpler to accept the
asymmetry: `/api/render` only *reads* (`peek`) an entry a prior preview
already computed, and renders fresh (uncached, exactly like before this
existed) if there isn't one yet. A user previewing before downloading --
the common path this UI encourages -- gets the full benefit; a user who
downloads without ever previewing costs nothing extra either way.

Deliberately RAM, not disk: this container has no persistent volume (see
CLAUDE.md), and every render is already ephemeral by design -- there's
nothing here worth surviving a restart. The tradeoff that does matter is
memory footprint under concurrent load from multiple people: capped by
`max_bytes` (LRU eviction) rather than growing unbounded, since a handful
of large raw (non-distilled) templates could otherwise dominate it -- see
DEFAULT_MAX_BYTES below for sizing notes. This does *not* help multiple
people rendering *different* reports at the same time -- that's bounded by
how many `soffice` conversions the host can run at once, a separate,
not-yet-addressed concern.
"""

from __future__ import annotations

import hashlib
import json
import os
import threading
import time
from collections import OrderedDict
from dataclasses import dataclass, field
from typing import Any, Callable

# A no-template render is well under 1MB regardless of rack/slide count --
# measured directly: a 2-slide (3 racks + stats, 55 nodes) render came to
# ~290KB .pptx + ~400KB .pdf, and a 3-slide 60-node render was barely
# bigger than a 1-slide 5-node one, since node-photo assets are embedded
# once and referenced, not duplicated per node. The real size driver is a
# raw (non-distilled) uploaded template, capped elsewhere at 80MB
# (MAX_TEMPLATE_BYTES in app.py) -- this cap just bounds how many of those
# the cache holds onto at once. 512MB comfortably fits this host's spare
# RAM (multiple GB free in the container's host at time of writing) with
# real headroom; override via env var if that stops being true.
DEFAULT_MAX_BYTES = int(os.environ.get("QRACK_CACHE_MAX_BYTES", 512 * 1024 * 1024))
# Preview sessions are inherently short-lived (upload, tweak, preview,
# download, done) -- 15 minutes of inactivity is meant to comfortably
# outlast a real editing session without holding onto abandoned ones
# indefinitely.
DEFAULT_TTL_SECONDS = int(os.environ.get("QRACK_CACHE_TTL_SECONDS", 15 * 60))
# How long a concurrent caller waits for someone else's in-progress
# render of the *same* content before giving up and computing it itself
# (see get_or_compute) -- a little above the render pipeline's own
# timeout (PREVIEW_TIMEOUT_SECONDS in app.py) so a waiter doesn't bail
# out right as the real result was about to land.
DEFAULT_WAIT_TIMEOUT_SECONDS = 30.0


def make_key(**fields: Any) -> str:
    """Hashes whatever fields are passed -- callers pass exactly the
    request fields that determine the rendered output (report, rack
    label, visible stats, template, rack sizes), explicitly excluding
    anything that doesn't (e.g. which page a preview asks for)."""
    canonical = json.dumps(fields, sort_keys=True, default=str)
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


@dataclass
class CacheEntry:
    pptx_bytes: bytes
    pdf_bytes: bytes
    total_our_slides: int
    template_slide_count: int
    size: int = field(init=False)
    last_accessed: float = field(init=False)

    def __post_init__(self) -> None:
        self.size = len(self.pptx_bytes) + len(self.pdf_bytes)
        self.last_accessed = time.monotonic()


class RenderCache:
    """Thread-safe (FastAPI runs sync `def` endpoints in a thread pool,
    so this is genuinely concurrent, not just async-concurrent) LRU cache
    with a size cap and a TTL, plus single-flight dedup so concurrent
    requests for the *same* not-yet-cached content share one computation
    instead of each running their own `soffice` conversion."""

    def __init__(self, max_bytes: int = DEFAULT_MAX_BYTES, ttl_seconds: float = DEFAULT_TTL_SECONDS):
        self._max_bytes = max_bytes
        self._ttl_seconds = ttl_seconds
        self._lock = threading.Lock()
        self._entries: "OrderedDict[str, CacheEntry]" = OrderedDict()
        self._total_bytes = 0
        self._pending: dict[str, threading.Event] = {}

    def peek(self, key: str) -> CacheEntry | None:
        """Read-only lookup -- never claims responsibility for computing
        a miss. See module docstring for why `/api/render` uses this
        instead of `get_or_compute`."""
        with self._lock:
            self._evict_expired_locked()
            entry = self._entries.get(key)
            if entry is not None:
                entry.last_accessed = time.monotonic()
                self._entries.move_to_end(key)
            return entry

    def get_or_compute(
        self, key: str, compute_fn: Callable[[], CacheEntry], wait_timeout: float = DEFAULT_WAIT_TIMEOUT_SECONDS
    ) -> CacheEntry:
        """Returns the cached entry for `key`, calling `compute_fn()` to
        produce (and cache) it on a miss. A concurrent caller that misses
        on the same key while another is already computing it waits on
        that one instead of duplicating the work -- up to `wait_timeout`,
        past which (or if the in-progress call fails) it falls back to
        computing it itself rather than risking a stuck request forever
        over one dropped connection. `compute_fn` raising propagates
        directly to its own caller (nothing gets cached); any waiters
        wake up, see nothing cached, and each retry on their own -- not
        maximally efficient if this content is broken and everyone piles
        up trying it, but simple, and broken content isn't the common
        case this optimizes for.
        """
        while True:
            with self._lock:
                self._evict_expired_locked()
                entry = self._entries.get(key)
                if entry is not None:
                    entry.last_accessed = time.monotonic()
                    self._entries.move_to_end(key)
                    return entry
                event = self._pending.get(key)
                mine = event is None
                if mine:
                    event = threading.Event()
                    self._pending[key] = event

            if not mine:
                if not event.wait(timeout=wait_timeout):
                    continue  # timed out waiting -- try claiming it ourselves instead
                continue  # woke up -- loop around to re-check the cache

            try:
                entry = compute_fn()
            finally:
                with self._lock:
                    finished = self._pending.pop(key, None)
                if finished is not None:
                    finished.set()

            with self._lock:
                self._insert_locked(key, entry)
            return entry

    def _insert_locked(self, key: str, entry: CacheEntry) -> None:
        existing = self._entries.pop(key, None)
        if existing is not None:
            self._total_bytes -= existing.size
        self._entries[key] = entry
        self._total_bytes += entry.size
        self._evict_to_fit_locked()

    def _evict_expired_locked(self) -> None:
        now = time.monotonic()
        expired = [k for k, e in self._entries.items() if now - e.last_accessed > self._ttl_seconds]
        for k in expired:
            self._total_bytes -= self._entries.pop(k).size

    def _evict_to_fit_locked(self) -> None:
        while self._total_bytes > self._max_bytes and self._entries:
            _, evicted = self._entries.popitem(last=False)  # oldest = least-recently-used
            self._total_bytes -= evicted.size


CACHE = RenderCache()
