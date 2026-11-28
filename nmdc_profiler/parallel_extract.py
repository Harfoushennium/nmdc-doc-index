"""Bounded per-source worker pool; authoritative staging and flag order remain in parent."""
from concurrent.futures import ProcessPoolExecutor
import multiprocessing
from .source_access import install_resilient_source_access
from .layout_compat import install_layout_compatibility
_CATALOG=None
def _init(catalog):
    global _CATALOG
    install_resilient_source_access();install_layout_compatibility()
    _CATALOG=catalog
def _extract(relative):
    rows=list(_CATALOG.process(_CATALOG.data_dir/relative,relative))
    return rows,list(_CATALOG.drain_processor_flags())

class FullScanPool:
    def __init__(self,catalog,workers):
        self.catalog=catalog
        self.sources=sorted((k for k in catalog.work_items if catalog.selection_statuses.get(k)=="SELECTED"),key=str.casefold)
        self.positions={k:i for i,k in enumerate(self.sources)}
        self.workers=workers
        self.pending={};self.next=0;self.floor=0;self.pool=None
    def __enter__(self):
        self.pool=ProcessPoolExecutor(max_workers=self.workers,mp_context=multiprocessing.get_context("spawn"),initializer=_init,initargs=(self.catalog,))
        return self
    def _fill(self):
        while self.next<len(self.sources) and len(self.pending)<self.workers*2:
            idx=self.next;self.next+=1
            if idx>=self.floor:self.pending[idx]=self.pool.submit(_extract,self.sources[idx])
    def process(self,path,relative):
        idx=self.positions.get(relative)
        if idx is None or idx<self.floor:return self.catalog.process(path,relative)
        self.floor=idx
        for k in list(self.pending):
            if k<idx:self.pending.pop(k).cancel()
        self._fill()
        future=self.pending.pop(idx)
        self.floor=idx+1
        try:
            rows,flags=future.result()
        except Exception:
            # Reproduce the existing serial parser's behavior; no worker-side
            # writes or partial records are accepted if the native worker fails.
            return self.catalog.process(path,relative)
        self.catalog._processor_flags.extend(flags)
        return rows
    def __exit__(self,*args):
        if self.pool is not None:
            self.pool.shutdown(wait=True,cancel_futures=True)
