"""CLI entry point for NMDC Document Index Cycle-1 profiling."""
from __future__ import annotations
import argparse
from pathlib import Path
from typing import Optional,Sequence
from nmdc_profiler import run

ROOT=Path(__file__).resolve().parent

def main(argv:Optional[Sequence[str]]=None)->int:
    p=argparse.ArgumentParser(description="Read-only NMDC Cycle-1 workbook profiler")
    p.add_argument("--data-dir",type=Path,default=ROOT/"DATA")
    p.add_argument("--output-dir",type=Path,default=ROOT/"outputs"/"cycle1")
    p.add_argument("--rules",type=Path,default=ROOT/"config"/"classification_rules.csv")
    a=p.parse_args(argv); result=run(a.data_dir,a.output_dir,a.rules)
    print(f"Profiled {len(result['workbooks'])} workbooks; {len(result['classification_rows'])} classification rows")
    return 0

if __name__=="__main__": raise SystemExit(main())
