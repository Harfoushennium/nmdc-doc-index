from __future__ import annotations
from pathlib import Path
from typing import Dict
from .core import rel_posix
from .ooxml import profile_workbook
from .rules import load_rules
from .selection import assign_selection
from .classification import classification_rows,enrich_profiles_with_classification
from .reporting import write_outputs


def run(data_dir:Path,output_dir:Path,rules_path:Path)->Dict[str,object]:
    root=data_dir.parent; rules=load_rules(rules_path)
    files=sorted((p for p in data_dir.rglob("*") if p.is_file() and p.suffix.lower()==".xlsx"),key=lambda p:rel_posix(p,root).lower())
    workbooks=[profile_workbook(p,root,data_dir) for p in files]
    exact,versions=assign_selection(workbooks,rules)
    rows=classification_rows(workbooks,rules); enrich_profiles_with_classification(workbooks,rows)
    write_outputs(workbooks,rows,exact,versions,output_dir)
    return {"workbooks":workbooks,"classification_rows":rows,"exact_duplicate_groups":exact,"version_groups":versions}
