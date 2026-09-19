import os
import sys
import base64
import hashlib
import shutil
from pathlib import Path

def assemble_package():
    root = Path(".")
    package = root / "production-package"
    (package / "engine").mkdir(parents=True, exist_ok=True)
    (package / "config").mkdir(parents=True, exist_ok=True)
    (package / "vba").mkdir(parents=True, exist_ok=True)
    (package / "source").mkdir(parents=True, exist_ok=True)
    
    # Assemble base workbook
    chunks = sorted(list((root / "excel" / "base_chunks").glob("*.part*")))
    if len(chunks) != 10:
        raise ValueError(f"Expected 10 base chunks, found {len(chunks)}")
    
    encoded = "".join(p.read_text(encoding="utf-8") for p in chunks)
    data = base64.b64decode(encoded)
    
    expected_hash = "9225a41707c2dd483499a697d8358556364d42a2a35be9a09027f170c45f7682"
    actual_hash = hashlib.sha256(data).hexdigest().lower()
    if actual_hash != expected_hash:
        raise ValueError(f"Base workbook hash mismatch: {actual_hash} != {expected_hash}")
        
    wb_dest = package / "source" / "NMDC_Document_Index_Base.xlsx"
    wb_dest.write_bytes(data)
    print("Base workbook assembled & SHA256 verified:", actual_hash)
    
    # Copy engine
    shutil.copy2(root / "dist" / "nmdc_index_engine.exe", package / "engine" / "nmdc_index_engine.exe")
    
    # Copy config
    for cfg in ["classification_rules.csv", "project_identity_overrides.csv", "source_exclusions.csv"]:
        shutil.copy2(root / "config" / cfg, package / "config" / cfg)
        
    # Copy vba
    for bas in (root / "excel" / "vba").glob("*.bas"):
        shutil.copy2(bas, package / "vba" / bas.name)
    for frm in (root / "excel" / "vba").glob("*.frm"):
        shutil.copy2(frm, package / "vba" / frm.name)
    for frx in (root / "excel" / "vba").glob("*.frx"):
        shutil.copy2(frx, package / "vba" / frx.name)
        
    # Copy scripts
    shutil.copy2(root / "packaging" / "Create_NMDC_Document_Index.vbs", package / "Create_NMDC_Document_Index.vbs")
    shutil.copy2(root / "packaging" / "README.md", package / "README.md")
    
    print("Production package successfully assembled in:", package.resolve())

if __name__ == "__main__":
    assemble_package()
