# -*- coding: utf-8 -*-
import runpy
from pathlib import Path

runpy.run_path(
    str(Path(__file__).parent.parent / "app_tmob_streamlit_estable_corregida.py"),
    run_name="__main__",
)
