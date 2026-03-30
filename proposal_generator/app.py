"""Diagnostic startup - tests imports one by one."""
import sys
import traceback
from pathlib import Path

_project_root = str(Path(__file__).resolve().parent.parent)
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

import streamlit as st

st.set_page_config(page_title="Diagnostic", page_icon="🔧", layout="wide")
st.title("🔧 Solar Proposal Generator - Startup Diagnostic")

errors = []

# Test 1: Basic imports
for mod_name in ["yaml", "openpyxl", "pandas", "pptx", "lxml", "PIL"]:
    try:
        __import__(mod_name)
        st.success(f"✅ {mod_name}")
    except Exception as e:
        st.error(f"❌ {mod_name}: {e}")
        errors.append(mod_name)

# Test 2: streamlit_sortables
try:
    from streamlit_sortables import sort_items
    st.success("✅ streamlit_sortables")
except Exception as e:
    st.error(f"❌ streamlit_sortables: {e}")
    errors.append("streamlit_sortables")

# Test 3: Project modules
for mod in [
    "proposal_generator.utils",
    "proposal_generator.ppa_calc",
    "proposal_generator.demand_calc",
    "proposal_generator.subsidy_calc",
    "proposal_generator.load_calc",
    "proposal_generator.fip_calc",
    "proposal_generator.generator",
]:
    try:
        __import__(mod)
        st.success(f"✅ {mod}")
    except Exception as e:
        st.error(f"❌ {mod}: {e}")
        st.code(traceback.format_exc())
        errors.append(mod)

# Test 4: composition_profiles.yaml
try:
    import yaml
    profiles_path = Path(__file__).parent / "composition_profiles.yaml"
    with open(profiles_path, encoding="utf-8") as f:
        data = yaml.safe_load(f)
    st.success(f"✅ composition_profiles.yaml ({len(data.get('profiles', {}))} profiles)")
except Exception as e:
    st.error(f"❌ composition_profiles.yaml: {e}")
    st.code(traceback.format_exc())
    errors.append("profiles")

# Test 5: Template file
tpl = Path(__file__).parent.parent / "templates"
if tpl.exists():
    files = list(tpl.iterdir())
    st.success(f"✅ templates/ ({len(files)} files)")
    for f in files:
        st.write(f"  - {f.name}")
else:
    st.warning(f"⚠️ templates/ not found at {tpl}")

# Summary
st.divider()
if errors:
    st.error(f"🚨 {len(errors)} modules failed: {', '.join(errors)}")
else:
    st.balloons()
    st.success("🎉 All checks passed! Ready to restore full app.")

st.info(f"Python {sys.version}\nPlatform: {sys.platform}\nCWD: {Path.cwd()}")
