"""Release archive contract: complete guidance and provenance survive wheel delivery."""
import hashlib
import json
import os
import shutil
import subprocess
import sys
import zipfile
from email.parser import Parser
from pathlib import Path

import pytest


@pytest.mark.parametrize("server_version", ["0.3.0-dev.18+verify2.0", "0.3.0-dev.18+package.check"])
def test_wheel_delivers_complete_guidance_with_derived_provenance(tmp_path, server_version):
    root = Path(__file__).resolve().parents[1]
    project = tmp_path / "project"
    project.mkdir()
    for name in ("pyproject.toml", "setup.py", "MANIFEST.in", "README.md", "LICENSE"):
        shutil.copy2(root / name, project / name)
    shutil.copytree(root / "skills", project / "skills")
    source = project / "src/msaccess_vcs_mcp"
    source.mkdir(parents=True)
    (source / "__init__.py").write_text(f'__version__ = "{server_version}"\n', encoding="utf-8")
    shutil.copy2(root / "src/msaccess_vcs_mcp/workflow_requirement.json", source)
    dist = project / "dist"
    env = os.environ.copy()
    env["TEMP"] = env["TMP"] = str(tmp_path)
    built = subprocess.run(
        [sys.executable, "-m", "build", "--wheel", "--no-isolation", "--outdir", str(dist)],
        cwd=project, env=env, capture_output=True, text=True, timeout=90,
    )
    assert built.returncode == 0, built.stdout + built.stderr
    with zipfile.ZipFile(next(dist.glob("*.whl"))) as archive:
        metadata_name = next(name for name in archive.namelist() if name.endswith(".dist-info/METADATA"))
        metadata = Parser().parsestr(archive.read(metadata_name).decode())
        prefix = "msaccess_vcs_mcp/guidance/"
        manifest = json.loads(archive.read(prefix + "release.json"))
        assert manifest["package_version"] == metadata["Version"]
        assert manifest["mcp_version"] == server_version
        assert manifest["distribution"] == "msaccess-vcs-mcp"
        expected = {path.relative_to(project / "skills").as_posix() for path in (project / "skills").rglob("*.md")}
        expected.add("workflow_requirement.json")
        assert set(manifest["files"]) == expected
        assert {name[len(prefix):] for name in archive.namelist() if name.startswith(prefix)} == expected | {"release.json"}
        for name in expected:
            contents = archive.read(prefix + name)
            assert hashlib.sha256(contents).hexdigest() == manifest["files"][name]
            original = source / name if name == "workflow_requirement.json" else project / "skills" / name
            assert contents == original.read_bytes()
