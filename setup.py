"""Include workflow guidance in the release, with generated version provenance."""
import ast
import hashlib
import json
import shutil
from pathlib import Path

from setuptools import setup
from setuptools.command.build_py import build_py


class BuildGuidance(build_py):
    def run(self):
        super().run()
        root = Path(__file__).resolve().parent
        source = root / "src/msaccess_vcs_mcp"
        destination = Path(self.build_lib) / "msaccess_vcs_mcp/guidance"
        # Replace only this build-owned directory so removed references cannot linger.
        if destination.exists():
            assert destination.resolve().is_relative_to(Path(self.build_lib).resolve())
            shutil.rmtree(destination)
        shutil.copytree(root / "skills", destination)
        shutil.copy2(source / "workflow_requirement.json", destination)
        version_tree = ast.parse((source / "__init__.py").read_text(encoding="utf-8"))
        server_version = next(
            ast.literal_eval(node.value)
            for node in version_tree.body
            if isinstance(node, ast.Assign)
            and any(isinstance(target, ast.Name) and target.id == "__version__" for target in node.targets)
        )
        manifest = {
            "distribution": self.distribution.get_name(),
            "package_version": self.distribution.get_version(),
            "mcp_version": server_version,
            "files": {
                path.relative_to(destination).as_posix(): hashlib.sha256(path.read_bytes()).hexdigest()
                for path in sorted(destination.rglob("*")) if path.is_file()
            },
        }
        (destination / "release.json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")


setup(cmdclass={"build_py": BuildGuidance})
