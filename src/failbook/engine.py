"""Formula recalculation backed by LibreOffice.

Acceptance checks must not trust the cached values stored in a workbook: a tool
that writes formulas (openpyxl, for example) stores no results at all, and a
tool that edits XML directly can leave stale ones behind. Every value a check
looks at is therefore recomputed from scratch by a real spreadsheet engine.

LibreOffice is used because it is free, scriptable, cross-platform and close
enough to Excel for the functions the cases rely on. The engine name and
version are recorded in every result so that verdicts can be traced.
"""

from __future__ import annotations

import atexit
import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Iterable

_PROFILE_XCU = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry"
           xmlns:xs="http://www.w3.org/2001/XMLSchema"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <item oor:path="/org.openoffice.Office.Calc/Formula/Load">
    <prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop>
  </item>
  <item oor:path="/org.openoffice.Office.Calc/Formula/Load">
    <prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop>
  </item>
</oor:items>
"""

_MAC_PATH = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
_WIN_PATHS = (
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
)


class EngineUnavailable(RuntimeError):
    """Raised when no recalculation engine can be found."""


def find_soffice() -> str | None:
    env = os.environ.get("FAILBOOK_SOFFICE")
    if env:
        return env
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for candidate in (_MAC_PATH, *_WIN_PATHS):
        if os.path.exists(candidate):
            return candidate
    return None


class LibreOfficeEngine:
    """Recalculate .xlsx files by round-tripping them through headless LibreOffice."""

    name = "libreoffice"

    def __init__(self, soffice: str | None = None, timeout: float = 180.0):
        self.soffice = soffice or find_soffice()
        if not self.soffice:
            raise EngineUnavailable(
                "LibreOffice was not found. Install it (e.g. `apt install libreoffice-calc`, "
                "`brew install --cask libreoffice`) or set FAILBOOK_SOFFICE to the soffice binary."
            )
        self.timeout = timeout
        self._profile: Path | None = None
        self._version: str | None = None

    # -- profile -----------------------------------------------------------------
    def _profile_dir(self) -> Path:
        if self._profile is None:
            root = Path(tempfile.mkdtemp(prefix="failbook-lo-"))
            (root / "user").mkdir(parents=True)
            (root / "user" / "registrymodifications.xcu").write_text(_PROFILE_XCU, encoding="utf-8")
            self._profile = root
            atexit.register(shutil.rmtree, root, True)
        return self._profile

    @property
    def version(self) -> str:
        if self._version is None:
            try:
                out = subprocess.run(
                    [self.soffice, "--version"], capture_output=True, text=True, timeout=60
                ).stdout.strip()
            except Exception:  # pragma: no cover - diagnostic only
                out = ""
            self._version = out.splitlines()[0] if out else "LibreOffice (unknown version)"
        return self._version

    def describe(self) -> str:
        return self.version

    # -- recalculation -------------------------------------------------------------
    def recalc(self, files: Iterable[Path], outdir: Path) -> dict[Path, Path]:
        """Recalculate every file and write the results to ``outdir``.

        Returns a mapping from each source path to the recalculated copy. Files
        are converted in a single LibreOffice process, so batching is much faster
        than converting one file at a time. Source base names must be unique.
        """
        files = [Path(f).resolve() for f in files]
        if not files:
            return {}
        names = [f.name for f in files]
        if len(set(names)) != len(names):
            raise ValueError("recalc() needs files with unique base names")
        outdir = Path(outdir).resolve()
        outdir.mkdir(parents=True, exist_ok=True)
        profile = self._profile_dir().as_uri()
        cmd = [
            self.soffice,
            "--headless",
            "--norestore",
            "--nolockcheck",
            "--nodefault",
            f"-env:UserInstallation={profile}",
            "--convert-to",
            "xlsx:Calc Office Open XML",
            "--outdir",
            str(outdir),
            *[str(f) for f in files],
        ]
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=self.timeout)
        result: dict[Path, Path] = {}
        missing = []
        for f in files:
            target = outdir / (f.stem + ".xlsx")
            if target.exists():
                result[f] = target
            else:
                missing.append(f.name)
        if missing:
            raise RuntimeError(
                "LibreOffice could not recalculate: "
                + ", ".join(missing)
                + f"\nstdout: {proc.stdout[-2000:]}\nstderr: {proc.stderr[-2000:]}"
            )
        return result


_DEFAULT: LibreOfficeEngine | None = None


def default_engine() -> LibreOfficeEngine:
    global _DEFAULT
    if _DEFAULT is None:
        _DEFAULT = LibreOfficeEngine()
    return _DEFAULT
