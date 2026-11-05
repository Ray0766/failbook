import shutil

import pytest

HAS_LIBREOFFICE = bool(shutil.which("soffice") or shutil.which("libreoffice"))
needs_libreoffice = pytest.mark.skipif(not HAS_LIBREOFFICE, reason="LibreOffice is not installed")
