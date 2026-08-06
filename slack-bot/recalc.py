"""Forces LibreOffice to recalculate and cache formula values in an xlsx
file, so Slack's inline file preview shows real numbers instead of blanks.
Best-effort: if it fails, the file is still valid and Excel/Sheets will
compute the formulas normally when a person opens it."""

import pathlib
import subprocess
import tempfile

_MACRO = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
    Sub RecalculateAndSave()
      ThisComponent.calculateAll()
      ThisComponent.store()
      ThisComponent.close(True)
    End Sub
</script:module>"""


def recalculate(xlsx_path, timeout=60):
    with tempfile.TemporaryDirectory() as profile_dir:
        profile_url = pathlib.Path(profile_dir).as_uri()

        subprocess.run(
            ["soffice", "--headless", "--terminate_after_init", f"-env:UserInstallation={profile_url}"],
            timeout=30, capture_output=True, check=True,
        )

        macro_dir = pathlib.Path(profile_dir) / "user" / "basic" / "Standard"
        macro_dir.mkdir(parents=True, exist_ok=True)
        (macro_dir / "Module1.xba").write_text(_MACRO)

        subprocess.run(
            [
                "soffice", "--headless", "--norestore",
                f"-env:UserInstallation={profile_url}",
                "vnd.sun.star.script:Standard.Module1.RecalculateAndSave?language=Basic&location=application",
                str(xlsx_path),
            ],
            timeout=timeout, capture_output=True, check=True,
        )
