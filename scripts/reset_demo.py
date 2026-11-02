"""Clear what the agent produced so the demo can run again. The inbox is left alone."""
import shutil
from pathlib import Path

HOME = Path(__file__).resolve().parent.parent

for folder in ("filed", "drafts"):
    shutil.rmtree(HOME / folder, ignore_errors=True)
(HOME / "students.xlsx").unlink(missing_ok=True)
print("Demo reset: removed filed/, drafts/ and students.xlsx")
