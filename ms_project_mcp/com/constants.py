"""Microsoft Project COM/VBA enum values, verified against the official
learn.microsoft.com "VBA reference for Microsoft Project".

We use late binding (``win32com.client.Dispatch``), so the ``pj*`` symbolic
names are NOT available at runtime — we pass these integers directly. Keep this
module as the single source of truth for every magic number in the codebase.

Gotchas that have bitten people (do not "fix" these to look tidier):
  * PjTaskLinkType is FinishToStart = 1, NOT 0.  FinishToFinish = 0.
  * FieldNameToFieldConstant returns a Long, not a string.
  * MSPDI XML is saved via the FormatID string "MSProject.xml" (no enum member).
  * BaselineSave lives on Application and selects the target via `Into`
    using PjSaveBaselineTo values (0, 11..20) — NOT the 0..10 PjBaselines values.
  * Task.Duration / Task.Work are in MINUTES.
"""

from __future__ import annotations


# --- Task.ConstraintType  (enum PjConstraint) ---------------------------------
class Constraint:
    ASAP = 0   # As Soon As Possible
    ALAP = 1   # As Late As Possible
    MSO = 2    # Must Start On
    MFO = 3    # Must Finish On
    SNET = 4   # Start No Earlier Than
    SNLT = 5   # Start No Later Than
    FNET = 6   # Finish No Earlier Than
    FNLT = 7   # Finish No Later Than


CONSTRAINT_NAMES = {
    0: "ASAP", 1: "ALAP", 2: "MustStartOn", 3: "MustFinishOn",
    4: "StartNoEarlierThan", 5: "StartNoLaterThan",
    6: "FinishNoEarlierThan", 7: "FinishNoLaterThan",
}

# Accept friendly user input (short codes and full names), case-insensitive.
CONSTRAINT_BY_NAME = {
    "ASAP": 0, "ALAP": 1,
    "MSO": 2, "MUSTSTARTON": 2,
    "MFO": 3, "MUSTFINISHON": 3,
    "SNET": 4, "STARTNOEARLIERTHAN": 4,
    "SNLT": 5, "STARTNOLATERTHAN": 5,
    "FNET": 6, "FINISHNOEARLIERTHAN": 6,
    "FNLT": 7, "FINISHNOLATERTHAN": 7,
}


# --- Task.Type  (enum PjTaskFixedType) ----------------------------------------
class TaskType:
    FIXED_UNITS = 0
    FIXED_DURATION = 1
    FIXED_WORK = 2


TASK_TYPE_NAMES = {0: "FixedUnits", 1: "FixedDuration", 2: "FixedWork"}
TASK_TYPE_BY_NAME = {
    "FIXEDUNITS": 0, "UNITS": 0,
    "FIXEDDURATION": 1, "DURATION": 1,
    "FIXEDWORK": 2, "WORK": 2,
}


# --- Resource.Type  (enum PjResourceTypes) ------------------------------------
class ResourceType:
    WORK = 0
    MATERIAL = 1
    COST = 2


RESOURCE_TYPE_NAMES = {0: "Work", 1: "Material", 2: "Cost"}
RESOURCE_TYPE_BY_NAME = {"WORK": 0, "MATERIAL": 1, "COST": 2}


# --- 2nd arg to Application.FieldNameToFieldConstant  (enum PjFieldType) -------
# NOTE: there is no pjAssignment member here.
class FieldType:
    TASK = 0
    RESOURCE = 1
    PROJECT = 2


# --- Dependency link type  (enum PjTaskLinkType) ------------------------------
# CAUTION: this ordering is counterintuitive. FinishToStart is 1.
class LinkType:
    FF = 0   # Finish-to-Finish
    FS = 1   # Finish-to-Start  (the normal/default dependency)
    SF = 2   # Start-to-Finish
    SS = 3   # Start-to-Start


LINK_TYPE_NAMES = {0: "FF", 1: "FS", 2: "SF", 3: "SS"}
LINK_TYPE_BY_NAME = {
    "FF": 0, "FINISHTOFINISH": 0,
    "FS": 1, "FINISHTOSTART": 1,
    "SF": 2, "STARTTOFINISH": 2,
    "SS": 3, "STARTTOSTART": 3,
}


# --- Baselines ----------------------------------------------------------------
# Read selector (Task.BaselineX lookups, BaselineClear) — enum PjBaselines.
class Baseline:
    B0 = 0
    B1 = 1
    B2 = 2
    B3 = 3
    B4 = 4
    B5 = 5
    B6 = 6
    B7 = 7
    B8 = 8
    B9 = 9
    B10 = 10


# Application.BaselineSave(..., Into=) target selector — enum PjSaveBaselineTo.
# Maps a human baseline number (0..10) to the Into value the API expects.
SAVE_BASELINE_INTO = {
    0: 0, 1: 11, 2: 12, 3: 13, 4: 14, 5: 15,
    6: 16, 7: 17, 8: 18, 9: 19, 10: 20,
}


# --- File formats -------------------------------------------------------------
class FileFormat:           # enum PjFileFormat (numeric `Format` arg)
    MPP = 0
    TXT = 3
    CSV = 4
    XLS = 5
    MPT = 11
    XLSX = 20
    XLSB = 21


# Preferred: pass FormatID (a string) to FileSaveAs. Keys are the friendly
# extensions we accept from callers.
FORMAT_ID = {
    "mpp": "MSProject.mpp",
    "mpt": "MSProject.mpt",
    "xml": "MSProject.xml",     # MSPDI
    "mspdi": "MSProject.xml",
    "csv": "MSProject.csv",
    "txt": "MSProject.txt",
    "xls": "MSProject.xls",
    "xlsx": "MSProject.xls",
    "pdf": "MSProject.pdf",
    "xpf": "MSProject.xpf",
}


# --- File save behaviour  (enum PjSaveType, used by FileCloseEx) --------------
class SaveType:
    DO_NOT_SAVE = 0
    SAVE = 1
    PROMPT = 2
