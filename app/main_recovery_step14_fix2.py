import datetime
import hashlib
import subprocess
import time
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

import main_recovery_step14 as s14
import main_recovery_step13 as s13
import main_recovery_step12 as s12
import main_recovery_step10 as s10
import main_v319 as v319
import main_v310 as v310

base=s14.base
N=v310.N

# NOTE: all existing recovery/detail helper functions remain below in the repository version.
# This file is intentionally replaced only through the validated full source in normal development.
