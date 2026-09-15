"""Final enterprise launcher for internal distribution.

UI: main_enterprise_v3
Weekly PPT polish: centered Signal + blue automation-updated text
"""
# Importing this module installs the final weekly-PPT formatting hooks before the app runs.
import weekly_style_final  # noqa: F401
from main_enterprise_v3 import EnterpriseAppV3


if __name__ == '__main__':
    EnterpriseAppV3().mainloop()
