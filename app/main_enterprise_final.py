"""Final enterprise launcher for internal distribution.

UI: main_enterprise_v3
Weekly PPT polish: centered Signal + change-aware #0033FF update text
Progress: owned directly by main_enterprise_v3 (single visible percentage)
"""
import weekly_style_final  # noqa: F401
from main_enterprise_v3 import EnterpriseAppV3


if __name__ == '__main__':
    EnterpriseAppV3().mainloop()
