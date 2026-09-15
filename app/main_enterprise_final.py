"""Final enterprise launcher for internal distribution.

UI: main_enterprise_v3
Weekly PPT polish: centered Signal + change-aware dark-navy update text
Progress polish: one percentage display only
"""
# Import hooks before the app starts.
import weekly_style_final  # noqa: F401
import progress_display_final  # noqa: F401
from main_enterprise_v3 import EnterpriseAppV3


if __name__ == '__main__':
    EnterpriseAppV3().mainloop()
