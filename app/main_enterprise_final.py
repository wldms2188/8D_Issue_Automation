import weekly_style_final  # noqa: F401
import excel_change_style_final  # noqa: F401
from main_enterprise_v3 import EnterpriseAppV3
import team_autocomplete_final  # noqa: F401
import project_autocomplete_final  # noqa: F401
import project_weekly_match_final  # noqa: F401
import problem_summary_final  # noqa: F401
import english_8d_final  # noqa: F401
import attachment_dedupe_final  # noqa: F401
import final_output_polish  # noqa: F401
import origin_v1_refined_patch  # noqa: F401
import output_variant_final  # noqa: F401

# Keep the user-confirmed stable baseline intact and apply only the five
# weekly-meeting fixes after every existing patch has loaded.
import user_stable_weekly_fixes  # noqa: F401

if __name__ == '__main__':
    EnterpriseAppV3().mainloop()
