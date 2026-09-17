"""Install the final issue-origin recommender without changing the validated core."""
import main_enterprise as ent
import origin_classifier_final as clf

# EnterpriseOriginDialog resolves this function at dialog construction time.
ent.legacy.step8._recommend_origin = clf.recommend_origin
