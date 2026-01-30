REGIONAL_GPK_APPEAL_PKLS = {
    "result4_gpk_appeal.pkl",
    "result4_obl_gpk.pkl",
}

def detect_court_type(pkl_name):
    if pkl_name in REGIONAL_GPK_APPEAL_PKLS:
        return "regional"
    return "district"