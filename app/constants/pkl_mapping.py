# Python 3.8 compatible

class PKLInfo:
    def __init__(self, specialization, instance, court_type):
        self.specialization = specialization    # GPK, KAS, AP, ...
        self.instance = instance                # first / appeal
        self.court_type = court_type            # district / regional


PKL_MAPPING = {
    # 1 инстанция — районные / городские суды
    "result4_with_2.pkl":   PKLInfo("GPK", "first",  "district"),
    "result4_with_2a.pkl":  PKLInfo("KAS", "first",  "district"),
    "result4_AP.pkl":       PKLInfo("AP",  "first",  "district"),
    "result4_AP1.pkl":      PKLInfo("AP1", "first",  "district"),
    "result4_U1.pkl":       PKLInfo("U1",  "first",  "district"),
    "result4_M_U1.pkl":     PKLInfo("M_U1","first",  "district"),

    # 1 инстанция — областной суд
    "result4_with_3.pkl":   PKLInfo("GPK", "first",  "regional"),

    # апелляция — областной суд
    "result4_with_33.pkl":  PKLInfo("GPK", "appeal", "regional"),
}


def get_pkl_info(pkl_name):
    if pkl_name not in PKL_MAPPING:
        raise ValueError("Неизвестный формат pkl: %s" % pkl_name)
    return PKL_MAPPING[pkl_name]
