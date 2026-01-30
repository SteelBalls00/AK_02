from app.constants.pkl_mapping import PKL_MAPPING


def select_pkl_for_context(pkl_files, specialization, instance):
    """
    pkl_files: список файлов в папке суда
    specialization: GPK / AP / U1
    instance: first / appeal
    """

    for pkl_name, info in PKL_MAPPING.items():
        if (
            pkl_name in pkl_files
            and info.specialization == specialization
            and info.instance == instance
        ):
            return pkl_name

    return None
