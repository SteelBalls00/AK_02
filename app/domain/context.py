class DataContext:
    def __init__(self, specialization, instance, court_type):
        self.specialization = specialization
        self.instance = instance
        self.court_type = court_type

    @classmethod
    def from_pkl_info(cls, pkl_info):
        return cls(
            specialization=pkl_info.specialization,
            instance=pkl_info.instance,
            court_type=pkl_info.court_type,
        )

    def as_key(self):
        return (self.specialization, self.instance, self.court_type)
