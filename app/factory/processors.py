

class ProcessorFactory:

    @staticmethod
    def get(context):
        if context.specialization == "GPK":
            if context.instance == "first":
                if context.court_type == "district":
                    return GPKFirstDistrictProcessor()
                if context.court_type == "regional":
                    return GPKFirstRegionalProcessor()

            if context.instance == "appeal":
                return GPKAppealRegionalProcessor()

        raise ValueError("Processor not found")
