'''
Processor — «как превратить данные в таблицу»

Задача:

взять raw_data

взять week_index

посчитать

вернуть структуру для таблицы
'''
from app.processors.gpk_first_district import GPKFirstDistrictProcessor
from app.processors.ap1_first_district import AP1FirstDistrictProcessor
from app.processors.kas_appeal_regional import KASAppealRegionalProcessor
from app.processors.kas_first_regional import KASFirstRegionalProcessor
from app.processors.m_aos_first_district import MAOSFirstDistrictProcessor
from app.processors.u1_first_district import U1FirstDistrictProcessor
from app.processors.m_u1_first_district import MU1FirstDistrictProcessor
from app.processors.ap_first_district import APFirstDistrictProcessor
from app.processors.kas_first_district import KASFirstDistrictProcessor

from app.processors.gpk_first_regional import GPKFirstRegionalProcessor
from app.processors.gpk_appeal_regional import GPKAppealRegionalProcessor



class ProcessorFactory:

    @staticmethod
    def get(context):
        key = context.as_key()

        if key == ("GPK", "first", "district"):
            return GPKFirstDistrictProcessor()

        if key == ("KAS", "first", "district"):
            return KASFirstDistrictProcessor()

        if key == ("AP", "first", "district"):
            return APFirstDistrictProcessor()

        if key == ("AP1", "first", "district"):
            return AP1FirstDistrictProcessor()

        if key == ("U1", "first", "district"):
            return U1FirstDistrictProcessor()

        if key == ("M_U1", "first", "district"):
            return MU1FirstDistrictProcessor()

        if key == ("GPK", "first", "regional"):
            return GPKFirstRegionalProcessor()

        if key == ("GPK", "appeal", "regional"):
            return GPKAppealRegionalProcessor()

        if key == ("KAS", "first", "regional"):
            return KASFirstRegionalProcessor()

        if key == ("KAS", "appeal", "regional"):
            return KASAppealRegionalProcessor()

        if key == ("M_AOS", "first", "regional"):
            return MAOSFirstDistrictProcessor()


        raise ValueError("Нет процессора для контекста: %s" % (key,))
