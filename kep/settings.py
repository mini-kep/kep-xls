from dataclasses import dataclass


def latest_date():
    return (2019, 7)


XL_MAPPER = {(2019, 7): 'краткосрочные_июль.xlsx'}
URL_MAPPER = {(2019, 7): 'ind07(2).rar'}


@dataclass
class Namer:
    year: int
    month: int

    def xl(self):
        return self.get(XL_MAPPER)

    def url(self):
        filename = self.get(URL_MAPPER)
        return f'https://gks.ru/storage/mediabank/{filename}'

    def get(self, mapper):
        key = (self.year, self.month)
        try:
            return mapper[key]
        except KeyError:
            msg = {'Supported values': mapper.keys(),
                   'got': key}
            raise ValueError(msg)


@dataclass
class Variable:
    sheet: str
    cell: str
    freq: str
    name: str


VARIABLES = [
    ('1.1. ', 'B9', 'A', 'GDP_RUB'),
    ('1.1. ', 'C9', 'Q', 'GDP_RUB'),

    ('1.1. ', 'B32', 'A', 'GDP_INDEX'),
    ('1.1. ', 'C32', 'Q', 'GDP_INDEX'),

    ('1.6. ', 'C5', 'Q', 'INVEST_RUB'),
    ('1.6. ', 'B5', 'A', 'INVEST_RUB'),

    ('1.6. ', 'B28', 'A', 'INVEST_INDEX'),
    ('1.6. ', 'C51', 'Q', 'INVEST_INDEX'),

    ('1.5. ', 'G73', 'M', 'COMM_FREIGHT'),

    ('1.12 ', 'G5', 'M', 'RETAIL_SALES'),
    ('1.12 ', 'B5', 'A', 'RETAIL_SALES'),

    ('2.1 ', 'B11', 'A', 'GOV_INC_CONS_ACCUM'),
    ('2.1 ', 'B404', 'A', 'GOV_EXP_CONS_ACCUM'),

    ('3.5 ', 'B7', 'A', 'CPI'),
    ('3.5 ', 'C7', 'Q', 'CPI'),
    ('3.5 ', 'G7', 'M', 'CPI'),

    ('4.2 ', 'B7', 'A', 'WAGE'),
    ('4.2 ', 'C7', 'Q', 'WAGE'),
    ('4.2 ', 'G7', 'M', 'WAGE'),

]


def get_variable_definitions():
    return [Variable(*r) for r in VARIABLES]
