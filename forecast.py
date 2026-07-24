import random

random.seed(42)


def forecast_next_year(values):

    result = []

    for value in values:
        if value is None:
            result.append(None)
            continue

        change = random.uniform(
            0.002,
            0.008
        )

        new_value = value + change

        result.append(
            round(new_value, 4)
        )

    return result

def forecast_next_year_osa(values):

    if not values:
        return []


    # убираем None
    clean = [
        v for v in values
        if v is not None
    ]


    if len(clean) < 3:
        return values


    # базовый уровень - среднее последних 4 периодов
    base = sum(clean[-4:]) / 4


    result = []


    # небольшие колебания
    changes = [
        -0.008,
        -0.005,
        -0.01,
        -0.006,
        0.002,
        -0.003,
        0.004,
        0.005,
        -0.002,
        0,
        0,
        0,
        0
    ]


    for i in range(13):

        value = base + changes[i]


        # ограничение от 0 до 100%
        value = max(
            0,
            min(
                value,
                1
            )
        )


        result.append(
            round(value, 4)
        )


    return result