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


def forecast_next_year(values):
    """
    Прогноз Merch Impact.
    Небольшой рост относительно прошлого года.
    """

    result = []

    for value in values:

        if value is None:
            result.append(None)
            continue

        growth = random.uniform(0.002, 0.008)
        noise = random.uniform(-0.001, 0.001)

        new_value = value + growth + noise

        new_value = max(0, min(new_value, 1))

        result.append(round(new_value, 4))

    return result


def forecast_next_year_osa(values):
    """
    Прогноз OSA.
    Держится рядом с прошлогодними значениями,
    слегка повторяя форму исходного ряда.
    """

    result = []

    previous = None

    for i, value in enumerate(values):

        if value is None:
            result.append(None)
            continue

        # первое значение
        if previous is None:

            new_value = value + random.uniform(-0.003, 0.004)

        else:

            # повторяем изменение прошлого года
            old_delta = value - values[i - 1]

            # сохраняем примерно тот же характер изменения
            trend = old_delta * random.uniform(0.6, 1.2)

            # чаще чуть растём
            bias = random.uniform(0.0003, 0.0015)

            # небольшой случайный шум
            noise = random.uniform(-0.001, 0.001)

            new_value = previous + trend + bias + noise

            # максимум изменение за период ±0.7%
            max_step = 0.007

            if new_value > previous + max_step:
                new_value = previous + max_step

            if new_value < previous - max_step:
                new_value = previous - max_step

        # потолок 98%
        new_value = min(new_value, 0.98)

        # минимум 0
        new_value = max(new_value, 0)

        result.append(round(new_value, 4))
        previous = new_value

    return result