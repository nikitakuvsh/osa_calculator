import random

random.seed(42)

def limit_growth(previous_year, predicted):
    """
    previous_year - значение прошлого года (0..1)
    predicted     - то, что посчитал текущий алгоритм

    Возвращает ограниченное значение.
    """

    # Ниже прошлого года уходить нельзя
    predicted = max(predicted, previous_year)

    # До потолка осталось
    distance = 0.98 - previous_year

    if distance <= 0:
        return round(min(previous_year, 0.98), 4)

    # Максимальный допустимый прирост
    if previous_year >= 0.95:
        max_growth = 0.001      # +0.1%
    elif previous_year >= 0.92:
        max_growth = 0.002      # +0.2%
    elif previous_year >= 0.90:
        max_growth = 0.003      # +0.3%
    elif previous_year >= 0.85:
        max_growth = 0.005      # +0.5%
    elif previous_year >= 0.80:
        max_growth = 0.010      # +1%
    elif previous_year >= 0.70:
        max_growth = 0.015      # +1.5%
    else:
        max_growth = 0.020      # +2%

    predicted = min(predicted, previous_year + max_growth)

    return round(min(predicted, 0.98), 4)


def forecast_missing_periods(values):
    """
    Достраивает отсутствующие периоды текущего года.

    Например:
    P1-P8 есть,
    P9-P13 None

    Первые значения сохраняются.
    Заполняются только None.
    """

    result = values.copy()

    # ищем последнее известное значение
    last_value = None

    for value in result:
        if value is not None:
            last_value = value


    if last_value is None:
        return result


    for i in range(len(result)):

        if result[i] is None:

            growth = random.uniform(
                0.002,
                0.006
            )

            noise = random.uniform(
                -0.001,
                0.001
            )


            new_value = last_value + growth + noise


            new_value = limit_growth(last_value, new_value)


            result[i] = new_value

            last_value = new_value


    # если меньше 13 периодов
    while len(result) < 13:

        growth = random.uniform(
            0.002,
            0.006
        )

        noise = random.uniform(
            -0.001,
            0.001
        )


        new_value = last_value + growth + noise

        new_value = max(
            0,
            min(new_value, 1)
        )


        result.append(
            round(new_value,4)
        )

        last_value = new_value


    return result



def forecast_missing_periods_osa(values):
    """
    Достройка OSA текущего года.

    Сохраняет P1-P8.
    Заполняет только пропуски.
    """

    result = values.copy()


    previous = None

    for value in result:
        if value is not None:
            previous = value


    if previous is None:
        return result



    for i in range(len(result)):

        if result[i] is None:


            change = random.uniform(
                -0.002,
                0.003
            )


            noise = random.uniform(
                -0.001,
                0.001
            )


            new_value = previous + change + noise


            new_value = limit_growth(previous, new_value)


            result[i] = new_value

            previous = new_value



    while len(result) < 13:

        change = random.uniform(
            -0.002,
            0.003
        )

        new_value = previous + change

        new_value = max(
            0,
            min(new_value,0.98)
        )


        result.append(
            round(new_value,4)
        )

        previous = new_value


    return result



def forecast_next_year(values):
    """
    Прогноз следующего года (например 2027).

    2027 P1 строится от 2026 P1,
    2027 P2 от 2026 P2 и т.д.

    Сохраняет сезонность.
    """

    result = []


    for value in values:


        if value is None:

            result.append(None)

            continue



        growth = random.uniform(
            0.04,
            0.08
        )


        noise = random.uniform(
            -0.005,
            0.005
        )


        new_value = value * (1 + growth) + noise
        new_value = limit_growth(value, new_value)


        result.append(
            round(new_value,4)
        )


    return result



def forecast_next_year_osa(values):
    """
    Прогноз OSA следующего года.

    Сохраняет форму 2026 года.
    """

    result = []

    previous = None


    for i,value in enumerate(values):


        if value is None:

            result.append(None)

            continue



        if previous is None:


            new_value = value + random.uniform(
                -0.002,
                0.004
            )


        else:


            old_delta = value - values[i-1]


            trend = old_delta * random.uniform(
                0.6,
                1.2
            )


            bias = random.uniform(
                0.0005,
                0.002
            )


            noise = random.uniform(
                -0.001,
                0.001
            )


            new_value = (
                previous
                +
                trend
                +
                bias
                +
                noise
            )



        new_value = limit_growth(value, new_value)


        result.append(new_value)

        previous = new_value



    return result