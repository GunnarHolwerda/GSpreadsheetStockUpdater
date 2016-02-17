def construct_time_variables(today):
    """
    Method stolen from Joshwa Moellenkamp

    Using a provided datetime.datetime object representing the
    current date, construct an assortment of values used by this script.
    Keyword arguments:
    today - A datetime.datetime representing the current object.
    return - (day_of_week, # Sunday, Monday, etc.
              tomorrow,    # Monday, Tuesday, etc.
              int_month,   # 1, 2, ..., 12
              str_month,   # January, February, etc.
              day,         # Day of the month
              year)        # Year
    """

    months = {
        1: "January",
        2: "February",
        3: "March",
        4: "April",
        5: "May",
        6: "June",
        7: "July",
        8: "August",
        9: "September",
        10: "October",
        11: "November",
        12: "December",
    }

    weekdays = {
        0: "Monday",
        1: "Tuesday",
        2: "Wednesday",
        3: "Thursday",
        4: "Friday",
        5: "Saturday",
        6: "Sunday",
    }
    day_of_week = weekdays.get(today.weekday())
    str_month = months.get(today.month)
    day = today.day
    if 4 <= day % 100 <= 20:
        str_day = str(day) + "th"
    else:
        str_day = str(day) + {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")

    return day_of_week, str_month, str_day

def remove_dollar_sign_and_commas(cell_value):
    new_value = cell_value[1:]
    new_value = new_value.replace(',', '')

    return new_value
