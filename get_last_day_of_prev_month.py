from datetime import datetime, timedelta

def get_last_day_of_previous_month():
    today = datetime.today()

    first_day_of_current_month = today.replace(day=1)
    last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)

    return last_day_of_previous_month.date().strftime('%d.%m.%Y')

def get_last_day_of_before_previous_month():
    today = datetime.today()

    first_day_of_current_month = today.replace(day=1)

    last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)

    first_day_of_previous_month = (last_day_of_previous_month.replace(day=1)).replace(day=1)

    last_day_of_before_previous_month = first_day_of_previous_month - timedelta(days=1)

    return last_day_of_before_previous_month.date().strftime('%d.%m.%Y')

print(get_last_day_of_before_previous_month())
