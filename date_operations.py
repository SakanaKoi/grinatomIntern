from datetime import datetime, timedelta


# Нахождение данных о днях прошедшей недели
def get_previous_month_details():
    today = datetime.today()
    first_day_of_current_month = today.replace(day=1)
    last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)
    first_day_of_previous_month = last_day_of_previous_month.replace(day=1)

    previous_month = str(int(first_day_of_previous_month.strftime("%m")))
    last_day_pm = str(int(last_day_of_previous_month.strftime("%d")))
    first_weekday_pm = first_day_of_previous_month.weekday() + 1

    return previous_month, last_day_pm, first_weekday_pm
