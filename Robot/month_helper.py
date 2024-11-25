import datetime
import calendar


class MonthHelper:
    def __init__(self):
        self.today = datetime.date.today()

    def get_month_name(self, month):
        #Возвращает наименование месяца
        month_names = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
        return month_names[month-1]

    def get_previous_month(self):
        first_day = self.today.replace(day=1)
        last_month = first_day - datetime.timedelta(days=1)
        return last_month.month, last_month.year

    def get_first_and_last_date(self):
        month, year = self.get_previous_month()
        first_date = datetime.date(year, month, 1)
        last_date = datetime.date(year, month, calendar.monthrange(year, month)[1])
        return first_date, last_date

    def get_previous_month_dates(self):
        first_date, last_date = self.get_first_and_last_date()
        first_date_str = f"{first_date.day} {self.get_month_name(first_date.month)} {first_date.year}"
        last_date_str = f"{last_date.day} {self.get_month_name(last_date.month)} {last_date.year}"
        return first_date_str, last_date_str