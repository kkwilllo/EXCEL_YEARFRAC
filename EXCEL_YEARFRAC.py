import datetime

class EXCEL_YEARFRAC():
    def __init__(self):
        pass

    def appears_le_year(self, date1:datetime.date, date2:datetime.date):
        # Returns True if date1 and date2 "appear" to be 1 year or less apart.
        # This compares the values of year, month, and day directly to each other.
        # Requires date1 <= date2; returns boolean. Used by basis 1.
        if date1.year == date2.year:
            return True
        if (((date1.year + 1) == date2.year) and
            ((date1.month > date2.month) or
            ((date1.month == date2.month) and (date1.day >= date2.day)))):
            return True
        return False

    # Check if year is a leap year
    def is_leap_year(self, year:int):
        return year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)

    def feb29_between(self, date1:datetime.date, date2:datetime.date):
        # Check each year in the range
        for year in range(date1.year, date2.year + 1):
            if self.is_leap_year(year):
                leap_day = datetime.date(year, 2, 29)
                if date1 <= leap_day <= date2:
                    return True
        return False

    def diffdays(self, date1:datetime.date, date2:datetime.date):
        return (date2-date1).total_seconds()  / 86400

    def basis1(self, date1:Union[datetime.date,None], date2:Union[datetime.date,None]):
        if date1 and date2:
            if type(date1) == datetime.datetime:
                date1 = date1.date()
            elif type(date2) == datetime.datetime:
                date2 = date2.date()
        else:
            return None
        # Swap so date1 <= date2 in all cases:
        if date1 > date2:
            date1, date2 = date2, date1
        if date1 == date2:
            return 0.0
        if self.appears_le_year(date1, date2):
            print('leap_year')
            if (date1.year == date2.year and self.is_leap_year(date1.year)):
                year_length = 366.
            elif (self.feb29_between(date1, date2) or
                (date2.month == 2 and date2.day == 29)): # fixed, 2008-04-18
                print('leap year feb29')
                year_length = 366.
            else:
                print('leap year else')
                year_length = 365.
            return self.diffdays(date1, date2) / year_length
        else:
            num_years = (date2.year - date1.year) + 1
            days_in_years = self.diffdays(datetime.date(date1.year, 1, 1), datetime.date(date2.year+1, 1, 1))
            average_year_length = days_in_years / num_years
            return self.diffdays(date1, date2) / average_year_length

'''
# Sample Run
### YEARFRAC() ###
import datetime
listDate = [
    
    {'start_date':'2024-01-01', 'end_date':'2024-12-31'},
    {'start_date':'2024-01-01', 'end_date':'2025-01-02'},
    {'start_date':'2024-01-01', 'end_date':'2024-02-29'},
    {'start_date':'2024-01-01', 'end_date':'2024-03-01'},
    {'start_date':'2023-01-01', 'end_date':'2023-03-01'},
    {'start_date':'2024-02-29', 'end_date':'2025-02-28'},
    {'start_date':'2024-01-01', 'end_date':'2028-12-31'},
    {'start_date':'2024-03-01', 'end_date':'2025-03-01'},
    {'start_date':'2024-02-29', 'end_date':'2025-03-01'},
    
    {'start_date':'2024-02-29', 'end_date':'2028-02-28'},
    {'start_date':'2024-02-29', 'end_date':'2028-02-29'},
    {'start_date':'2024-03-01', 'end_date':'2028-03-01'},
]

for _dict in listDate:
    print(_dict)
    start_date = datetime.datetime.strptime(_dict['start_date'],'%Y-%m-%d')
    end_date = datetime.datetime.strptime(_dict['end_date'],'%Y-%m-%d')
    print(EXCEL_YEARFRAC().basis1(start_date, end_date))
'''
