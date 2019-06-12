import datetime, calendar

class MyLastThursday:

    def LastThInMonth(self):
        now = datetime.date.today()
        # Create a datetime.date for the last day of the given month
        year = now.year;
        month = now.month;

        daysInMonth = calendar.monthrange(year, month)[1]  # Returns (month, numberOfDaysInMonth)
        dt = datetime.date(year, month, daysInMonth)

        # Back up to the most recent Thursday
        offset = 4 - dt.isoweekday()
        if offset > 0: offset -= 7  # Back up one week if necessary
        dt += datetime.timedelta(offset)  # dt is now date of last Th in month
        print(dt.strftime('%d%b%Y').upper())
        return dt.strftime('%d%b%Y').upper()