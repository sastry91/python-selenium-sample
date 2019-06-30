import datetime, calendar

class MyLastThursday:

    def LastThInMonth(self):
        now = datetime.date.today()
        # Create a datetime.date for the last day of the given month
        year = now.year;
        month = now.month;
        computedDate = self.computeDate(year, month)
        if computedDate > now :
            print(computedDate)
            return computedDate.strftime('%d%b%Y').upper()
        else:
            computedDate = self.computeDate(year, month+1)
            print(computedDate)
            return computedDate.strftime('%d%b%Y').upper()


    def computeDate(self,year,month):
        daysInMonth = calendar.monthrange(year, month)[1]  # Returns (month, numberOfDaysInMonth)
        dt = datetime.date(year, month, daysInMonth)

        # Back up to the most recent Thursday
        offset = 4 - dt.isoweekday()
        if offset > 0: offset -= 7  # Back up one week if necessary
        dt += datetime.timedelta(offset)  # dt is now date of last Th in month
        return dt;
