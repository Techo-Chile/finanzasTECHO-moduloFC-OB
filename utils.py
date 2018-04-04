from datetime import date
from datetime import timedelta 
from datetime import datetime


def filterDate(rowList, date_init, date_end, columnIndex):
        ret = []
        for row in rowList:
            try:
                d = datetime.strptime(str(row[columnIndex]), '%d/%m/%Y').date()
            except ValueError:
                continue
            if date_init <= d <= date_end:
                ret.append(row)
        return ret

def filterType(rowList, orderType, columnIndex):
        ret = []
        for row in rowList:
            if orderType in row[columnIndex]:
                ret.append(row)
        return ret
    
def filterReemb(rowList):
        ret = []
        for row in rowList:
            try:
              if ('X' in row[0]) | ('x' in row[0]):
                  ret.append(row)
            except Exception:
              continue
        return ret

def ywd_to_date(year, week, weekday):
    first = date(year, 1, 1)
    first_year, _first_week, first_weekday = first.isocalendar()
    if first_year == year:
        week -= 1
    return first + timedelta(days=week*7+weekday-first_weekday)