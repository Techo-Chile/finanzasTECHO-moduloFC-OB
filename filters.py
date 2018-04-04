from datetime import datetime  
from datetime import timedelta 
from datetime import datetime


def filterDate(rowList, dt, columnIndex):
        ret = []
        for row in rowList:
            try:
                d = datetime.strptime(str(row[columnIndex]), '%m/%d/%Y')
            except ValueError:
                continue
            if (dt == d) | (dt - timedelta(days=7) < d):
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