class Day:
  def __init__(self, date):
    self.date = date
  
  def getOutputData(self):
    out = {
      "date" : self.date,
      "ymd" : "{0:%Y/%m/%d}".format(self.date),
      "day" : self.date.day,
      "weekjpn": WEEK_JPNDAYS[date.weekday()],
      "weekeng": WEEK_ENGDAYS[date.weekday()]
   }
    return out

class Event:
  def __init__(self, mark, name, type, location, description):
    self.mark = mark
    self.name = name
    self.type = type.lower() if type != None else "closed"
    self.location = location
    self.description = description.replace("_x000D", "<br>")

  def setTimeStr(time):
    ts = row[5].value.split("～")
    data["stime"] = ts[0]
    data["etime"] = ts[1]

