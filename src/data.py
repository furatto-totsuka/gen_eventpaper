WEEK_JPNDAYS = ["月", "火", "水", "木", "金", "土", "日"]
WEEK_ENGDAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]

class EventManager:
  def __init__(self, filename):
    import openpyxl
    blist = openpyxl.load_workbook(filename)
    slist = blist.active
    self.events = []
    for row in slist.rows:
      if row[0].row != 1 and row[0].value != None:
        n = get_eventname(row[0].value)
        type = row[1].value
        location = "ふらっとステーション・とつか" if row[2].value == None else row[2].value
        description = "" if row[3].value == None else row[3].value
        events[n] = {"location": location, 
            "type": type,
            "description": description}

  def createEvent(self, mark, name, description = None, location = None):
    """
      イベントレコードを作成する
    """
    dbename = Event.getEventName(data["name"])
    type = events[dbename]["type"]
    location = events[dbename]["location"] if location == None else location      
    description = str(events[dbename]["description"]) if description == None else description
    return Event(mark, name, type, location, description)
    
class Day:
  def __init__(self, date, holiday=False):
    self.date = date
    if holiday:
      self.setHoliday(true)
  
  def setEvents(self, events):
    self.events = events

  def setHoliday(self, holiday):
    self.holiday = holiday
    if self.holiday:
        self.events = []
        self.text = "定休日"

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

  @classmethod
  def getEventName(cls, oldname):
    """
      データベース向けのイベント名称を取得する
      (具体的にはイベントタイトルの「第n回」などの表記を取り除き正規化する)
    """
    import re
    import unicodedata
    # 無効な文字の除去　
    oldname = re.sub(u"[\(（]第.*回[\)）]", "", oldname)
    oldname = re.sub(u"『.*』", "", oldname)
    # 日本語的な揺れ除去
    oldname = unicodedata.normalize("NFKC", oldname.strip())
    return oldname
