u"""日本語形式の曜日文字列"""
WEEK_JPNDAYS = ["月", "火", "水", "木", "金", "土", "日"]
u"""英語形式の曜日文字列"""
WEEK_ENGDAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]

class EventList():
  u"""イベントリストを格納するリストオブジェクト"""
  def __init__(self):
    self._evtlist = list()

  def append(self, object):
    self._evtlist.append(object)

  def getEventListToRawData(self):
    u"""
      リスト内のデータを配列データに変換する
    """
    out = []
    for day in self._evtlist:
      out.append(day.getOutputData())

    return out

  def getMonthFirstDay(self):
    u"""イベントリストに格納されている月の、1日目を示す日付オブジェクトを取得する"""
    d = self._evtlist[0].getDate()
    from datetime import datetime
    return datetime(d.year, d.month, 1)

  def insertHolidays(self):
    u"""
      配列に定休日データを追加する
    """
    from datetime import datetime
    import calendar
    d = self.getMonthFirstDay()
    lastday = calendar.monthrange(d.year, d.month)[1]
    days = []
    # データの用意
    for data in self._evtlist:
      days.append(data.getDate())
    sorted(days)

    # 木曜日の追加
    for day in range(1, lastday):
      dd = datetime(d.year, d.month, day)
      if dd.weekday() == 3 and not dd.day in days:
        self.append(Day(dd, True))
    self._evtlist = sorted(self._evtlist, key=lambda c: c.getDate())
  
class EventManager:
  def __init__(self, filename):
    import openpyxl
    blist = openpyxl.load_workbook(filename)
    slist = blist.active
    self.events = {}
    for row in slist.rows:
      if row[0].row != 1 and row[0].value != None:
        n = Event.getEventName(row[0].value)
        type = row[1].value
        location = "ふらっとステーション・とつか" if row[2].value == None else row[2].value
        description = "" if row[3].value == None else row[3].value
        self.events[n] = {"location": location, 
            "type": type,
            "description": description}

  def createEvent(self, mark, name, description = None, location = None, type = None):
    """
      イベントレコードを作成する
    """
    dbename = Event.getEventName(name)
    if type == None:
      type = self.events[dbename]["type"]
      location = self.events[dbename]["location"] if location == None else location      
      description = str(self.events[dbename]["description"]) if description == None else description
    return Event(mark, name, type, location, description)

class Day:
  u"""個々の日付を表すクラス"""
  def __init__(self, date, holiday=False):
    self.date = date
    self.text = ""
    if holiday:
      self.setHoliday(True)
  
  def getDate(self):
    u"""日付値を返す"""
    return self.date
  
  def isHoliday(self):
    u"""定休日かどうかを返す"""
    return self.holiday

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
      "weekjpn": WEEK_JPNDAYS[self.date.weekday()],
      "weekeng": WEEK_ENGDAYS[self.date.weekday()],
    }
    if self.text != "":
      out["text"] = self.text
    elif self.events:
      evt = []
      for event in self.events:
        e = {
          "mark": event.mark,
          "name": event.name,
          "type": event.type,
          "location": event.location,
          "description": event.description
        }
        if 'stime' in dir(event):
          e["stime"] = event.stime
          e["etime"] = event.etime
        evt.append(e)
      out["list"] = evt
    return out

class Event:
  def __init__(self, mark, name, type, location, description):
    self.mark = mark
    self.name = name
    self.type = type.lower() if type != None else "closed"
    self.location = location
    self.description = description.replace("_x000D", "<br>")

  def setTimeStr(self, time):
    ts = time.split("～")
    import zenhan
    self.stime = zenhan.z2h(ts[0])
    self.etime = zenhan.z2h(ts[1])

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
