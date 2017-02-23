# python
# -*- coding: utf-8 -*-
from openpyxl import load_workbook
from jinja2 import Environment, FileSystemLoader
from pprint import pprint
from datetime import datetime
WEEK_JPNDAYS = ["月", "火", "水", "木", "金", "土", "日"]
WEEK_ENGDAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]

def main():
  # イベントリスト作成
  events = get_eventlist(u"イベント詳細一覧表.xlsx")
  caldata = get_monthevent(u"02月ふらっとイベント表.xlsx", events)

  # テンプレート展開
  baseday = caldata[0]["date"]
  vars = {
    "year": baseday.year,
    "month" : baseday.month,
    "events": caldata
  }
  env = Environment(loader=FileSystemLoader('./tmpl/', encoding='utf8'))
  tmpl= env.get_template("base.jinja2")
  html = tmpl.render(vars)
  print(html)

### イベントリストを取得
def get_eventlist(filename):
  blist = load_workbook(filename)
  slist = blist.active
  events = {}
  for row in slist.rows:
    if row[0].row != 1:
      events[row[0].value] = {"location": row[1].value, "description": row[2].value}
  return events

### イベント表をチェックする
def get_monthevent(filename, events):
  bevent = load_workbook(filename)
  sevent = bevent.active
  caldata = []
  daylist = []
  errdata = []
  date = None
  for row in sevent.rows:
    try:
      if row[0].row != 1:
        data = {}
        if date == None or date != row[1].value:
          if len(daylist) != 0: # 前日の予定をイベントリストに追加
            caldata.append({
              "date": date,
              "day" : date.day,
              "weekjpn": WEEK_JPNDAYS[date.weekday()],
              "weekeng": WEEK_ENGDAYS[date.weekday()],
              "list": daylist})
          date = row[1].value
          daylist = []
        data["name"] = row[4].value
        data["type"] = "nosection"
        data["description"] = str(events[data["name"]]["description"]).replace("_x000D_", "<br>")
        if row[5].value != "": #時刻取得(時刻がないものについてはパースしない)
          ts = row[5].value.split("～")
          data["stime"] = ts[0]
          data["etime"] = ts[1]
        daylist.append(data)

    except KeyError as e:
      # 取得エラーはあとで報告
      errdata.append({
        "date": row[1].value,
        "name": row[4].value
      })

  # エラー確認
  if len(errdata) != 0:
    print(u"イベント詳細に登録されていないイベントがあります。広報メンバーに確認してください")
    for err in errdata:
      print(err["date"].strftime(u"%m/%d") + ":" + err["name"])
    raise "処理に失敗しました"

  return caldata 

if __name__ == '__main__':
  main()