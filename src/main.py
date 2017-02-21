# python
# -*- coding: utf-8 -*-
from openpyxl import load_workbook
from jinja2 import Environment, FileSystemLoader
from pprint import pprint
from datetime import datetime
WEEK_JPNDAYS = ["月", "火", "水", "木", "金", "土", "日"]
WEEK_ENGDAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
# TODO: あとでメソッドに分割
# 必要なワークブック読み込み
blist = load_workbook(u"イベント詳細一覧表.xlsx")
bevent = load_workbook(u"02月ふらっとイベント表.xlsx")
slist = blist.active
sevent = bevent.active
# イベントリスト作成
events = {}
for row in slist.rows:
  if row[0].row != 1:
    events[row[0].value] = {"location": row[1].value, "description": row[2].value}
# イベントカレンダーデータ取得
caldata = []
errdata = []
for row in sevent.rows:
  try:
    if row[0].row != 1:
      data = {}
      date = row[1].value
      data["name"] = row[4].value
      data["type"] = "nosection"
      data["day"] = date.day
      data["weekjpn"] = WEEK_JPNDAYS[date.weekday()]
      data["weekeng"] = WEEK_ENGDAYS[date.weekday()]
      data["description"] = events[data["name"]]["description"]
      if row[5].value != "": #時刻取得(時刻がないものについてはパースしない)
        ts = row[5].value.split("～")
        data["stime"] = ts[0]
        data["etime"] = ts[1]
      caldata.append(data)
  except KeyError as e:
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


# テンプレート展開
vars = {
  "year": 2017,
  "month" : 4,
  "events": caldata
}
env = Environment(loader=FileSystemLoader('./tmpl/', encoding='utf8'))
tmpl= env.get_template("base.jinja2")
html = tmpl.render(vars)
print(html)
