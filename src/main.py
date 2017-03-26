# python
# -*- coding: utf-8 -*-
import argparse
from openpyxl import load_workbook
from jinja2 import Environment, FileSystemLoader
from pprint import pprint
from datetime import datetime
WEEK_JPNDAYS = ["月", "火", "水", "木", "金", "土", "日"]
WEEK_ENGDAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]

parser = argparse.ArgumentParser(description='ふらっとステーション・とつか ふらっとイベントだより生成ツール')
parser.add_argument('filename', 
                    type=str,
                    nargs=None,  
                    help=u'イベント表ファイルのパスを指定します')
parser.add_argument('-e', '--eventlist', 
                    type=str,
                    required=True,
                    help=u'イベント詳細定義ファイルのパスを指定します')
parser.add_argument('-f', '--continue_is_fault', 
                    default=False, 
                    action='store_true',
                    help=u'イベント詳細が見つからない項目があった場合でも、そのままリストを生成します(省略時False)')
parser.add_argument('-n', '--notice', 
                    type=str,
                    help=u'フッタに表示する通知文を指定します')
parser.add_argument('-t', '--template',
                    type=str,
                    default="doc",
                    help=u'テンプレートを指定します(省略時doc)')
def main(args):  
  # イベントリスト作成
  events = get_eventlist(args.eventlist)
  caldata = get_monthevent(args.filename, events, args.continue_is_fault)

  # テンプレート展開
  baseday = caldata[0]["date"]
  vars = {
    "year": baseday.year,
    "month" : baseday.month,
    "events": caldata,
    "notice": args.notice
  }
  env = Environment(loader=FileSystemLoader('./tmpl/', encoding='utf8'))
  tmpl= env.get_template(args.template + ".jinja2")
  html = tmpl.render(vars)
  print(html)

### イベントリストを取得
def get_eventlist(filename):
  blist = load_workbook(filename)
  slist = blist.active
  events = {}
  for row in slist.rows:
    if row[0].row != 1 and row[0].value != None:
      n = get_eventname(row[0].value)
      type = row[1].value
      location = "ふらっとステーション・とつか" if row[2].value == None else row[2].value
      description = "" if row[3].value == None else row[3].value
      events[n] = {"location": location, 
          "type": type,
          "description": description}
  return events

### イベント表をチェックする(振り分け関数)
def get_monthevent(filename, events, continue_is_fault):
  book = load_workbook(filename)
  sheet = book.active
  if sheet['A1'].value == "No":
    return get_monthevent_v1(sheet, events, continue_is_fault)
  else:
    return get_monthevent_v2(sheet, events, continue_is_fault)

### イベント表をチェックする
def get_monthevent_v2(worksheet, events, continue_is_fault):
  pass

### イベント表をチェックする
def get_monthevent_v1(worksheet, events, continue_is_fault):
  caldata = []
  daylist = []
  errdata = []
  date = None
  for row in worksheet.rows:
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
        data["mark"] = row[3].value
        data["name"] = row[4].value
        dbename = get_eventname(data["name"])
        t = events[dbename]["type"]
        data["type"] = t.lower() if t != None else "closed"
        data["description"] = str(events[dbename]["description"]).replace("_x000D_", "<br>")
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

  if len(errdata) != 0:
    # エラー確認
    import sys
    print(u"イベント詳細に登録されていないイベントがあります。広報メンバーに確認してください", file=sys.stderr)
    for err in errdata:
      print(err["date"].strftime(u"%m/%d") + ":" + err["name"], file=sys.stderr)
    if not continue_is_fault:
      raise "処理に失敗しました"

  # 木曜日を挿入する処理
  import calendar
  d = caldata[0]["date"]
  lastday = calendar.monthrange(d.year, d.month)[1]
  days = []
  for data in caldata:
    days.append(data["day"])
  sorted(days)

  for day in range(1, lastday):
    dd = datetime(d.year, d.month, day)
    if dd.weekday() == 3 and not dd.day in days:
      caldata.append({
        "date": dd,
        "day" : dd.day,
        "weekjpn": WEEK_JPNDAYS[dd.weekday()],
        "weekeng": WEEK_ENGDAYS[dd.weekday()],
        "text": "定休日"})
  caldata = sorted(caldata, key=lambda c: c["day"])

  return caldata 

### データベース向けのイベント名称を取得する(具体的にはイベントタイトルの「第n回」などの表記を取り除き正規化する)
def get_eventname(oldname):
  import re
  import unicodedata
  # 無効な文字の除去　
  oldname = re.sub(u"[\(（]第.*回[\)）]", "", oldname)
  oldname = re.sub(u"『.*』", "", oldname)
  # 日本語的な揺れ除去
  oldname = unicodedata.normalize("NFKC", oldname.strip())
  return oldname

try:
  args = parser.parse_args()
  main(args)
except FileNotFoundError as fnfe:
  print(u"引数に指定したファイルが存在しません。")
