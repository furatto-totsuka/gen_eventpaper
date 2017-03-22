# python
# -*- coding: utf-8 -*-
import argparse
from openpyxl import load_workbook
from jinja2 import Environment, FileSystemLoader
import re
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
def main(args):  
  # イベントリスト作成
  events = get_eventlist(args.eventlist)
  caldata = get_monthevent(args.filename, events, args.continue_is_fault)

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
  tcaldata = []
  header = []
  # 読み取りフェーズ
  for row in worksheet.rows:
    if row[0].row == 1:
      for i in range(0, 6):
        header.append(row[i].value)
    # 読み取り
    if row[0].row != 1 and row[0].value != None:
      data = {}
      mark = row[0].value[0]
      name = row[0].value[1:]
      # 日時整理
      datetimestr = re.split(u"[\(\（][日月火水木金土][\)\）]\s*", 
        str(row[2].value))
      dates = datetimestr[0].split("、")
      time = datetimestr[1]
      m = 0
      days = []
      for i in range(0, len(dates)):
        daystr = dates[i]
        d = 0
        if daystr.find("/") != -1:
          # 月がある
          dm = daystr.split("/")
          d = int(dm[1])
          m = int(dm[0])
        else:
          d = int(daystr)

        days.append(datetime(2016, m, d))

      # 概要文収集
      d = [] 
      n = [1]
      n += range(3, 7, 1)
      for i in n:
        if row[i].value != None:
          s = str(row[i].value)
          if i == 3:
              s = header[i] + ":" + s 
          d.append(s.strip())
      description = "/".join(d)
      # データ入力
      data["mark"] = mark
      data["name"] = name
      data["type"] = get_eventtype(name, events)
      data["description"] = get_eventdesc(name, events, description)
      print(data["name"] + ":" + data["description"])

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
  for day in range(1, lastday):
    dd = datetime(d.year, d.month, day)
    if dd.weekday() == 3:
      caldata.append({
        "date": dd,
        "day" : dd.day,
        "weekjpn": WEEK_JPNDAYS[dd.weekday()],
        "weekeng": WEEK_ENGDAYS[dd.weekday()],
        "text": "定休日"})
  caldata = sorted(caldata, key=lambda c: c["day"])

  return caldata 

# TODO: あとでクラス化検討
### データベース向けのイベント名称を取得する(具体的にはイベントタイトルの「第n回」などの表記を取り除き正規化する)
def get_eventname(oldname):
  import unicodedata
  # 無効な文字の除去
  oldname = re.sub(u"[\(（]第.*回[\)）]", "", oldname)
  oldname = re.sub(u"『.*』\)", "", oldname)
  # 日本語的な揺れ除去
  oldname = unicodedata.normalize("NFKC", oldname.strip())
  return oldname

### イベントタイプを取得
def get_eventtype(oldname, events):
  dbename = get_eventname(oldname)
  t = events[dbename]["type"]
  tn = t.lower() if t != None else "closed"
  return tn

### イベント概要を取得
def get_eventdesc(oldname, events, innerdescription):
  tn = ""
  if innerdescription != None:
    tn = innerdescription
  else:
    dbename = get_eventname(oldname)
    tn = events[dbename]["description"]
  return tn.replace("_x000D_", "<br>")

try:
  args = parser.parse_args()
  main(args)
except FileNotFoundError as fnfe:
  print(u"引数に指定したファイルが存在しません。")
