# python
# -*- coding: utf-8 -*-
import argparse
import openpyxl
from data import EventManager, Day, EventList
from jinja2 import Environment, FileSystemLoader
from pprint import pprint
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
  events = EventManager(args.eventlist)
  caldata = get_monthevent(args.filename, events, args.continue_is_fault)
  caldata.insertHolidays()

  # テンプレート展開
  baseday = caldata.getMonthFirstDay()
  vars = {
    "year": baseday.year,
    "month" : baseday.month,
    "events": caldata.getEventListToRawData(),
    "notice": args.notice
  }
  env = Environment(loader=FileSystemLoader('./tmpl/', encoding='utf8'))
  tmpl= env.get_template(args.template + ".jinja2")
  html = tmpl.render(vars)
  print(html)
  
### イベント表をチェックする(振り分け関数)
def get_monthevent(filename, events, continue_is_fault):
  book = openpyxl.load_workbook(filename)
  sheet = book.active
  if sheet['A1'].value == "No":
    return get_monthevent_v1(sheet, events, continue_is_fault)
  else:
    return get_monthevent_v2(sheet, events, continue_is_fault)

def get_monthevent_v2(worksheet, events, continue_is_fault):
  u"""イベント表をチェックする。v2フォーマット処理"""
  caldata = EventList()
  ym = calcym(worksheet.title)
  # 一回目スキャン(リストは並び替えられていない)
  list = []
  for row in worksheet.rows:
    if row[0].row != 1:
      list.append({
        "date": int(row[0].value),
        "week": row[1].value,
        "name": row[2].value,
        "time": row[3].value,
        "content": row[4].value,
        "cost": row[5].value,
        "remark": row[6].value,
      })
  list = sorted(list, key=lambda c: c["date"])

def calcym(wstitle):
  u"""わくわくだよりのタイトルから、何年何月のわくわくだよりかチェックする"""
  import re
  go = re.match("第(\d+)号(\d+)月号", wstitle)
  year = int((int(go.group(1)) + 3) / 12) + 2014
  month = int(go.group(2))
  return (year, month)

### イベント表をチェックする
def get_monthevent_v1(worksheet, events, continue_is_fault):
  caldata = EventList()
  daylist = []
  errdata = []
  date = None
  for row in worksheet.rows:
    try:
      if row[0].row != 1:
        if date == None or date != row[1].value:
          if len(daylist) != 0: # 前日の予定をイベントリストに追加
            d = Day(date)
            d.setEvents(list(daylist))
            caldata.append(d)
          date = row[1].value
          daylist = []
        e = events.createEvent(row[3].value, row[4].value)
        if row[5].value != "": #時刻取得(時刻がないものについてはパースしない)
          e.setTimeStr(row[5].value)
        daylist.append(e)
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

  return caldata 

try:
  args = parser.parse_args()
  main(args)
except FileNotFoundError as fnfe:
  print(u"引数に指定したファイルが存在しません。")
