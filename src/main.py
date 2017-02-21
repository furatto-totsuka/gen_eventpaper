# python
# -*- coding: utf-8 -*-
from openpyxl import load_workbook
from jinja2 import Template
from pprint import pprint
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
