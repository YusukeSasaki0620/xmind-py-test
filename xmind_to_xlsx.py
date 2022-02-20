#-*- coding: utf-8 -*-

import glob,os,codecs,sys
import xmind
import openpyxl

input_dir = "input_files"
output_dir = 'output_files'
LF = "\n"

def escape_crlf(str):
  return str.replace( '\n' , '\\n').replace('\r', '')

def sub_topic_recursive_processing(current_sheet , sub_topic, current_row, current_column = 1):
  current_sheet.cell(current_row, current_column).value = escape_crlf(sub_topic.getTitle())
  sub_topics = sub_topic.getSubTopics() or []
  for sub_topic in sub_topics:
    sub_topic_recursive_processing(current_sheet, sub_topic, current_row, current_column+1)
    current_row+=1

if __name__ == '__main__':
  args = sys.argv
  for index, arg in enumerate(args):
    if index == 1: input_dir = arg
    elif index == 2: output_dir = arg
    else: pass

  print("input_dir:" + input_dir)
  print("output_dir:" + output_dir)
  input_file_paths = glob.glob(input_dir + '/' + "*.xmind")

  print(input_file_paths)

  for input_file_path in input_file_paths:
    file_name = os.path.splitext(os.path.basename(input_file_path))[0]
    output_file_path = output_dir + '/' + file_name + '.xlsx'

    # 新規作成
    wb = openpyxl.Workbook()
    # ループ処理で邪魔なので、デフォルトシートを削除
    for default_sheet in wb.worksheets:
      wb.remove(default_sheet)

    workbook = xmind.load(input_file_path)
    sheets = workbook.getSheets()
    for sheet in sheets:
      newSheet = wb.create_sheet(title=escape_crlf(sheet.getTitle()))
      rt = sheet.getRootTopic()
      newSheet.cell(1,1).value = escape_crlf(rt.getTitle())
      sub_topics = rt.getSubTopics() or []
      current_row = 3
      for sub_topic in sub_topics:
        sub_topic_recursive_processing(newSheet, sub_topic, current_row)
        current_row+=1

    # 上書き保存
    wb.save(output_file_path)
