#-*- coding: utf-8 -*-

import glob,os,codecs,sys
import xmind
import openpyxl

input_dir = "input_files"
output_dir = 'output_files'


def sub_topic_recursive_processing(current_sheet , sub_topic, current_row, current_column = 1):
  # セル書き込み
  current_sheet.cell(current_row, current_column).value = sub_topic.getTitle()
  sub_topics = sub_topic.getSubTopics() or []

  for sub_topic in sub_topics:
    # 深さ優先探索の再起なので、呼び出すたびにcolunm増
    current_row = sub_topic_recursive_processing(current_sheet, sub_topic, current_row, current_column + 1)

  # 終了条件：末端に来たときのみrow増加
  return current_row + (not bool(sub_topics))

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

    # エクセルファイル新規作成
    wb = openpyxl.Workbook()
    # ループ処理で邪魔なので、デフォルトシートを削除
    for default_sheet in wb.worksheets:
      wb.remove(default_sheet)

    # xmindファイル読み込み
    workbook = xmind.load(input_file_path)
    sheets = workbook.getSheets()
    for sheet in sheets:
      # xmindシートごとに処理
      # エクセルシート新規作成
      newSheet = wb.create_sheet(title=sheet.getTitle())
      # 固定出力：ルートトピックは最上端
      rt = sheet.getRootTopic()
      newSheet.cell(1,1).value = rt.getTitle()

      current_row = 3 # ２行目は固定出力なので一旦飛ばす

      sub_topics = rt.getSubTopics() or []
      # サブトピックを再帰処理
      for sub_topic in sub_topics:
        current_row = sub_topic_recursive_processing(newSheet, sub_topic, current_row)
    # 上書き保存
    wb.save(output_file_path)
