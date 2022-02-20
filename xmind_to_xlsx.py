#-*- coding: utf-8 -*-

import glob,os,codecs,sys
import xmind
import openpyxl

input_dir = "input_files"
output_dir = 'output_files'
LF = "\n"
TAB = "  "
LIST = "- "
# NOTE: xmind書き込みの際にメソッドが別れているので区別できるように
SHEET = "# "
ROOTTOPIK = "## "

def escape_crlf(str):
  return str.replace( '\n' , '\\n').replace('\r', '')

def sub_topic_recursive_processing(output_file , sub_topic, index = 0):
  output_file.write(TAB * index + LIST + escape_crlf(sub_topic.getTitle()) + LF)
  sub_topics = sub_topic.getSubTopics() or []
  for sub_topic in sub_topics:
    sub_topic_recursive_processing(output_file, sub_topic, index+1)

if __name__ == '__main__':
  # wb = openpyxl.Workbook()
  # wb.save('sample.xlsx')
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
      # f.write(SHEET + escape_crlf(sheet.getTitle()) + LF)
      rt = sheet.getRootTopic()
      newSheet.cell(1,1).value = escape_crlf(rt.getTitle())
      # f.write(ROOTTOPIK + escape_crlf(rt.getTitle()) + LF)
      # sub_topics = rt.getSubTopics() or []
      # for sub_topic in sub_topics:
      #   sub_topic_recursive_processing(f, sub_topic)

    # 上書き保存
    wb.save(output_file_path)
