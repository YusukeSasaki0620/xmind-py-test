#-*- coding: utf-8 -*-

import glob,os,codecs,sys
import xmind

input_dir = "input_files"
output_dir = 'output_files'
LF = "\n"
TAB = "  "
LIST = "- "
# NOTE: xmind書き込みの際にメソッドが別れているので区別できるように
SHEET = "# "
ROOTTOPIK = "## "

def sub_topic_recursive_processing(output_file , sub_topic, index = 0):
  output_file.write(TAB * index + LIST + sub_topic.getTitle() + LF)
  sub_topics = sub_topic.getSubTopics() or []
  for sub_topic in sub_topics:
    sub_topic_recursive_processing(output_file, sub_topic, index+1)

if __name__ == '__main__':
  args = sys.argv
  for index, arg in enumerate(args):
    if index == 1: input_dir = arg
    elif index == 2: output_dir = arg
    else: pass

  print("input_dir:" + input_dir)
  print("output_dir:" + output_dir)
  os.chdir(input_dir)
  input_file_names = glob.glob("*.xmind")
  os.chdir("../")

  print(input_file_names)

  for file_name in input_file_names:
    input_file_path = input_dir + '/' + file_name
    output_file_path = output_dir + '/' + os.path.splitext(file_name)[0] + '.md'
    with codecs.open(output_file_path, 'w', 'utf-8') as f:
      workbook = xmind.load(input_file_path)
      sheets = workbook.getSheets()
      for sheet in sheets:
        f.write(SHEET + sheet.getTitle() + LF)
        rt = sheet.getRootTopic()
        f.write(ROOTTOPIK + rt.getTitle() + LF)
        sub_topics = rt.getSubTopics() or []
        for sub_topic in sub_topics:
          sub_topic_recursive_processing(f, sub_topic)

