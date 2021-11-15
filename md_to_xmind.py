#-*- coding: utf-8 -*-

import glob,os,codecs,re
import xmind

input_dir = "output_files"
os.chdir(input_dir)
input_file_names = glob.glob("*.md")
os.chdir("../")
output_dir = 'tmp'
LF = "\n"
TAB = "  "
LIST = "- "

def sub_topic_recursive_processing(output_file , sub_topic, index = 0):
  output_file.write(TAB * index + LIST + sub_topic.getTitle() + LF)
  sub_topics = sub_topic.getSubTopics() or []
  for sub_topic in sub_topics:
    sub_topic_recursive_processing(output_file, sub_topic, index+1)


print(input_file_names)

for file_name in input_file_names:
  input_file_path = input_dir + '/' + file_name
  output_file_path = output_dir + '/' + os.path.splitext(file_name)[0] + '.xmind'
  with codecs.open(input_file_path, 'r','utf-8') as f:
    w = xmind.load(output_file_path)
    curr_line = f.readline()
    pre_line = ''
    while curr_line:
      if re.match('# ', curr_line):
        print('sheet:' + re.sub('^# ', '', curr_line))
        sheet = w.createSheet()
        # TODO: 日本語が書き込めない
        # sheet.setTitle(re.sub('^# ', '', curr_line))
      elif re.match("(  ){0,}- ", curr_line):
        print('topic:' + curr_line)
      else:
        print('other:' + curr_line)
      pre_line = curr_line
      curr_line = f.readline()

    xmind.save(w)


