#-*- coding: utf-8 -*-

import glob,os,codecs,re
import xmind
from xmind.core.workbook import WorkbookDocument
from xmind.core.topic import TopicElement

input_dir = "output_files"
os.chdir(input_dir)
input_file_names = glob.glob("*.md")
os.chdir("../")
output_dir = 'tmp'
LF = "\n"
TAB = "  "
LIST = "- "
# NOTE: xmind書き込みの際にメソッドが別れているので区別できるように
SHEET = "# "
ROOTTOPIK = "## "
LISTREG = "(  ){0,}- "

print(input_file_names)

for file_name in input_file_names:
  input_file_path = input_dir + '/' + file_name
  output_file_path = output_dir + '/' + os.path.splitext(file_name)[0] + '.xmind'
  with codecs.open(input_file_path, 'r','utf-8') as f:
    w = WorkbookDocument()

    curr_line = f.readline()
    pre_line = ''
    curr_sheet = None
    curr_topic = None
    curr_topic_prefix = None
    while curr_line:
      match = re.match(SHEET, curr_line) or re.match(ROOTTOPIK, curr_line) or re.match(LISTREG, curr_line)
      if match:
        prefix = match.group()
        str_data = re.sub('^' + prefix, '', curr_line).replace(LF,'').encode('utf8')
        if prefix == SHEET:
          print('sheet:' + str_data)
          sheet = w.createSheet()
          sheet.setTitle(str_data)
          w.addSheet(sheet)
          curr_sheet = sheet
        elif prefix == ROOTTOPIK:
          print('roottopic:' + str_data)
          rt = curr_sheet.getRootTopic()
          rt.setTitle(str_data)
          curr_topic = rt
          curr_topic_prefix = prefix
        else:
          # TODO: いい感じに階層戻るやつをかんがえる
          print('topic:' + str_data)
          tp=TopicElement()
          tp.setTitle(str_data)
          curr_topic.addSubTopic(tp)
          print("index:" + str(tp.getIndex()))
          # curr_topic = tp
          # curr_topic_prefix = prefix
      else:
        # TODO: TOPIK内改行をどうするか、、、
        print('other:' + curr_line)
      pre_line = curr_line
      curr_line = f.readline()

    # NOTE: 問答無用で作成される初期シートを削除して帳尻を合わせる
    w.removeSheet(w.getPrimarySheet())

    xmind.save(w, output_file_path)


