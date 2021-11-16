#-*- coding: utf-8 -*-

import glob,os,codecs,re,sys
import xmind
from xmind.core.workbook import WorkbookDocument
from xmind.core.topic import TopicElement
from xmind.core.mixin import TopicMixinElement

# NOTE:default 1000, 規模がでかくなるとエラーになる
sys.setrecursionlimit(200000)

input_dir = "output_files"
output_dir = 'tmp'
LF = "\n"
TAB = "  "
LIST = "- "
# NOTE: xmind書き込みの際にメソッドが別れているので区別できるように
SHEET = "# "
ROOTTOPIK = "## "
LISTREG = "(  ){0,}- "


def unescape_lf(str):
  return str.replace( '\\n' , '\n')

if __name__ == '__main__':
  args = sys.argv
  for index, arg in enumerate(args):
    if index == 1: input_dir = arg
    elif index == 2: output_dir = arg
    else: pass

  print("input_dir:" + input_dir)
  print("output_dir:" + output_dir)
  input_file_paths = glob.glob(input_dir + '/' + "*.md")

  print(input_file_paths)

  for input_file_path in input_file_paths:

    file_name = os.path.splitext(os.path.basename(input_file_path))[0]
    output_file_path = output_dir + '/' + file_name + '.xmind'
    with codecs.open(input_file_path, 'r','utf-8') as f:
      w = WorkbookDocument()

      curr_line = f.readline()
      pre_line = ''
      curr_sheet = None
      # NOTE:戻るために自分の親たちを覚えておく
      curr_topics = None
      # NOTE:戻る量を計算するため直前のネスト量を覚えておく
      curr_topic_prefix = None
      while curr_line:
        match = re.match(SHEET, curr_line) or re.match(ROOTTOPIK, curr_line) or re.match(LISTREG, curr_line)
        if match:
          prefix = match.group()
          str_data = re.sub('^' + prefix, '', curr_line).replace(LF,'').encode('utf8')
          str_data = unescape_lf(str_data)
          if prefix == SHEET:
            sheet = w.createSheet()
            sheet.setTitle(str_data)
            w.addSheet(sheet)
            curr_sheet = sheet
          elif prefix == ROOTTOPIK:
            rt = curr_sheet.getRootTopic()
            rt.setTitle(str_data)
            # NOTE: curr_topicsをrootで初期化
            curr_topics = [rt]
            curr_topic_prefix = prefix
          else:
            tp=TopicElement()
            tp.setTitle(str_data)

            now_prefix_len = len(prefix)
            curr_prefix_len = len(curr_topic_prefix)

            # NOTE: currがrootなら無条件で子供に追加
            if curr_topic_prefix == ROOTTOPIK:
              curr_topics[-1].addSubTopic(tp)
              curr_topics.append(tp)
              curr_topic_prefix = prefix
            else:
              # NOTE: 子ノードの追加ケース
              # 深さ優先探索でmdを作ってある前提なので、len(now) > len(curr) -> currの子供
              if now_prefix_len > curr_prefix_len:
                curr_topics[-1].addSubTopic(tp)
                curr_topics.append(tp)
                curr_topic_prefix = prefix
              else:
                # NOTE: len(now) <= len(curr) -> currのN個親の子供、差分からNを求める
                sub_topic_num = ( curr_prefix_len - now_prefix_len ) / len(TAB) + 1
                while sub_topic_num > 0:
                  curr_topics.pop()
                  sub_topic_num -= 1
                curr_topics[-1].addSubTopic(tp)
                curr_topics.append(tp)
                curr_topic_prefix = prefix
        else:
          # NOTE: topic内改行をエスケープしてmdに出力しているので基本でない
          print('other:' + curr_line)
        pre_line = curr_line
        curr_line = f.readline()

      # NOTE: 問答無用で作成される初期シートを削除して帳尻を合わせる
      w.removeSheet(w.getPrimarySheet())

      xmind.save(w, output_file_path)


