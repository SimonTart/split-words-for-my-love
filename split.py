from pydub import AudioSegment
from pydub.silence import split_on_silence
import xlrd
import os

silence_sound = AudioSegment.from_file('./silence.mp3', format="mp3")

def addSilence(chunk):
    return silence_sound + chunk + silence_sound


def increaseDB(chunk):
    # 0.1 = 290
    # base = 28000
    add = (28000 - chunk.max)/290 * 0.1
    return chunk + add

def split_words(folder):
    file_path = './input/{0}/{0}'.format(folder)
    sound = AudioSegment.from_file(file_path + '.mp3', format="mp3")
    chunks = split_on_silence(sound,
        min_silence_len=1000,
        silence_thresh=-50
    )

    loc = (file_path + ".xls")
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)

    print("音频中共划分出{0}个单音频".format(len(chunks)))
    print("excel中共{0}个单词".format(sheet.nrows-1))

    # 检查单词是否重复
    word_exist = {}
    for i in range(1, sheet.nrows):
        word = sheet.cell_value(i, 2)
        if word_exist.get(word):
            print("{0}重复了".format(word))
        else:
            word_exist.setdefault(word, True)

    # 开始输出
    count = 0
    for i in range(1, sheet.nrows):
        word = sheet.cell_value(sheet.nrows - i, 2).strip()
        path = "./output/{0}/{1}.mp3".format(folder, word)
        if os.path.exists(path):
            print("{0}已经存在".format(word))
        else:
            chunk = addSilence(increaseDB(chunks[-i]))
            chunk.export(path, format="mp3")
            # print(chunk.max)
            count = count + 1

    print("此次共生成{0}个单词音频".format(count))


split_words('回到家')
# split_words('起床')
# split_words('常用句型三')


