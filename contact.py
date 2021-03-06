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

def processSentence(sen):
    sen = sen.strip()
    if sen.endswith('.') or sen.endswith('?') or sen.endswith('!'):
        sen = sen[:-1]
    return sen

def exportChunks(chunks):
    for i in range(0, len(chunks)):
        path = "./output/test/{0}.mp3".format(i)
        chunks[i].export(path, format="mp3")

def contact(folder):
    file_path = './input/{0}/{0}'.format(folder)

    # 处理excel
    loc = (file_path + ".xls")
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)

    # 检查句子是否重复
    sentence_exist = {}
    for i in range(1, sheet.nrows):
        sentence = sheet.cell_value(i, 3)
        if sentence_exist.get(sentence):
            print("{0}重复了".format(sentence))
        else:
            sentence_exist.setdefault(sentence, True)


    #处理音频
    woman_sound = AudioSegment.from_file(file_path + '_女.mp3', format="mp3")
    woman_chunks = split_on_silence(woman_sound,
        min_silence_len=1000,
        silence_thresh=-55
    )
    # exportChunks(woman_chunks)

    woman_slow_sound = AudioSegment.from_file(file_path + '_女慢.mp3', format="mp3")
    woman_slow_chunks = split_on_silence(woman_slow_sound,
        min_silence_len=1500,
        silence_thresh=-55
    )
    # exportChunks(woman_slow_chunks)

    man_sound = AudioSegment.from_file(file_path + '_男.mp3', format="mp3")
    man_chunks = split_on_silence(man_sound,
        min_silence_len=500,
        silence_thresh=-55
    )
    # exportChunks(man_chunks)


    print("excel中共{0}个句子".format(sheet.nrows-1))
    print("女生 音频中共划分出{0}个音频".format(len(woman_chunks)))
    print("女生慢 音频中共划分出{0}个音频".format(len(woman_slow_chunks)))
    print("男生 音频中共划分出{0}个音频".format(len(man_chunks)))

    # 开始输出
    count = 0
    for i in range(1, sheet.nrows):
        sentence = sheet.cell_value(sheet.nrows - i, 3)
        sentence = processSentence(sentence)

        path = "./output/{0}/{1}.mp3".format(folder, sentence)
        if os.path.exists(path):
            print("{0}已经存在".format(sentence))
            continue

        chinese_chunk = increaseDB(woman_chunks[-i*2])
        man_chunk = increaseDB(man_chunks[-i])
        woman_slow_chunk = increaseDB(woman_slow_chunks[-i])

        contacted_chunk = silence_sound*2 + chinese_chunk + silence_sound*3 + man_chunk + silence_sound*3 + woman_slow_chunk + silence_sound*2


        contacted_chunk.export(path, format="mp3")
        count = count + 1

    print("此次共生成{0}个单词音频".format(count))


contact('起床句子')