import discord
import xlrd
import random
import unicodedata
import os

TOKEN = os.environ["DISCORD_TOKEN"]
client = discord.Client()
channel = None
wb = None
q = "test"
a = "test"
q_array = ["a", "b", "c", "d"]

def zenkaku_translate(string):
    text = "！＂＃＄％＆＇（）＊＋，－．／０１２３４５６７８９：；＜＝＞？＠ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ［＼］＾＿｀>？＠ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ｛｜｝～"
    if string[0] in text:
        string = string.translate(str.maketrans({chr(0xFF01 + i): chr(0x21 + i) for i in range(94)}))
    return string

def xls_open(string):
    global wb,sheet,q,a
    try: 
        s = 'quiz/' + string + '.xlsx'
        wb = xlrd.open_workbook(s)
    except:
        q = string + '.xlsx は存在しません'
        return
    sheet = wb.sheet_by_name('Sheet1')
        
def typing_qanda():
    global wb,sheet,q,a
    col_num = sheet.nrows
    while True:
        q_num = random.randint(0,col_num-1)
        if '答えなさい' not in sheet.cell_value(q_num,0):
            break
        
    q = sheet.cell_value(q_num,0)
    a = str(sheet.cell_value(q_num,1))
    a = zenkaku_translate(a)
    if ".0" in a:
        a = a.replace(".0","")
        
def yontaku_qanda():
    global wb,sheet,q,q_array,a
    col_num = sheet.nrows / 4
    while True:
        q_num = random.randint(0,col_num-1) * 4
        if '┗' not in sheet.cell_value(q_num,0) and '分岐' not in sheet.cell_value(q_num,0):
            break
    q = sheet.cell_value(q_num,0)
    for i in range(4):
        q_array[i] = sheet.cell_value(q_num + i, 1)
        if sheet.cell_value(q_num,2) == q_array[i]:
            a = str(i+int(1))

def what_chr(string):
    name = unicodedata.name(string[0])
    if "HIRAGANA" in name:
        return "ひらがな"
    if "KATAKANA" in name:
        return "カタカナ"
    else:
        return "英数"
    

@client.event
async def on_ready():
    print('ログインしました')

@client.event
async def on_message(message):
    global wb,channel
    if channel is not None and message.channel != channel:
        return
    elif message.author.bot:
        return
    elif '&quiz ' in message.content:
        xls_open(message.content.replace('&quiz ',""))
        if q != 'ファイルオープンに失敗しました':
            channel = message.channel
        if 'tai' in message.content:
            typing_qanda()
            await message.channel.send(q + "(" + what_chr(a) + ")")
        elif 'kyu' in message.content:
            typing_qanda()
            await message.channel.send(q)
        elif 'taku' in message.content:
            yontaku_qanda()
            await message.channel.send(q)
            for i in range(4):
                await message.channel.send(str(i+int(1)) + ". " + q_array[i])
        else:
            channel = message.channel
            await message.channel.send("例外が発生しました")
    elif wb is not None and message.content.upper() == a:
        await message.channel.send("正解だ！")
        wb = None
    elif wb is not None and message.content != a:
        await message.channel.send("不正解だ～  答:" + a)
        wb = None
    elif message.content == "&quit":
        client.close()
    else:
        return
        

client.run(TOKEN)