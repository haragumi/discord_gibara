import discord
import xlrd
import random
import unicodedata

TOKEN = os.environ["DISCORD_TOKEN"]
client = discord.Client()
channel = None
wb = None
q = "test"
a = "test"

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
        await message.channel.send(q + "(" + what_chr(a) + ")")
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