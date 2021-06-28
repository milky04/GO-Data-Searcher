#discord用の
import discord
#今回主にDataFrameを使うために
import pandas as pd
#正規表現を行うために
import re
#reduce関数を使うために
from functools import reduce
#Googleスプレッドシートの認証等の初期設定を行い、スプレッドシートのシート1を開くために
import AuthenticationGSS
#処理の状態を管理するために
from ProcessPool import ProcessPool

#Botのアクセストークンを設定
TOKEN = 'botのトークン'

#接続に必要なオブジェクトを生成
client = discord.Client()

#起動時に動作する処理
@client.event
async def on_ready():
    #起動したらコマンドプロンプトにログイン通知を表示
    print('ログインしました')
    #名前の下に"'GO Data Searcher'をプレイ中"を表示
    await client.change_presence(activity=discord.Game(name='GO Data Searcher'))

#Googleスプレッドシートの認証等の初期設定を行い、スプレッドシートのシート1を開く
worksheet = AuthenticationGSS.authenticate()

#メッセージ受信時に動作する処理
@client.event
async def on_message(message):
    #メッセージ送信者がBotだった場合は無視する
    if message.author.bot:
        return

    #↓の為に全てのコマンドを変数に格納
    MSG_SS_CMD_LIST = ('/command', '/help', '/cup', '/exname', '/party', '/pand')
    #コマンドの実行が指定されたチャンネルではない場合はメンション付きのメッセージを返して無効にする
    if message.channel.name != 'チャンネル名':
        if message.content.startswith(MSG_SS_CMD_LIST):
            await message.channel.send(f"{message.author.mention} \r\n{'このチャンネルではコマンドの実行は不可です。一般チャンネルでコマンドの実行をしてください。'}")
            return

    #「/command」と発言したら内容が返る処理
    if message.content.startswith('/command'):
        #embedで埋め込みメッセージに変換
        embed_command = discord.Embed(title = 'コマンドリスト', color = discord.Color.teal())
        embed_command.add_field(name = '/party', value = '相性補完の良いポケモンを提案', inline = False)
        embed_command.add_field(name = '/pand', value = '/partyの絞り込み版', inline = False)
        embed_command.add_field(name = '/help', value = '/partyと/pandのhelpを表示', inline = False)
        embed_command.add_field(name = '/exname', value = 'アローラなどの入力名の分かりづらいポケモン名をリスト表示', inline = False)
        embed_command.add_field(name = '/cup', value = '/pandで使用可能なカップ名をリスト表示', inline = False)
        await message.channel.send(embed = embed_command)


    #「/help」と発言したら内容が返る処理
    if message.content.startswith('/help'):
        #embedで埋め込みメッセージに変換
        embed_party_help = discord.Embed(title = '/party', description = '相性補完の提案をします', color = discord.Color.teal())
        embed_party_help.add_field(name = '例：', value = '/party hl ギラティナA\r\n/party sl マリルリ エアームド マッギョ ゴンべ ミュウ (6-3対応)', inline = False)
        embed_pand_help = discord.Embed(title = '/pand', description = '/のあとに条件を入れると絞り込み出来ます(順不同)\r\nプレミアカップ→ /pc\r\nリトルカップ(S4準拠)→ /リトル\r\nメガシンカあり→ /メガ\r\n未実装を含む→ /未実装\r\nカップ指定→ /ハロウィン など(/cup参照)', color = discord.Color.teal())
        embed_pand_help.add_field(name = '例：', value = '/pand/未実装/リトル LL モノズ\r\n/pand/メガ/pc hl ピクシー ルカリオ', inline = False)
        await message.channel.send(embed = embed_party_help)
        await message.channel.send(embed = embed_pand_help)


    #「/cup」と発言したら内容が返る処理
    if message.content.startswith('/cup'):
        #embedで埋め込みメッセージに変換
        embed_cup = discord.Embed(title = 'カップ名リスト', description = '/pc\r\n/ハロウィン\r\n/ひこう\r\n/リトル\r\n/カントー\r\n/ホリデー\r\n/ラブラブ\r\n/レトロ\r\n/エレメント', color = discord.Color.teal())
        await message.channel.send(embed = embed_cup)


    #セル範囲を受け取ってDataFrameに変換する関数。cellrangeにセル範囲を指定
    def convert_ws_to_df(cellrange):
        #セル範囲を受け取ってDataFrameに変換
        ws_to_df = pd.DataFrame(worksheet.get(cellrange))
        #DataFrameのNoneを''に変換
        df_fill = ws_to_df.fillna('')
        #Indexとカラムを非表示にする
        df_result = df_fill.to_string(index = False, header = False)
        #結果を左揃えにする
        return re.sub('\s*(.+?)\n|\s*(.+?)$', '\\1\\2\n', df_result)

    #「/exname」と発言したら内容が返る処理
    if message.content.startswith('/exname'):
        #I11～Jの値のあるセルまでの値を受け取ってDataFrameに変換してembedで埋め込みメッセージに変換
        embed_exname1 = discord.Embed(title = '特殊ポケモン名リスト　1/3', description = convert_ws_to_df('I11:J61'), color = discord.Color.teal())
        embed_exname2 = discord.Embed(title = '特殊ポケモン名リスト　2/3', description = convert_ws_to_df('I62:J112'), color = discord.Color.teal())
        embed_exname3 = discord.Embed(title = '特殊ポケモン名リスト　3/3', description = convert_ws_to_df('I113:J'), color = discord.Color.teal())
        await message.channel.send(embed = embed_exname1)
        await message.channel.send(embed = embed_exname2)
        await message.channel.send(embed = embed_exname3)


    #各リーグ名を変数に格納
    LEAGUE_LIST = ('ll', 'sl', 'hl', 'ml')
    #セルB2～B8までを取得(リーグ・ポケモン名を入力・リセットするために事前に取得)
    wsc1 = worksheet.range(2, 2, 8, 2)
    #あとからセルB2,B4～B8の値を消す(リセットする)ための''を格納したタプルを事前に用意
    RESET_LEAGUE_AND_POKENAME = ('', 'oyo', '', '', '', '', '')

    #セルに値を格納する関数。cellにセル範囲、valにそれぞれのセルに入力する値を格納
    def input_cells(data):
        cell = data[0]
        val = data[1]
        #値がoyoだった場合入力しない
        if (val != 'oyo'):
            cell.value = val
        return cell

    #「/party」と発言したら発言をGSSに入出力する処理
    if message.content.startswith('/party'):
        #メッセージを受け取った時にもし処理の状態がロックされていないならロックする
        if ProcessPool.is_lock() == False:
            ProcessPool.lock()
        #メッセージを受け取った時にもし処理の状態がロックされているならメッセージを返して無効にする
        else:
            await  message.channel.send(f"{message.author.mention} \r\n{'他のユーザーのコマンドを処理中です。時間を置いて再度実行してください。'}")
            return
        #リーグが指定されていなかった場合にメンション付きのメッセージを返して無効にする
        if not any(map(message.content.lower().__contains__, LEAGUE_LIST)):
            await  message.channel.send(f"{message.author.mention} \r\n{'リーグを指定してください。'}")
            #ロックを解除する
            ProcessPool.unlock()
            return
        #発言を「 」(半角スペース)か「,」か「　」(全角スペース)か「、」で区切りで配列に変換
        search_pokemon = re.split('[ ,　、]', message.content)
        #リーグとポケモン名の間にダミーのデータ(oyo)を追加(後からセルB3を入力せずに飛ばすため.。oyoのセル座標はB3)
        search_pokemon.insert(2, 'oyo')
        #配列の要素数を数えて格納
        search_pokemon_length = len(search_pokemon)
        #要素数が8を超えたら無効にする(ポケモン名が6匹以上入力されたり区切り文字が多いなどメッセージの記述が正しくない場合に無効にしてメンション付きのメッセージを返す。本来なら要素数は7だけどダミーのデータ(oyo)を追加しているのでそれを含めて8。)
        if search_pokemon_length > 8:
            await  message.channel.send(f"{message.author.mention} \r\n{'ポケモンが6匹以上入力されているかコマンドの記述が正しくありません。/helpを参考に正しく記述してください。'}")
            #ロックを解除する
            ProcessPool.unlock()
            return
        try:
            #区切られた単語をセルB3を飛ばしてB2とB4～最大B8までに入力
            #リーグ名をセルB2に、ポケモン名1～最大5までをセルB4～最大B8までに格納(入力されたメッセージによって可変)
            wsc1 = list(map(input_cells, zip(wsc1, search_pokemon[1:])))
            #セルB2とB4～最大B8までの値を更新
            worksheet.update_cells(wsc1)

            #B2～D8までの検索値を受け取ってDataFrameに変換してembedで埋め込みメッセージに変換
            embed_party_searchvalue = discord.Embed(title = '検索値', description = convert_ws_to_df('B2:D8'), color = discord.Color.teal())
            #メンション付きの値をbotが返す
            await message.channel.send(f"{message.author.mention}", embed = embed_party_searchvalue)

            #B11～E61までの値(50件)を受け取ってDataFrameに変換してembedで埋め込みメッセージに変換
            embed_party1 = discord.Embed(title = '検索結果', description = convert_ws_to_df('B11:E61'), color = discord.Color.teal())
            #メンション付きの値をbotが返す
            await message.channel.send(f"{message.author.mention}", embed = embed_party1)

            #検索結果が50件よりも多かった場合の処理
            try:
                #B62～E111までの値(字数制限で現状暫定E111まで(前のと合わせて最大100件))を受け取ってDataFrameに変換してembedで埋め込みメッセージに変換
                embed_party2 = discord.Embed(title = '検索結果2', description = convert_ws_to_df('B62:E111'), color = discord.Color.teal())
                #メンション付きの値をbotが返す
                await message.channel.send(f"{message.author.mention}", embed = embed_party2)
            #検索結果が50件よりも少なかった場合はこの処理をパスする
            except KeyError:
                pass
        finally:
            #セルB2,B4～B8の値を消す(リセットのため)
            wsc1 = list(map(input_cells, zip(wsc1, RESET_LEAGUE_AND_POKENAME)))
            #セルB2,B4～最大B8までの値を更新
            worksheet.update_cells(wsc1)
            #ロックを解除する
            ProcessPool.unlock()


    #各カップ名を変数に格納
    CUP_LIST = ('/ハロウィン', '/ひこう', '/リトル', '/カントー', '/ホリデー', '/ラブラブ', '/レトロ', '/エレメント')
    #セルG2～G47に値を入力するためのリストを用意(このリストがセルG2～G47と同義)
    conditions_cells = [''] * 46
    #セルG2～G47を取得
    wsc_conditions = worksheet.range(2, 7, 47, 7)
    #「/pand」と発言したら発言をGSSに入出力する処理
    if message.content.startswith('/pand'):
        #メッセージを受け取った時にもし処理の状態がロックされていないならロックする
        if ProcessPool.is_lock() == False:
            ProcessPool.lock()
        #メッセージを受け取った時にもし処理の状態がロックされているならメッセージを返して無効にする
        else:
            await  message.channel.send(f"{message.author.mention} \r\n{'他のユーザーのコマンドを処理中です。時間を置いて再度実行してください。'}")
            return
        #リーグが指定されていなかった場合にメンション付きのメッセージを返して無効にする
        if not any(map(message.content.lower().__contains__, LEAGUE_LIST)):
            await  message.channel.send(f"{message.author.mention} \r\n{'リーグを指定してください。'}")
            #ロックを解除する
            ProcessPool.unlock()
            return
        #メッセージにカップ名が2つ以上指定されていた場合にメッセージを返して無効にする
        #matchesは合計値、keywordは単語(1個ずつ)を表す
        #入力されたメッセージに対して、CUP_LIST内のそれぞれの単語が含まれている個数を合計する
        #matches + message.content.count(keyword)で上記の処理を実行する
        #メッセージに含まれている単語(keyword)をcountで1個ずつ数える。数えた数をmatchesに加算していく(matchesがその合計となる)。合計値が1よりも大きくなったらメッセージを返して無効にする
        if reduce(lambda matches, keyword: matches + message.content.count(keyword), CUP_LIST, 0) > 1:
            await  message.channel.send(f"{message.author.mention} \r\n{'カップ名は複数指定できません。'}")
            #ロックを解除する
            ProcessPool.unlock()
            return
        #発言を「 」(半角スペース)か「,」か「　」(全角スペース)か「、」で区切りで配列に変換
        search_pokemon = re.split('[ ,　、]', message.content)
        #リーグとポケモン名の間にダミーのデータ(oyo)を追加(後からセルB3を入力せずに飛ばすため)
        search_pokemon.insert(2, 'oyo')
        #配列の要素数を数えて格納
        search_pokemon_length = len(search_pokemon)
        #要素数が8を超えたら無効にする(ポケモン名が6匹以上入力されたり区切り文字が多いなどメッセージの記述が正しくない場合に無効にしてメンション付きのメッセージを返す。本来なら要素数は7だけどダミーのデータ(oyo)を追加しているのでそれを含めて8。)
        if search_pokemon_length > 8:
            await  message.channel.send(f"{message.author.mention} \r\n{'ポケモンが6匹以上入力されているかコマンドの記述が正しくありません。/helpを参考に正しく記述してください。'}")
            #ロックを解除する
            ProcessPool.unlock()
            return
        try:
            #pcが指定された場合にセルG2に'pc'を入力する処理
            if all(map(message.content.lower().__contains__, '/pc')):
                conditions_cells[0] = 'pc'
            #メガが指定された場合にセルG3に'メガ'を入力する処理
            if '/メガ' in message.content:
                conditions_cells[1] = 'メガ'
            #未実装が指定された場合にセルG4に'未実装'を入力する処理
            if '/未実装' in message.content:
                conditions_cells[2] = '未実装'
            #ハロウィンカップが指定された場合にセルG5～G9に'どく','むし','あく','ゴースト','フェアリー'を入力する処理
            if '/ハロウィン' in message.content:
                conditions_cells[3] = 'どく'
                conditions_cells[4] = 'むし'
                conditions_cells[5] = 'あく'
                conditions_cells[6] = 'ゴースト'
                conditions_cells[7] = 'フェアリー'
            #ひこうカップが指定された場合にセルG5に'ひこう'を入力する処理
            if '/ひこう' in message.content:
                conditions_cells[3] = 'ひこう'
            #リトルカップが指定された場合にセルG22に'リトル'を入力する処理
            if '/リトル' in message.content:
                conditions_cells[20] = 'リトル'
            #カントーカップが指定された場合にセルG23に'カントー'を入力する処理
            if '/カントー' in message.content:
                conditions_cells[21] = 'カントー'
            #ホリデーカップが指定された場合にセルG5～G9に'ノーマル','くさ','でんき','こおり','ひこう','ゴースト'を入力する処理
            if '/ホリデー' in message.content:
                conditions_cells[3] = 'ノーマル'
                conditions_cells[4] = 'くさ'
                conditions_cells[5] = 'でんき'
                conditions_cells[6] = 'こおり'
                conditions_cells[7] = 'ひこう'
                conditions_cells[8] = 'ゴースト'
            #ラブラブカップが指定された場合にセルG30に'ラブラブ'を入力する処理
            if '/ラブラブ' in message.content:
                conditions_cells[28] = 'ラブラブ'
            #レトロカップが指定された場合にセルG31～G33に'あく','はがね','フェアリー'を入力する処理
            if '/レトロ' in message.content:
                conditions_cells[29] = 'あく'
                conditions_cells[30] = 'はがね'
                conditions_cells[31] = 'フェアリー'
            #エレメントカップが指定された場合にセルG5～G7に'ほのお','みず','くさ'、セルG22に'リトル'を入力する処理
            if '/エレメント' in message.content:
                conditions_cells[3] = 'ほのお'
                conditions_cells[4] = 'みず'
                conditions_cells[5] = 'くさ'
                conditions_cells[20] = 'リトル'

            #区切られた単語をセルB3を飛ばしてB2とB4～最大B8までに入力
            #リーグ名をセルB2に、ポケモン名1～最大5までをセルB4～最大B8までに格納(入力されたメッセージによって可変)
            wsc1 = list(map(input_cells, zip(wsc1, search_pokemon[1:])))
            #セルG2～G47に指定された条件を格納して入力
            wsc_conditions = list(map(input_cells, zip(wsc_conditions, conditions_cells)))
            #セルB2とB4～最大B8までとG2～G47の値を更新
            worksheet.update_cells(wsc1 + wsc_conditions)

            #B2～D8までの検索値を受け取ってDataFrameに変換してembedで埋め込みメッセージに変換
            embed_party_searchvalue = discord.Embed(title = '検索値', description = convert_ws_to_df('B2:D8'), color = discord.Color.teal())
            #条件指定された場合の処理
            try:
                #G2～G47までの値を受け取ってDataFrameに変換してembed_party_searchvalueにフィールドとして追加
                embed_party_searchvalue.add_field(name = '条件', value = convert_ws_to_df('G2:G47'), inline = False)
            #条件指定されなかった場合はこの処理をパスする
            except KeyError:
                pass
            #メンション付きの値をbotが返す
            await message.channel.send(f"{message.author.mention}", embed = embed_party_searchvalue)

            #B11～E61までの値(50件)を受け取ってDataFrameに変換してembedで埋め込みメッセージに変換
            embed_party1 = discord.Embed(title = '検索結果', description = convert_ws_to_df('B11:E61'), color = discord.Color.teal())
            #メンション付きの値をbotが返す
            await message.channel.send(f"{message.author.mention}", embed = embed_party1)

            #検索結果が50件よりも多かった場合の処理
            try:
                #B62～E111までの値(字数制限で現状暫定E111まで(前のと合わせて最大100件))を受け取ってDataFrameに変換してembedで埋め込みメッセージに変換
                embed_party2 = discord.Embed(title = '検索結果2', description = convert_ws_to_df('B62:E111'), color = discord.Color.teal())
                #メンション付きの値をbotが返す
                await message.channel.send(f"{message.author.mention}", embed = embed_party2)
            #検索結果が50件よりも少なかった場合はこの処理をパスする
            except KeyError:
                pass
        finally:
            #セルB2,B4～B8,G2～G47の値を消す(リセットのため)
            wsc1 = list(map(input_cells, zip(wsc1, RESET_LEAGUE_AND_POKENAME)))
            #セルG2～G47の値に''を格納
            conditions_cells = [''] * 46
            wsc_conditions = list(map(input_cells, zip(wsc_conditions, conditions_cells)))

            #セルB2,B4～B8,G2～G47に入力された値を消す＆更新
            worksheet.update_cells(wsc1 + wsc_conditions)
            #ロックを解除する
            ProcessPool.unlock()


#Botの起動とDiscordサーバーへの接続
client.run(TOKEN)
