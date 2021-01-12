#Googleスプレッドシートを操作するために
import gspread
#jsonを扱うために
import json
#ServiceAccountCredentials：Googleの各サービスへアクセスできるservice変数を生成
from oauth2client.service_account import ServiceAccountCredentials

#Googleスプレッドシートの認証等の初期設定を行い、スプレッドシートのシート1を開くための関数
def authenticate():
    #2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

    #認証情報設定
    #スプレッドシート名：GO Data Searcher heroku本番用
    #ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
    credentials = ServiceAccountCredentials.from_json_keyfile_name('ダウンロードしたJSONファイル名.json', scope)

    #OAuth2の資格情報を使用してGoogle APIにログインします。
    gc = gspread.authorize(credentials)

    #共有設定したスプレッドシートキーを変数[SPREADSHEET_KEY]に格納する。
    SPREADSHEET_KEY = 'スプレッドシートキー'

    #共有設定したスプレッドシートのシート1を開く
    return gc.open_by_key(SPREADSHEET_KEY).sheet1
