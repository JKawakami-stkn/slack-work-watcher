from slackbot.bot import default_reply
from slackbot.bot import listen_to
from slackbot.bot import respond_to
from slacker import Slacker

from slackbot_settings import API_TOKEN

import re
import datetime


from libs import functions

# Bot専用のチャンネル
CHANNEL = "GLYGWPEG7"




###########################################################################
### 名前      ：show_command_list
### 説明      ：全ての投稿を監視し、パターンに当てはまらなかった場合に、
###                 コマンドの一覧を表示する
### 引数      ：(message) 情報
### 戻り値    ：なし
### 参照関数  ：
###########################################################################

@default_reply()
def show_command_list(message):
    message.send(" :kishir:- コマンド一覧 -:kishir: \n"
        + "・ユーザー登録".ljust(9, "　")  + ":    [登録](str)表示名\n"
        + "・作業開始".ljust(9, "　")      + ":    [開始](str)タスク名,(int)予定日数\n"
        + "・作業確認".ljust(9, "　")      + ":    [確認]\n"
        + "・作業終了".ljust(9, "　")      + ":    [終了](int)タスクID'\n"
        + "・WBS出力".ljust(9, "　")      + ":    [出力]")






###########################################################################
### 名前      ：regist_user
### 説明      ：Botへのメンションを監視し、登録コマンドが入力されたら、
###                   DBにユーザーの情報を登録する。
### 引数      ：(message) 情報
### 戻り値    ：なし
### 参照関数  ：
###########################################################################

@respond_to(r'^\[登録\].+$')
def regist_user(message):

    # メッセージを取り出す
    text = message.body['text']

    # ユーザーIDを取り出す
    user_id = message.body['user']

    # メッセージテキストから表示ユーザー名を取得
    user_name = re.sub(r'\[.*\]', "", text).strip() # タグと空白を削除

    # DBにユーザー情報登録
    if(functions.setUser(user_id, user_name)):
        # 登録成功メッセージ
        message.reply(user_name + "(" + user_id + ")を作業者リストに登録しました。")
    else:
        # 登録失敗メッセージ
        message.reply("ID( " + user_id +" )は既に登録されています。")





###########################################################################
### 名前      ：start_task
### 説明      ：Botへのメンションを監視し、開始コマンドが入力されたら、
###                   DBにタスクの情報を登録し、遅れなどの情報を表示する
### 引数      ：(message) 投稿されたメッセージに関する情報
### 戻り値    ：なし
### 参照関数  ：
###########################################################################

@respond_to(r'^\[開始\].+,\s*\d+$')
def start_task(message):

    # メッセージを取り出す
    text = message.body['text']

    # メッセージテキストからタスク名と作業日数を取得
    data = re.sub(r'\[.*\]', "", text).strip().split(",") # タグと空白を削除

    # データを変数に格納
    task_name = data[0] # タスク名
    work_days = data[1] # 作業日数

    # ユーザーidを取り出す
    user_id = message.body["user"]

    # 作業者が登録されているか確認する
    if(functions.getUserData( user_id ) is not None):

        # 作業者名を取得
        user_name = functions.getUserData(user_id)

        # 前回タスクの終了予定日を取得
        last_finish_schedule = functions.get_last_finish_schedule(user_id)

        # 開始予定日を設定する。初めての登録の場合今日、そうでない場合前回タスク終了予定日の次の日をとする。(休日も働け)
        if(last_finish_schedule == datetime.date.today()):
            start_schedule = datetime.date.today()
        else:
            start_schedule = last_finish_schedule + datetime.timedelta(days=int(1))

        # 開始日を今日に設定
        start = datetime.date.today()

        # 開始予定日と工数から終了予定日を計算
        finish_schedule =  start_schedule + datetime.timedelta(days=int(work_days))

        # タスクを登録する
        if(functions.setTask(task_name, user_id, start, start_schedule, finish_schedule)):
            message.reply("[" + task_name + "]を開始します。開始:" + str(start) + "  開始予定:" + str(start_schedule) + "  終了予定:" + str(finish_schedule))

            # メッセージの送信先をBot用のチャンネルに変更しメッセージを送信
            message.body['channel'] = CHANNEL
            message.send(user_name +"が[" + task_name + "]を開始しました。開始:" + str(start) + "  開始予定:" + str(start_schedule) + "  終了予定:" + str(finish_schedule))
        else:
            message.reply("[" + task_name + "]の登録に失敗しました。")

    else:
        message.reply("あなた(" + user_id + ")は作業者ではありません。")





###########################################################################
### 名前      ：finish_task
### 説明      ：Botへのメンションを監視し、終了コマンドが入力されたら、
###                 DBに終了日時を登録し、遅れなどの情報を表示する
### 引数      ：(message) 投稿されたメッセージに関する情報
### 戻り値    ：なし
### 参照関数  ：
###########################################################################

@respond_to(r'^\[終了\][0-9]+$')
def finish_task(message):

    # メッセージを取り出す
    text = message.body['text']

    # メッセージテキストからタスクIDまたはタスク名を取得
    task_id = re.sub(r'\[.*\]', "", text).strip() # タグと空白を削除

    # メッセージからユーザーID取得
    user_id = message.body["user"]

    # タスクIDからタスクの全情報を取得
    task_data = functions.getTaskData(task_id)


    # 値が取れているか
    if(len( task_data) != 0):

        # 必要なデータを抽出
        task_name = task_data[0][1]
        staff_id = task_data[0][2]
        finish = task_data[0][4]
        finish_schedule = task_data[0][6]

        # 自分が担当している進行中のタスク
        if(staff_id == user_id and  finish is None):
            # ユーザIDからユーザ名を取得
            user_name = functions.getUserData(user_id)

            # 現在日時を該当タスクの終了日に設定
            finish = datetime.date.today()

            # 終了日を更新
            functions.setFinish(user_id, task_id,  finish)

            message.reply("[" + task_name + "]を終了しました。終了:" + str(finish))

            # メッセージの送信先をBot用のチャンネルに変更
            message.body['channel'] = CHANNEL
            message.send(user_name +"が[" + task_name + "]を終了しました。 終了:" + str(finish) + "    終了予定:" + str(finish_schedule))
        else:
            message.reply("自分が担当していないタスクか、すでに終了したタスクです。")
    else:
        message.reply("タスクが見つかりませんでした。")


###########################################################################
### 名前      ：check_task
### 説明      ：全ての投稿を監視し、確認コマンドが入力されたら、現在進行中のタスクと
###                担当者、遅れの情報などを表示する。
### 引数      ：(message) 投稿されたメッセージに関する情報
### 戻り値    ：なし
### 参照関数  ：
###########################################################################

@listen_to(r'^\[確認\]$')
@respond_to(r'^\[確認\]$')
def check_task(message):

    # 整形したデータを格納
    data_dict = {}

    # 出力するメッセージ
    message_str = ""

    # 全てのユーザーのデータを取得
    users_data = functions.getAllUsers()

    # ヘッダー部分
    message_str +="\nNo".ljust(9)\
                +"タスク名".ljust(30, '　')\
                +"開始日".center(14)\
                +"開始予定日".center(14) \
                +"終了予定日".center(10) + "\n"

    # ユーザーループ
    for user_data in users_data:

        # 名前とIDを取り出す
        user_id = user_data[0]
        user_name = user_data[1]

        # ユーザーのタスクを取得する
        task_data = functions.getChargeTask(user_id)

        # 情報を辞書に追加
        data_dict[user_name] = task_data

        # メッセージ生成
        message_str += "\n######   *" + user_name + "*    #############\n" # ユーザ名

        # 割り当てタスクなし
        if(len(task_data) == 0):
            message_str  +=  "進行中のタスクがありません。\n"

        # タスクループ
        for task in task_data:

            # タスク情報
            message_str  +=  str(task[0]).ljust(10)\
                         +   str(task[1]).ljust(30, '　')\
                         +   str(task[3]).center(14)\
                         +   str(task[5]).center(14)\
                         +   str(task[6]).center(14) + "\n"

    message.reply(message_str)





###########################################################################
### 名前      ：create_wbs
### 説明      ：WBSを作成する
### 引数      ：(message) 投稿されたメッセージに関する情報
### 戻り値    ：なし
### 参照関数  ：sendFile
###########################################################################

@respond_to(r'^\[出力\]$')
def create_wbs(message):
    # WBSを作成しファイルのパスを格納
    path = functions.createWBS()

    # ファイル送信
    slacker = Slacker(API_TOKEN)
    # slacker.files.upload(file_=path, channels=CHANNEL)
    slacker.files.upload(file_=path, channels=message.body['channel'])






