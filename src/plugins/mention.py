from slackbot.bot import default_reply
from slackbot.bot import listen_to
from slackbot.bot import respond_to

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
    message.send(' - コマンド一覧 - \n\
        ・登録コマンド：[登録](str[a-zA-Z])表示名\n\
        ・開始コマンド：[開始](str)タスク名,(int)予定日数\n\
        ・終了コマンド：[終了](str)タスク名\n\
        ・確認コマンド：[確認]')





###########################################################################
### 名前      ：regist_user
### 説明      ：Botへのメンションを監視し、登録コマンドが入力されたら、
###                   DBにユーザーの情報を登録する。
### 引数      ：(message) 情報
### 戻り値    ：なし
### 参照関数  ：
###########################################################################

@respond_to(r'^\[登録\][a-zA-Z]+$')
def regist_user(message):

    # メッセージを取り出す
    text = message.body['text']

    # ユーザーIDを取り出す
    user_id = message.body['user']

    # メッセージテキストから表示名を取得
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
### 引数      ：(message) 情報
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

    #作業日数をもとに開始日と終了日を計算
    start = datetime.date.today()
    finish_schedule = start + datetime.timedelta(days=int(work_days))

    # ユーザーidを取り出す
    user_id = message.body["user"]

    # 作業者が登録されているか確認する
    if(functions.getUserName( user_id ) is not None):

        # 作業者名を取得
        user_name = functions.getUserName( user_id )

        # TODO : 仕事をしていない日などないはずなので開始予定日は前回タスク終了予定日の1日後とする
        # タスクを登録する
        if(functions.setTask(task_name, user_id, start, start, finish_schedule)):
            message.reply("[" + task_name + "]を開始します。開始:" + str(start) + "  終了予定:" + str(finish_schedule))

            # メッセージの送信先をBot用のチャンネルに変更しメッセージを送信
            message.body['channel'] = CHANNEL
            message.send(user_name +"が[" + task_name + "]を開始しました。開始:" + str(start) + "  終了予定:" + str(finish_schedule))
        else:
            message.reply("[" + task_name + "]の登録に失敗しました。")

    else:
        message.reply("あなた(" + user_id + ")は作業者ではありません。")





###########################################################################
### 名前      ：finish_task
### 説明      ：Botへのメンションを監視し、終了コマンドが入力されたら、
###                 DBに終了日時を登録し、遅れなどの情報を表示する
### 引数      ：(message) 情報
### 戻り値    ：なし
### 参照関数  ：
###########################################################################

@respond_to(r'^\[終了\].+$')
def finish_task(message):

    # メッセージを取り出す
    text = message.body['text']

    # メッセージテキストからタスク名を取得
    task = re.sub(r'\[.*\]', "", text).strip() # タグと空白を削除


    # 現在日時を該当タスクの終了日に設定
    finish = datetime.date.today()

    # TODO : 外部ファイルのメソッドを呼び出し、DBに登録する処理を行う

    message.reply("[" + task + "]を終了しました。終了:" + str(finish))

    # メッセージの送信先をBot用のチャンネルに変更
    message.body['channel'] = CHANNEL
    #message.send(user_name +"が[" + task + "]を終了しました。終了予定:" + str(finish_schedule) + "  終了:" + str(finish))





###########################################################################
### 名前      ：check_task
### 説明      ：全ての投稿を監視し、確認コマンドが入力されたら、現在進行中のタスクと
###                担当者、遅れの情報などを表示する。
###
### 引数      ：(message) 情報
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

    # 格納ループ
    for user_data in users_data:

        # 名前とIDを取り出す
        user_id = user_data[0]
        user_name = user_data[1]

        # ユーザーのタスクを取得する
        task_data = functions.getTask(user_id)

        # 情報を辞書に追加
        data_dict[user_name] = task_data

        # メッセージ生成
        message_str += "###    担当者 : " + user_name + "    ##########\n" # ユーザ名

        for task in task_data:

            # タスク情報
            message_str +=  "ID:"    + str(task[0])\
                         +   "  名前:"   + task[1]\
                         +   "  開始日:"     + str(task[3])\
                         +   "  開始予定日:" + str(task[5])\
                         +   "  終了予定日:" + str(task[6]) + "\n"

    message.reply(message_str)





