from slackbot.bot import default_reply
from slackbot.bot import listen_to
from slackbot.bot import respond_to
from slacker import Slacker

import re
import datetime

from slackbot_settings import API_TOKEN
from consts import *

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

    message_to_sent = " :kishir:- コマンド一覧 -:kishir: \n" +\
                    "・ユーザー登録"      +    ":    [登録](str)表示名\n"+\
                    "・作業開始　　"      +    ":    [開始](str)タスク名,(int)予定日数\n"+\
                    "・作業確認　　"      +    ":    [確認]\n"+\
                    "・作業終了　　"      +    ":    [終了](int)タスクID'\n"+\
                    "・WBS出力　　 "      +   ":    [出力]"

    message.send(message_to_sent)






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

    user_name = re.sub(r'\[.*\]', "", message.body['text']).strip()
    user_id = message.body['user']

    query_to_register_user = "INSERT INTO user (id, name) VALUES ( '" + user_id + "','" + user_name  +"')"
    query_result = functions.executionQuery(query_to_register_user)

    if(query_result):
        message_to_sent = user_name + "(" + user_id + ")を作業者リストに登録しました。"
    else:
        message_to_sent = "ID( " + user_id +" )は既に登録されています。"

    message.reply(message_to_sent)



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

    input_data = re.sub(r'\[.*\]', "", message.body['text']).strip().split(",") # タグと空白を削除

    task_name = input_data[0]
    work_days = input_data[1]
    user_id = message.body["user"]

    query_to_find_user = "SELECT name FROM user WHERE id = '" + user_id + "';"
    query_to_find_user_result = functions.executionQuery( query_to_find_user )

    if(query_to_find_user_result):

        worker_name =  functions.trimQueryResult(query_to_find_user_result)

        query_to_get_last_work_finish_scheduled = "SELECT MAX(finish_schedule) FROM task WHERE staff_id = '" + user_id + "' GROUP BY staff_id;"
        query_to_get_last_work_finish_scheduled_result = functions.executionQuery(query_to_get_last_work_finish_scheduled)

        if(query_to_get_last_work_finish_scheduled_result):

            last_work_finish_scheduled = functions.trimQueryResult(query_to_get_last_work_finish_scheduled_result)
            work_start_schedule = last_work_finish_scheduled + datetime.timedelta(days=int(1))

        else:

            work_start_schedule = datetime.date.today()

        work_start = datetime.date.today()
        work_finish_schedule =  work_start_schedule + datetime.timedelta(days=int(work_days))

        query_to_register_task = "INSERT INTO task (name, staff_id, start, start_schedule, finish_schedule) VALUES ('"+ task_name +"','" +  user_id + "','" +  str(work_start) + "','" +  str(work_start_schedule) + "','" +  str(work_finish_schedule) + "');"
        query_to_register_task_result = functions.executionQuery( query_to_register_task )

        if(query_to_register_task_result):

            message.reply("[" + task_name + "]を開始します。開始:" + str(work_start) + "  開始予定:" + str(work_start_schedule) + "  終了予定:" + str(work_finish_schedule))

            message.body['channel'] = CHANNEL
            message.send(worker_name +"が[" + str(task_name) + "]を開始しました。開始:" + str(work_start) + "  開始予定:" + str(work_start_schedule) + "  終了予定:" + str(work_finish_schedule))

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

    task_id = re.sub(r'\[.*\]', "", message.body['text']).strip() # タグと空白を削除
    user_id = message.body["user"]

    query_to_get_task_data = "SELECT * FROM task WHERE id = '" + task_id + "';"
    query_to_get_task_data_result = functions.executionQuery( query_to_get_task_data )

    if(query_to_get_task_data_result):

        task_data = functions.trimQueryResult(query_to_get_task_data_result, TASK_COLUMN)

        task_name = task_data["name"]
        staff_id = task_data["staff"]
        finish = task_data["finish"]
        finish_schedule = task_data["finish_schedule"]

        if(staff_id == user_id and  finish is None):

            query_to_get_user_name = "SELECT * FROM task WHERE id = '" + task_id + "';"
            query_to_get_user_name_result = functions.executionQuery( query_to_get_user_name )

            user_name = functions.trimQueryResult(query_to_get_user_name_result)

            finish = datetime.date.today()

            functions.setFinish(user_id, task_id,  finish)

            # query_to_get_user_name = "UPDATE task SET  finish = '"+ str(finish) +"' WHERE  staff_id = '" + user_id + "' AND id = " + str(task_id) +";"
            query_to_get_user_name = "UPDATE task SET  finish = '"+ str(finish) +"' WHERE id = " + str(task_id) +";"
            query_to_get_user_name_result = functions.executionQuery( query_to_get_user_name )

            message.reply("[" + task_name + "]を終了しました。終了:" + str(finish))

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

    message_to_sent ="\nNo".ljust(9)\
                    +"タスク名".ljust(30, '　')\
                    +"開始日".center(14)\
                    +"開始予定日".center(14) \
                    +"終了予定日".center(10) + "\n"

    query_to_get_all_user = "SELECT * FROM user;"
    query_to_get_all_user_result = functions.executionQuery( query_to_get_all_user )
    users_data = functions.trimQueryResult(query_to_get_all_user_result, USER_COLUMN)

    for user_data in users_data:

        user_id = user_data["id"]
        user_name = user_data["name"]

        query_to_get_charge_task = "SELECT * FROM task WHERE staff_id = '" + user_id + "' AND finish IS NULL;"
        query_to_get_charge_task_result = functions.executionQuery( query_to_get_charge_task )
        charge_tasks_data = functions.trimQueryResult(query_to_get_charge_task_result, TASK_COLUMN)

        message_to_sent += "\n######   *" + user_name + "*    #############\n"

        if(charge_tasks_data):

            for task_data in charge_tasks_data:

                message_to_sent  +=  str(task_data["No"]).ljust(10)\
                                 +   str(task_data["name"]).ljust(30, '　')\
                                 +   str(task_data["start"]).center(14)\
                                 +   str(task_data["start_schedule"]).center(14)\
                                 +   str(task_data["finish_schedule"]).center(14) + "\n"

        else:
            message_to_sent  +=  "進行中のタスクがありません。\n"


    message.reply(message_to_sent)





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






