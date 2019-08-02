# -*- coding: utf-8 -*-

import mysql.connector
import datetime
import openpyxl as excel
import time
import requests

###########################################################################
### 名前      ：db_commit
### 説明      ：DBに接続する
### 引数      ：なし
### 戻り値    ：接続情報インスタンス
### 参照関数  ：
###########################################################################
def db_commit():
    # 接続先サーバーを指定
    db=mysql.connector.connect(host="localhost", user="root", password="password" ,charset='utf8')
    cursor=db.cursor(buffered=True)

    # 接続先dbを指定
    cursor.execute("USE job_watcher")
    # 接続
    db.commit()

    return db, cursor


###########################################################################
### 名前      ：setUser
### 説明      ：ユーザーを登録する
### 引数      ：(str)ユーザーID
### 戻り値    ：(str)ユーザー名
### 参照関数  ：
###########################################################################
def setUser(user_id, user_name):

    try:
        # dbに接続
        db, cursor = db_commit()

        # userテーブルにユーザーidとユーザー名を登録
        query = "INSERT INTO user (id, name) VALUES (%s, %s);"

        # 渡す値を()で囲まないと動かないので注意
        cursor.execute(query, (user_id, user_name) )

        # 実行
        db.commit()

    except mysql.connector.errors.IntegrityError:
        # 登録失敗時Falseを返す
        return False

    finally:
        cursor.close()
        db.close()

    # 登録成功時Trueを返す
    return True



###########################################################################
### 名前      ：getUserName
### 説明      ：ユーザーIDからユーザーのデータを取得する
### 引数      ：(str)ユーザーID
### 戻り値    ：(str)ユーザー名
### 参照関数  ：
###########################################################################
def getUserData(user_id):

    try:
        # dbに接続
        db, cursor = db_commit()

        # userテーブルのidと一致するユーザーのnameを取得
        query = "SELECT name FROM user WHERE id = '" + user_id + "';"

        # 実行
        cursor.execute(query)

        # 結果を取得
        res = cursor.fetchall()



    except mysql.connector.errors.IntegrityError:
        return None

    finally:
        cursor.close()
        db.close()

    return res[0][0]


###########################################################################
### 名前      ：setTask
### 説明      ：開始したタスクの情報を登録する
### 引数      ：(str)タスクの名前, (str)担当者ID, (date)開始日, (date)終了予定日
### 戻り値    ：(bool)結果
### 参照関数  ：
###########################################################################
def setTask(name, staff_id, start, start_schedule, finish_schedule):

    try:
        # dbに接続
        db, cursor = db_commit()

        # userテーブルにユーザーidとユーザー名を登録
        query = "INSERT INTO task (name, staff_id, start, start_schedule, finish_schedule) VALUES (%s, %s, %s, %s, %s);"

        # 渡す値を()で囲まないと動かないので注意
        cursor.execute(query, (name, staff_id, start, start_schedule, finish_schedule) )

        # 実行
        db.commit()

    except mysql.connector.errors.IntegrityError:
        # 登録失敗時Falseを返す
        return False

    finally:
        cursor.close()
        db.close()

    return True




###########################################################################
### 名前      ：setFinish
### 説明      ：終了したタスクの終了日を設定する
### 引数      ：(str)ユーザーID, (int)タスクID, (date)終了日
### 戻り値    ：担当しているタスク情報
### 参照関数  ：
###########################################################################
def setFinish(user_id, task, finish):
    try:
        # dbに接続
        db, cursor = db_commit()

        # IDが指定されている場合
        query = "UPDATE task SET  finish = '"+ str(finish) +"' WHERE  staff_id = '" + user_id + "' AND id = " + str(task) +";"

        # 実行
        cursor.execute(query)

        # 反映？
        db.commit()

    except mysql.connector.errors.IntegrityError:
        # 登録失敗時Falseを返す
        return False

    finally:
        cursor.close()
        db.close()

    # 登録成功時Trueを返す
    return True





###########################################################################
### 名前      ：getAllTask
### 説明      ：全てのタスク
### 引数      ：なし
### 戻り値    ：
### 参照関数  ：
###########################################################################
def getAllTask():

    res = None;

    try:
        # dbに接続
        db, cursor = db_commit()

        # userテーブルのidと一致するユーザーのnameを取得
        query = "SELECT * FROM task;"

        # 実行
        cursor.execute(query)

        # 結果を取得
        res = cursor.fetchall()

    except mysql.connector.errors.IntegrityError:
        # 取得失敗時Noneを返す
        return None

    finally:
        cursor.close()
        db.close()

    # 登録成功時、タスクの情報を返す
    return res




###########################################################################
### 名前      ：getTaskData
### 説明      ：特定のタスクの名前を取得
### 引数      ：(task_id)ユーザーID
### 戻り値    ：
### 参照関数  ：
###########################################################################
def getTaskData(task_id):

    res = None;

    try:
        # dbに接続
        db, cursor = db_commit()

        # userテーブルのidと一致するユーザーのnameを取得
        query = "SELECT * FROM task WHERE id = '" + task_id + "';"

        # 実行
        cursor.execute(query)

        # 結果を取得
        res = cursor.fetchall()

    except mysql.connector.errors.IntegrityError:
        # 登録失敗時Noneを返す
        return None

    finally:
        cursor.close()
        db.close()

    # 登録成功時、タスクの情報を返す
    return res


###########################################################################
### 名前      ：getChargeTask
### 説明      ：特定のユーザーが担当しているタスクの情報を取得
### 引数      ：(user_id)ユーザーID
### 戻り値    ：担当しているタスク情報
### 参照関数  ：
###########################################################################
def getChargeTask(user_id):

    try:
        # dbに接続
        db, cursor = db_commit()

        # taskテーブルから特定のユーザーが担当しているタスクの情報を取得
        query = "SELECT * FROM task WHERE staff_id = '" + user_id + "' AND finish IS NULL;"

        # 実行
        cursor.execute(query)

        # 結果を取得
        res = cursor.fetchall()


    except mysql.connector.errors.IntegrityError:
        return None

    finally:
        cursor.close()
        db.close()

    return res



###########################################################################
### 名前      ：getAllUsers
### 説明      ：全てのユーザーの全情報を取得
### 引数      ：なし
### 戻り値    ：ユーザー情報
### 参照関数  ：
###########################################################################
def getAllUsers():

    try:
        # dbに接続
        db, cursor = db_commit()

        # userテーブルのすべてのユーザー情報を取得
        query = "SELECT * FROM user;"

        # 実行
        cursor.execute(query)

        # 結果を取得
        res = cursor.fetchall()

    except mysql.connector.errors.IntegrityError:

        return None

    finally:
        cursor.close()
        db.close()

    return res




###########################################################################
### 名前      ：get_last_finish__schedule
### 説明      ：指定のユーザーが担当するタスクの中から終了予定日が最大のものを取得
### 引数      ：(user_id)取得するユーザーのID
### 戻り値    ：終了予定日
### 参照関数  ：
###########################################################################

def get_last_finish_schedule(user_id):

    try:
        # dbに接続
        db, cursor = db_commit()

        # userテーブルのすべてのユーザー情報を取得
        query = "SELECT MAX(finish_schedule) FROM task WHERE staff_id = '" + user_id + "' GROUP BY staff_id;"

        # 実行
        cursor.execute(query)

        # 結果を取得
        res = cursor.fetchall()

    except mysql.connector.errors.IntegrityError:
        return None

    finally:
        cursor.close()
        db.close()

    # データが取得できなかった場合
    if(len(res) != 0):
        return res[0][0]
    else:
        return datetime.date.today()




###########################################################################
### 名前      ：sendFile
### 説明      ：ファイルを送信
### 引数      ：(user_id)取得するユーザーのID
### 戻り値    ：終了予定日
### 参照関数  ：
###########################################################################
def sendFile(token, channel, filepath):
    print('')


###########################################################################
### 名前      ：createWBS
### 説明      ：CSVファイルを作成
### 引数      ：(user_id)取得するユーザーのID
### 戻り値    ：終了予定日
### 参照関数  ：getUserName()
###########################################################################

def createWBS():

    # 新規ワークブックを作成
    wb = excel.Workbook()
    ws = wb.active

    # column = ["No", "タスク", "", "開始予定日", "終了予定日", "開始日", "終了日", "担当者", "ステータス"]
    column = ["No", "タスク", "担当者", "開始日", "終了日", "開始予定日", "終了予定日",  "ステータス"]
    # ヘッダー部を作成
    for i in range(1, len(column) + 1):
        ws.cell(column = i, row = 2, value = column[i - 1])

    # タスク情報を取得
    tasks = getAllTask()

    # キー:ユーザーID, 値:ユーザ名 の辞書を作成
    users_data_dict = {}
    users_data = getAllUsers()

    for user_data in users_data:
        users_data_dict.setdefault(user_data[0], user_data[1])


    # タスク情報書き込み
    for i, task in enumerate(tasks): # 縦
        for j, data in enumerate(task): # 横

            # ユーザー名
            if(j == 2):
                # ユーザー名を表示する処理
                if(data in users_data_dict):
                    ws.cell(column = j + 1, row = i + 3, value = users_data_dict[data])
                    continue

            # 終了日
            if(j == 4):
                # ステータス設定
                if(data is not None):
                    ws.cell(column = j + 1, row = i + 3, value = str(data))
                    ws.cell(column = len(task) + 1, row = i + 3, value = "完了")
                    continue
                else:
                    ws.cell(column = j + 1, row = i + 3, value = "")
                    ws.cell(column = len(task) + 1, row = i + 3, value = "作業中")
                    print("作業中", j, data)
                    continue

            ws.cell(column = j + 1, row = i + 3, value = str(data))


    # 保存
    now = time.time()
    wb.save("../files/WBS" + str(now) + ".xlsx")

    print("実行しました")
    return(now)













