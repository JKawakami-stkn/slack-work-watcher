# -*- coding: utf-8 -*-
# 日時
import datetime
import time
# mysql
import mysql.connector

# excel操作
import openpyxl as excel
# from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from openpyxl.styles import Border, Side, PatternFill



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
### 名前      ：createWBS
### 説明      ：CSVファイルを作成
### 引数      ：(user_id)取得するユーザーのID
### 戻り値    ：終了予定日
### 参照関数  ：getUserName()
###########################################################################

# TODO : 予定より開始が早いタスクと予定より遅れているタスクをぱっと見で分かるように色分けする
def createWBS():

    # 新規ワークブックを作成
    wb = excel.Workbook()
    ws = wb.active

    # column = ["No", "タスク", "", "開始予定日", "終了予定日", "開始日", "終了日", "担当者", "ステータス"]
    column = ["No", "タスク", "担当者", "開始日", "終了日", "開始予定日", "終了予定日",  "ステータス"]

    # セルの枠線
    thin = Side(border_style="thin", color="000000")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)

    # 操作対処のセル
    this_cell = None

    # プロジェクト開始日
    project_start_date = None

    # ヘッダー部を作成
    for i in range(1, len(column) + 1):

        # セルを取得
        this_cell = ws.cell(column = i, row = 2, value = column[i - 1])
        this_cell.fill = PatternFill(fill_type='solid',fgColor='ffde8c') # セルの色
        this_cell.border = border # 枠線

        # 下のセルと結合
        ws.merge_cells( start_row = 2, start_column = i, end_row = 3, end_column=i )

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

            # TODO : 変わるvalueのみを変数に格納しborder処理を減らす

            width =  i + 4
            height = j + 1

            # プロジェクトの開始日を取得
            if(i == 0 and j== 3):
                project_start_date = data

            # ユーザー名
            if(j == 2):
                # ユーザー名を表示する処理
                if(data in users_data_dict):
                    this_cell = ws.cell(column = height, row = width, value = users_data_dict[data])
                    this_cell.border = border # 枠線
                    continue

            # 終了日
            if(j == 4):
                # ステータス設定
                if(data is not None):
                    this_cell = ws.cell(column = height, row = width, value = str(data))
                    this_cell.border = border # 枠線
                    this_cell = ws.cell(column = len(task) + 1, row = width, value = "完了")
                    this_cell.border = border # 枠線
                    continue
                else:
                    this_cell = ws.cell(column = height, row = width, value = "")
                    this_cell.border = border # 枠線
                    this_cell = ws.cell(column = len(task) + 1, row = width, value = "作業中")
                    this_cell.border = border # 枠線
                    continue

            this_cell = ws.cell(column = j + 1, row = width, value = str(data))

            this_cell.border = border # 枠線

    # 曜日名
    weekname = ["月","火","水","木","金","土", "日"]

    # 月
    this_month = 0



    # 366日分を繰り返してセルに書き込む
    tm = project_start_date

    # 月の始まりのセル
    start_month_width  = 9

    for i in range(1, 367):

        width =  i + 8
        height = 1

        # 月初めのみ出力
        if(this_month != tm.month):
            this_cell = ws.cell(column = width, row = height, value=tm.month)
            this_month = tm.month
            this_cell.border = border # 枠線
            # 最初の月じゃない場合に結合
            # FIXME : 最後の月が結合されない
            if(i != 1):
                # セルと結合
                ws.merge_cells( start_row = height, start_column = start_month_width, end_row = height, end_column = width - 1)
                start_month_width = width

        this_cell = ws.cell(column=i + 8, row=2, value=tm.day)
        this_cell.border = border # 枠線
        this_cell = ws.cell(column=i + 8, row=3, value=weekname[tm.weekday()])
        this_cell.border = border # 枠線

        # 翌日
        tm = tm + datetime.timedelta(days=1)


    # プロジェクト開始日のセル
    project_start_sell = 9
    project_start_date

    # 予定出力
    for i, data in enumerate(tasks):

        # タスク開始予定セルを取得
        start_cell = (data[5] - project_start_date).days + project_start_sell

        # タスク終了予定セルを取得
        finish_cell = (data[6] - project_start_date).days + project_start_sell

        # せる塗りつぶし
        for j in range(start_cell, finish_cell + 1):
            this_cell = ws.cell(column = j, row = i + 4)
            this_cell.fill = PatternFill(fill_type='solid',fgColor='f4ff78') # セルの色


    # 実績出力
    for i, data in enumerate(tasks):

        # タスク開始セルを取得
        start_cell = (data[3] - project_start_date).days + project_start_sell

        # タスク終了セルを取得
        if(data[4] is not None):
            finish_cell = (data[4] - project_start_date).days + project_start_sell
        else:
            finish_cell = (datetime.date.today() - project_start_date).days + project_start_sell

        # せる塗りつぶし
        for j in range(start_cell, finish_cell + 1):
            this_cell = ws.cell(column = j, row = i + 4)
            this_cell.fill = PatternFill(fill_type='solid',fgColor='4aaeff') # セルの色


    # 保存
    now = time.time()
    path = "../files/WBS" + str(now) + ".xlsx"
    wb.save(path)


    return(path)






