# -*- coding: utf-8 -*-

import datetime
import mysql.connector



def db_commit():
    # 接続先サーバーを指定
    db=mysql.connector.connect(host="localhost", user="root", password="password" ,charset='utf8')
    cursor=db.cursor(buffered=True)

    # 接続先dbを指定
    cursor.execute("USE job_watcher")
    # 接続
    db.commit()

    return db, cursor

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

        print(user_name + "( " + user_id + " )の登録に成功")

    except mysql.connector.errors.IntegrityError:
        # 登録失敗時Falseを返す
        print(user_name + "( " + user_id + " )の登録に失敗")
        return False

    finally:
        cursor.close()
        db.close()

    # 登録成功時Trueを返す
    return True



def getUserName(user_id):

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
        print("名前の取得失敗")
        return None

    cursor.close()
    db.close()

    return res[0][0]



def setTask(user_id, work_days):
    return None

def getTask():
    return None

