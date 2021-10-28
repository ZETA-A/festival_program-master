'''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

이 프로그램의 저작권은 ZETA-A(Github)에게 있습니다.
(문의 : open120488@gmail.com)

프로그램의 라이선스는 ZETA Custiom License가 적용되어있습니다.

ZETA Custiom License에 따라 사용하지않을경우 법적책임이 가해질 수 있습니다.

+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'''

from sqlite3.dbapi2 import Connection
import gspread
from gspread.models import Worksheet
from oauth2client.service_account import ServiceAccountCredentials

import sqlite3
import json
from datetime import datetime
import os
import os.path
import sys
from time import sleep
import datetime
<<<<<<< HEAD
import time
=======
>>>>>>> 6dd057e7bbbc1fbf97f6980eb6f3940145c59db7

from PyQt5.QtWidgets import QApplication, QLabel
import PyQt5
from PyQt5 import QtGui
from PyQt5 import QtCore
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtCore import QObject, QThread, pyqtSignal
from PyQt5 import QtWidgets, uic
from PyQt5.QtWidgets import *
from PyQt5.QtWidgets import QTableWidget

LOAD_DB = '_settings/outputDB.db'
LOAD_SETTING_JSON = '_settings/setting.json'
LOAD_ICON = '_settings/icon.png'
SETTING_GROUP = dict()
SERIAL = "vmfhrmfoadml-fkdltjstmrk-aksfyehldjTtmqslek"

'''
프로그램 시작
'''
<<<<<<< HEAD
# DB 파일체크
if os.path.isfile(LOAD_DB):
    print("DB 파일이 확인되었습니다")
    conn = sqlite3.connect(LOAD_DB)  # SQL 연결
    cur = conn.cursor()
    pass
else:
    print("파일이 존재하지않습니다")
    print(LOAD_DB + " 파일을 생성합니다")
    conn = sqlite3.connect(LOAD_DB)  # SQL 연결
    cur = conn.cursor()
    cur.execute(
        "CREATE TABLE log(학번 text, 딱지완료시간 text, 무궁화완료시간 text, 구슬완료시간 text, 딱지확인란 text, 무궁화확인란 text, 구슬확인란 text)")
    conn.commit()
    conn.close()

# JSON 파일체크
if os.path.isfile(LOAD_SETTING_JSON):
    print("JSON 파일이 확인되었습니다")
    pass
else:
    print("JSON 파일이 존재하지않습니다")
    print(LOAD_SETTING_JSON + " 파일을 생성합니다")
    SETTING_GROUP["URL"] = "Empty"
    SETTING_GROUP["NAME"] = "Empty"
    SETTING_GROUP["API"] = "Empty"
    with open(LOAD_SETTING_JSON, 'w', encoding="utf-8") as make_file:
        json.dump(SETTING_GROUP, make_file, indent="\t")

with open(LOAD_SETTING_JSON, 'r', encoding="utf-8") as f:
    json_data = json.load(f)
=======
def programStart():
    # DB 파일체크
    if os.path.isfile(LOAD_DB):
        print("DB 파일이 확인되었습니다")
        conn = sqlite3.connect(LOAD_DB)  # SQL 연결
        cur = conn.cursor()
        pass
    else:
        print("파일이 존재하지않습니다")
        print(LOAD_DB + " 파일을 생성합니다")
        conn = sqlite3.connect(LOAD_DB)  # SQL 연결
        cur = conn.cursor()
        cur.execute(
            "CREATE TABLE log(시간 text, 학번 text, 이름 text, 확인여부 text)")
        conn.commit()
        conn.close()

    # JSON 파일체크
    if os.path.isfile(LOAD_SETTING_JSON):
        print("JSON 파일이 확인되었습니다")
        pass
    else:
        print("JSON 파일이 존재하지않습니다")
        print(LOAD_SETTING_JSON + " 파일을 생성합니다")
        SETTING_GROUP["URL"] = "Empty"
        SETTING_GROUP["NAME"] = "Empty"
        SETTING_GROUP["API"] = "Empty"
        with open(LOAD_SETTING_JSON, 'w', encoding="utf-8") as make_file:
            json.dump(SETTING_GROUP, make_file, indent="\t")


    with open(LOAD_SETTING_JSON, 'r', encoding="utf-8") as f:
        json_data = json.load(f)
>>>>>>> 6dd057e7bbbc1fbf97f6980eb6f3940145c59db7

    spreadsheet_url = json_data['URL']
    sheet_name = json_data['NAME']
    json_file_name = json_data['API']
d = datetime.datetime.now()
licenseCheck = f"{d.year}{d.month}{d.day}"
print(licenseCheck)
if licenseCheck > f"20211030":
    print("사용기한이 지났습니다")
    os._exit(1)
else:
    print("사용기한이 남아있습니다")
    programStart()


d = datetime.datetime.now()
licenseCheck = f"{d.year}{d.month}{d.day}"

'''
구글 스프레드 시트 통신 명령
'''


def googleSheet():
    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive'
    ]

    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        json_file_name, scope)
    gc = gspread.authorize(credentials)

    # 스프레드시트 가져오기
    doc = gc.open_by_url(spreadsheet_url)

    # 시트 선택하기
    worksheet = doc.worksheet(sheet_name)

    worksheet.resize(1000, 7)
    print("시트 연결성공")

    column_data = worksheet.col_values(5 or 6 or 7)

    for index, x in enumerate(column_data):
        if x == 'O':

<<<<<<< HEAD
            cell_num_data = worksheet.acell('A' + str(index + 1)).value
            cell_DdacTime_data = worksheet.acell('B' + str(index + 1)).value
            cell_MuGungHwaTime_data = worksheet.acell(
                'C' + str(index + 1)).value
            cell_GuSeulChiGiTime_data = worksheet.acell(
                'D' + str(index + 1)).value
            cell_DdacZziChiGi_data = worksheet.acell(
                'E' + str(index + 1)).value
            cell_MuGungHwa_data = worksheet.acell('F' + str(index + 1)).value
            cell_GuSeulChiGi_data = worksheet.acell('G' + str(index + 1)).value

            print(cell_num_data)
            cur.execute("INSERT INTO log (학번, 딱지완료시간, 무궁화완료시간, 구슬완료시간, 딱지확인란, 무궁화확인란, 구슬확인란) VALUES(?, ?, ?, ?, ?, ?, ?)",
                        (cell_num_data, cell_DdacTime_data, cell_MuGungHwaTime_data, cell_GuSeulChiGiTime_data, cell_DdacZziChiGi_data, cell_MuGungHwa_data, cell_GuSeulChiGi_data))
=======
            cell_DdacTime_data = worksheet.acell('A' + str(index + 1)).value
            cell_MuGungHwaTime_data = worksheet.acell('B' + str(index + 1)).value
            cell_GuSeulChiGiTime_data = worksheet.acell('C' + str(index + 1)).value
            cell_num_data = worksheet.acell('D' + str(index + 1)).value
            cell_DdacZziChiGi_data = worksheet.acell('D' + str(index + 1)).value
            cell_MuGungHwa_data = worksheet.acell('E' + str(index + 1)).value
            cell_GuSeulChiGi_data = worksheet.acell('F' + str(index + 1)).value

            print(cell_num_data)
            cur.execute("INSERT INTO log (딱찌 완료 시간, 무궁화 완료 시간, 구슬 완료 시간, 학번, 이름, 딱치치기, 무궁화꽃, 구슬치기) VALUES(?, ?, ?, ?, ?, ?, ?, ?)",
                        (cell_DdacTime_data, cell_num_data, cell_DdacZziChiGi_data, cell_MuGungHwa_data, cell_GuSeulChiGi_data))
>>>>>>> 6dd057e7bbbc1fbf97f6980eb6f3940145c59db7
            cur.execute(
                "DELETE FROM log WHERE rowid < (SELECT max(rowid) FROM log GROUP BY 학번);")
            conn.commit()


'''
UI 구성 명령
'''
class LicenseEndWindow(QMainWindow):
    def __init__(self):
        super().__init__()

class LicenseEndWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # 오류 메시지 발생
        self.setGeometry(300, 300, 635, 198)
        self.setMinimumSize(635, 198)
        self.setMaximumSize(635, 198)
        self.setWindowTitle('라이선스가 만료되었습니다!')
        self.setWindowIcon(QIcon(LOAD_ICON))

        # 가장 큰 라벨
        self.bigLabel = QLabel(self)
        self.bigLabel.setGeometry(20, 20, 601, 51)
        self.bigLabel.setText("프로그램의 라이선스가 만료되었습니다!")
        self.bigLabel.setFont(QFont('Arial', 25))

        # 중간 라벨
        self.midLabel = QLabel(self)
        self.midLabel.setGeometry(20, 70, 471, 31)
        self.midLabel.setText("시리얼 넘버를 입력한 후, 확인을 눌러주세요")
        self.midLabel.setFont(QFont('Arial', 17))

        # 시리얼 입력 인덱스
        self.serialIndex = QLineEdit(self)
        self.serialIndex.setGeometry(20, 100, 591, 51)
        self.serialIndex.setFont(QFont('Arial', 12))

        # 작은 라벨
        self.smallLabel = QLabel(self)
        self.smallLabel.setGeometry(20, 160, 241, 21)
        self.smallLabel.setText("( 시리얼 넘버의 '-'도 모두 입력해주세요 )")
        self.smallLabel.setFont(QFont('Arial', 9))

        # 확인 버튼
        self.okBtn = QPushButton(self)
        self.okBtn.setGeometry(440, 160, 75, 23)
        self.okBtn.setText("확인")
        self.okBtn.setFont(QFont('Arial', 9))
        self.okBtn.clicked.connect(self.verify)

        # 종료 버튼
        self.cancelBtn = QPushButton(self)
        self.cancelBtn.setGeometry(530, 160, 75, 23)
        self.cancelBtn.setText("종료")
        self.cancelBtn.setFont(QFont('Arial', 9))
        self.cancelBtn.clicked.connect(self.close)

    def verify(self):
        if SERIAL == self.serialIndex.text():
            MainWindow().show()
        else:
            self.serialIndex.setText("시리얼 넘버가 틀렸습니다!")

    def close(self):
        os._exit(1)


class MainWindow(QMainWindow, QThread):
    def __init__(self):
        super().__init__()

        # 윈도우 설정
        self.setGeometry(300, 300, 1000, 715)  # x, y, w, h
        self.setMinimumSize(1000, 715)
        self.setMaximumSize(1000, 715)
        self.setWindowTitle('구글시트 카운팅 프로그램')
        self.setWindowIcon(QIcon(LOAD_ICON))

        #  SQL 테이블
        self.sqlTable = QTableWidget(self)
        self.sqlTable.setGeometry(11, 10, 981, 651)
        self.sqlTable.setRowCount(1000)
        self.sqlTable.setColumnCount(7)
        self.sqlTable.setHorizontalHeaderLabels(
            ["학번", "딱지 완료 시간", "무궁화 완료 시간", "구슬 완료 시간", "딱지 확인란", "무궁화 확인란", "구슬 확인란"])
        self.sqlTable.setColumnWidth(0, 80)
        self.sqlTable.setColumnWidth(1, 200)
        self.sqlTable.setColumnWidth(2, 200)
        self.sqlTable.setColumnWidth(3, 200)
        self.sqlTable.setColumnWidth(4, 80)
        self.sqlTable.setColumnWidth(5, 82)
        self.sqlTable.setColumnWidth(6, 80)

        # 업데이트 버튼
        self.updateBtn = QPushButton(self)
        self.updateBtn.setGeometry(310, 667, 113, 41)
        self.updateBtn.clicked.connect(googleSheet)
        self.updateBtn.clicked.connect(self.LoadData)
        self.updateBtn.setText("업데이트")

        # 세팅 버튼
        self.settingBtn = QPushButton(self)
        self.settingBtn.setGeometry(445, 667, 113, 41)
        self.settingBtn.clicked.connect(self.setting_popup)
        self.settingBtn.clicked.connect(self.OpenLoadData)
        self.settingBtn.setText("설정")

        # 종료 버튼
        self.closeBtn = QPushButton(self)
        self.closeBtn.setGeometry(580, 667, 113, 41)
        self.closeBtn.clicked.connect(self.closeProgram)
        self.closeBtn.setText("종료")

        # QDialog 설정
        self.dialog = QDialog()

    # 버튼 이벤트 함수
    def setting_popup(self):
        # QDialog 세팅
        self.sheetLabel = QLabel(self.dialog)
        self.sheetLabel.setGeometry(10, 11, 71, 21)
        self.sheetLabel.setText("시트설정")
        self.sheetLabel.setFont(QFont('Arial', 13))
        self.setWindowIcon(QIcon(LOAD_ICON))

        self.sheetPermission = QLabel(self.dialog)
        self.sheetPermission.setGeometry(10, 40, 71, 21)
        self.sheetPermission.setText("시트주소 :")
        self.sheetPermission.setFont(QFont('Arial', 11))

        self.sheetPermission_lineEdit = QLineEdit(self.dialog)
        self.sheetPermission_lineEdit.setGeometry(90, 35, 351, 31)

        self.sheetName = QLabel(self.dialog)
        self.sheetName.setGeometry(10, 86, 71, 21)
        self.sheetName.setText("시트이름 :")
        self.sheetName.setFont(QFont('Arial', 11))

        self.sheetName_lineEdit = QLineEdit(self.dialog)
        self.sheetName_lineEdit.setGeometry(90, 81, 351, 31)

        self.programSettingLabel = QLabel(self.dialog)
        self.programSettingLabel.setGeometry(10, 141, 481, 21)
        self.programSettingLabel.setText("프로그램 설정 (상대경로입력, /를 이용해 경로 작성)")
        self.programSettingLabel.setFont(QFont('Arial', 13))

        self.sheetAPILabel = QLabel(self.dialog)
        self.sheetAPILabel.setGeometry(18, 176, 61, 21)
        self.sheetAPILabel.setText("시트API :")
        self.sheetAPILabel.setFont(QFont('Arial', 11))

        self.sheetAPI_lineEdit = QLineEdit(self.dialog)
        self.sheetAPI_lineEdit.setGeometry(90, 171, 351, 31)

        self.developerLabel = QLabel(self.dialog)
        self.developerLabel.setGeometry(20, 221, 261, 21)
        self.developerLabel.setText("개발자 : 제타(GitHub : ZETA-A)")
        self.developerLabel.setFont(QFont('Arial', 9))

        self.applyBtn = QPushButton(self.dialog)
        self.applyBtn.setGeometry(290, 220, 75, 23)
        self.applyBtn.setText("적용")
        self.applyBtn.clicked.connect(self.SaveData)

        self.cancleBtn = QPushButton(self.dialog)
        self.cancleBtn.setGeometry(380, 220, 75, 23)
        self.cancleBtn.setText("닫기")
        self.cancleBtn.clicked.connect(self.dialog_close)

        self.dialog.setWindowTitle("설정")
        self.dialog.setWindowModality(Qt.ApplicationModal)
        self.dialog.resize(473, 254)
        self.dialog.setMinimumSize(473, 254)
        self.dialog.setMaximumSize(473, 254)
        self.dialog.show()

    # Dialog 닫기 이벤트

    def dialog_close(self):
        self.dialog.close()

    def closeProgram(self):
        os._exit(1)

    def SaveData(self):
        SETTING_GROUP["URL"] = self.sheetPermission_lineEdit.text()
        SETTING_GROUP["NAME"] = self.sheetName_lineEdit.text()
        SETTING_GROUP["API"] = self.sheetAPI_lineEdit.text()

        with open(LOAD_SETTING_JSON, 'w', encoding='utf-8') as make_file:
            json.dump(SETTING_GROUP, make_file, indent="\t")

        spreadsheet_url = json_data['URL']
        sheet_name = json_data['NAME']
        json_file_name = json_data['API']

    def OpenLoadData(self):
        with open(LOAD_SETTING_JSON, 'r') as f:
            json_data = json.load(f)

        self.sheetPermission_lineEdit.setText(json_data["URL"])
        self.sheetName_lineEdit.setText(json_data["NAME"])
        self.sheetAPI_lineEdit.setText(json_data["API"])

    def LoadData(self):
        connection = sqlite3.connect(LOAD_DB)
        cur = connection.cursor()
        sqlstr = 'SELECT * FROM log LIMIT 1000'
        tablerow = 0
        results = cur.execute(sqlstr)
        for row in results:
            self.sqlTable.setItem(
                tablerow, 0, QtWidgets.QTableWidgetItem(row[0]))
            self.sqlTable.setItem(
                tablerow, 1, QtWidgets.QTableWidgetItem(row[1]))
            self.sqlTable.setItem(
                tablerow, 2, QtWidgets.QTableWidgetItem(row[2]))
            self.sqlTable.setItem(
                tablerow, 3, QtWidgets.QTableWidgetItem(row[3]))
            self.sqlTable.setItem(
                tablerow, 4, QtWidgets.QTableWidgetItem(row[4]))
            self.sqlTable.setItem(
                tablerow, 5, QtWidgets.QTableWidgetItem(row[5]))
            self.sqlTable.setItem(
                tablerow, 6, QtWidgets.QTableWidgetItem(row[6]))
            tablerow += 1


if __name__ == '__main__':
    app = QApplication(sys.argv)
    if licenseCheck > f"20220111":
        mainWindow = LicenseEndWindow()
    elif licenseCheck < f"20220111":
        mainWindow = MainWindow()
    mainWindow.show()
    sys.exit(app.exec_())
