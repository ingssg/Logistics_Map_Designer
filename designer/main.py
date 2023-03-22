# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


# def print_hi(name):
# Use a breakpoint in the code line below to debug your script.
# print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
# if __name__ == '__main__':
#    print_hi('PyCharm')*/

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
import os
import sys
import random

from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtCore import pyqtSlot, Qt
from PyQt5.QtGui import *
from PyQt5.QtCore import pyqtSlot, Qt
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import QFileInfo
import random
# import pandas as pd
import xlsxwriter
# from openpyxl import load_workbook
import openpyxl
from PyQt5.uic.properties import QtGui

import pymysql
from pymysql.constants import CLIENT

conn = None
cur = None

conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='1290', db='lghpdb', charset='utf8',
                       client_flag=CLIENT.MULTI_STATEMENTS, autocommit=True)
cur = conn.cursor()


def resource_path(relative_path):
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


# 1.homePage.ui
form = resource_path('homePage.ui')  # 여기에 ui파일명 입력
form_class = uic.loadUiType(form)[0]
# 2.setGrid.ui
form_second = resource_path('setGrid.ui')
form_secondwindow = uic.loadUiType(form_second)[0]
# 3.setAttribute.ui
form_third = resource_path('setAttribute.ui')
form_thirdwindow = uic.loadUiType(form_third)[0]
# 4.createMap.ui
form_fourth = resource_path('createMap.ui')
form_fourthwindow = uic.loadUiType(form_fourth)[0]
# 5.viewFile.ui
form_fifth = resource_path('viewFile.ui')
form_fifthwindow = uic.loadUiType(form_fifth)[0]
# 6.sixthFile.ui
form_sixth = resource_path('editFile.ui')
form_sixthwindow = uic.loadUiType(form_sixth)[0]


# 1.homePage.ui
class WindowClass(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle("맵 디자인")
        self.setFixedSize(1000, 800)
        self.createFile.clicked.connect(self.btn_createfile_to_setgrid)  # createFile button 클릭
        self.openFile.clicked.connect(self.btn_fileLoad)  # openFile button 클릭

    # 여기에 시그널-슬롯 연결 설정 및 함수 설정.
    # -createFile button 함수: setGrid.ui로 창전환
    def btn_createfile_to_setgrid(self):
        self.hide()
        self.second = secondwindow()
        self.second.exec_()
        self.show()

    # -openFile button 함수: 파일선택창
    def btn_fileLoad(self):
        # QFileDialog.getOpenFileName(self, '', '', 'xlsx파일 (*.xlsx);; All File(*)')  # !!저장파일 타입 정해지면, 확장자에 추가
        # 미리보기ui연결 수정클릭->수정페이지
        self.hide()
        self.fifth = fifthwindow()
        self.fifth.exec_()
        self.show()  # homepage로 돌아감


# 5. viewFile.ui
class fifthwindow(QDialog, QWidget, form_fifthwindow):
    def __init__(self, parent=None):
        global row, col, file_name
        super(fifthwindow, self).__init__()
        # self.initUi()
        self.setupUi(self)
        self.setWindowTitle("맵 미리보기")
        self.show()  # 파일선택후 창이 앞으로 띄워지게 하기위해 위에 위치
        self.setFixedSize(1000, 800)
        file = QFileDialog.getOpenFileName(self, '', '', 'xlsx파일 (*.xlsx);; All File(*)')  # !!저장파일 타입 정해지면, 확장자에 추가
        global filename  # 선언, 할당 분리
        filename = file[0]
        load_xlsx = openpyxl.load_workbook(file[0], data_only=True)
        load_sheet = load_xlsx['NewSheet1']
        self.table = QTableWidget(parent)

        # 파일 이름으로 db에서 해당 정보 연결
        file_name = QFileInfo(file[0]).baseName()
        sql = "CALL deleteProject('p1'); CALL createProject('p1', NULL, NULL, NULL); CALL createSimul('p1', 's1'); CALL updateSimulName(%s, 's1');"
        cur.execute(sql, [str(file_name)])

        # self._mainwin=parent
        vbox = QVBoxLayout()
        vbox.addWidget(self.table)
        grid = QGridLayout()
        vbox.addLayout(grid)
        edit = QPushButton("수정")
        grid.addWidget(edit, 0, 0)
        self.setLayout(vbox)
        self.setGeometry(200, 200, 400, 500)
        # max_column,max_row는 text 있는 셀까지만 셈. db연동 필요
        # print(load_sheet.max_column)
        # print(load_sheet.max_row)
        sql = "SELECT GridSizeX FROM grid " + "WHERE Grid_ID = %s;"
        cur.execute(sql, [str(file_name)])
        file_col = cur.fetchone()
        sql = "SELECT GridSizeY FROM grid " + "WHERE Grid_ID = %s;"
        cur.execute(sql, [str(file_name)])
        file_row = cur.fetchone()
        row = int(file_row[0])
        col = int(file_col[0])
        self.table.setColumnCount(col)
        self.table.setRowCount(row)
        # 반드시 item 생성해야 셀 색상 변경가능
        for i in range(row):
            for j in range(col):
                self.table.setItem(i, j, QTableWidgetItem())
        self.table.resizeColumnsToContents()
        self.table.resizeRowsToContents()
        # load_excel은 1부터, table은 0부터
        for i in range(1, row + 1):
            for j in range(1, col + 1):
                if load_sheet.cell(i, j).value == "y":
                    self.table.item(i - 1, j - 1).setBackground(Qt.yellow)
                    self.table.item(i - 1, j - 1).setText("y")
                    self.table.item(i - 1, j - 1).setForeground(Qt.yellow)
                if load_sheet.cell(i, j).value == "b":
                    self.table.item(i - 1, j - 1).setBackground(Qt.darkBlue)
                    self.table.item(i - 1, j - 1).setText("b")
                    self.table.item(i - 1, j - 1).setForeground(Qt.darkBlue)
                if load_sheet.cell(i, j).value == "g":
                    self.table.item(i - 1, j - 1).setBackground(Qt.darkGreen)
                    self.table.item(i - 1, j - 1).setText("g")
                    self.table.item(i - 1, j - 1).setForeground(Qt.darkGreen)
                if load_sheet.cell(i, j).value == "r":
                    self.table.item(i - 1, j - 1).setBackground(Qt.red)
                    self.table.item(i - 1, j - 1).setText("r")
                    self.table.item(i - 1, j - 1).setForeground(Qt.red)
                if load_sheet.cell(i, j).value == "d":
                    self.table.item(i - 1, j - 1).setBackground(Qt.darkGray)
                    self.table.item(i - 1, j - 1).setText("d")
                    self.table.item(i - 1, j - 1).setForeground(Qt.darkGray)
        edit.clicked.connect(self.btn_edit)
        # self.show() #파일 선택후 맵미리보기창이 뒤에 뜨게됨

    def btn_edit(self):
        self.hide()
        self.fifth = sixthwindow()
        self.fifth.exec_()


# 6. editFile.ui
class sixthwindow(QDialog, QWidget, form_sixthwindow):
    def __init__(self, parent=None):
        global temp_count_len, temp_count_wid, row, col, file_name, file_grid
        temp_count_len = int(col)
        temp_count_wid = int(row)
        sql = "SELECT * FROM grid " + "WHERE Grid_ID = %s;"
        cur.execute(sql, [str(file_name)])
        file_grid = cur.fetchone()

        super(sixthwindow, self).__init__()
        # self.initUi()
        # self.setupUi(self)
        self.setWindowTitle("맵 수정하기")
        self.setFixedSize(1000, 800)
        self.table = QTableWidget(parent)
        vbox = QVBoxLayout()
        vbox.addWidget(self.table)
        grid = QGridLayout()
        vbox.addLayout(grid)
        charge = QPushButton("충전")
        grid.addWidget(charge, 0, 0)
        chute = QPushButton("슈트")
        grid.addWidget(chute, 0, 1)
        ws = QPushButton("워크스테이션")
        grid.addWidget(ws, 0, 2)
        buffer = QPushButton("버퍼")
        grid.addWidget(buffer, 1, 0)
        block = QPushButton("블락")
        grid.addWidget(block, 1, 1)
        trash = QPushButton("삭제")
        grid.addWidget(trash, 0, 5)
        clear = QPushButton("초기화")
        grid.addWidget(clear, 1, 5)
        addrow = QPushButton("row추가")
        grid.addWidget(addrow, 0, 3)
        addcol = QPushButton("col추가")
        grid.addWidget(addcol, 0, 4)
        delrow = QPushButton("row삭제")
        grid.addWidget(delrow, 1, 3)
        delcol = QPushButton("col삭제")
        grid.addWidget(delcol, 1, 4)
        save = QPushButton("저장")
        grid.addWidget(save, 2, 6)
        self.setLayout(vbox)
        self.setGeometry(200, 200, 400, 500)
        # global변수 사용하기(file이름)
        load_xlsx = openpyxl.load_workbook(filename, data_only=True)
        load_sheet = load_xlsx['NewSheet1']
        # row = load_sheet.max_row
        # col = load_sheet.max_column
        self.table.setColumnCount(col)
        self.table.setRowCount(row)
        # 반드시 item 생성해야 셀 색상 변경가능
        for i in range(row):
            for j in range(col):
                self.table.setItem(i, j, QTableWidgetItem())
        self.table.resizeColumnsToContents()
        self.table.resizeRowsToContents()
        # load_excel은 1부터, table은 0부터
        for i in range(1, row + 1):
            for j in range(1, col + 1):
                if load_sheet.cell(i, j).value == "y":
                    self.table.item(i - 1, j - 1).setBackground(Qt.yellow)
                    self.table.item(i - 1, j - 1).setText("y")
                    self.table.item(i - 1, j - 1).setForeground(Qt.yellow)
                if load_sheet.cell(i, j).value == "b":
                    self.table.item(i - 1, j - 1).setBackground(Qt.darkBlue)
                    self.table.item(i - 1, j - 1).setText("b")
                    self.table.item(i - 1, j - 1).setForeground(Qt.darkBlue)
                if load_sheet.cell(i, j).value == "g":
                    self.table.item(i - 1, j - 1).setBackground(Qt.darkGreen)
                    self.table.item(i - 1, j - 1).setText("g")
                    self.table.item(i - 1, j - 1).setForeground(Qt.darkGreen)
                if load_sheet.cell(i, j).value == "r":
                    self.table.item(i - 1, j - 1).setBackground(Qt.red)
                    self.table.item(i - 1, j - 1).setText("r")
                    self.table.item(i - 1, j - 1).setForeground(Qt.red)
                if load_sheet.cell(i, j).value == "d":
                    self.table.item(i - 1, j - 1).setBackground(Qt.darkGray)
                    self.table.item(i - 1, j - 1).setText("d")
                    self.table.item(i - 1, j - 1).setForeground(Qt.darkGray)

        # self.table.selectionModel().selectionChanged.connect(self.on_selection)
        charge.clicked.connect(self.btn_charge)  # charge button 클릭
        chute.clicked.connect(self.btn_chute)  # chute button 클릭
        ws.clicked.connect(self.btn_ws)  # ws button 클릭
        buffer.clicked.connect(self.btn_buffer)  # buffer button 클릭
        block.clicked.connect(self.btn_block)  # block button 클릭
        trash.clicked.connect(self.btn_trash)  # trash button 클릭
        clear.clicked.connect(self.btn_clear)  # clear button 클릭
        addrow.clicked.connect(self.btn_addrow)  # addRow button 클릭
        addcol.clicked.connect(self.btn_addcol)  # addCol button 클릭
        delrow.clicked.connect(self.btn_delrow)  # delRow button 클릭
        delcol.clicked.connect(self.btn_delcol)  # delCol button 클릭
        save.clicked.connect(self.btn_save_map)  # saveMap button 클릭
        self.show()

        # def on_selection(self,selected):
        #    for ix in selected.indexes():
        #        print('select row: {0}, col: {1}'.format(ix.row(),ix.column()))

    @pyqtSlot()
    def btn_charge(self):
        global yellow, red, green, blue, gray, file_grid
        i_charge = file_grid[13]
        for ix in self.table.selectedIndexes():
            # print('s r:{0},c:{1}'.format(ix.row(),ix.column()))
            if i_charge == 1:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.yellow)
                self.table.item(ix.row(), ix.column()).setText("y")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.yellow)
                yellow = 1
            elif i_charge == 2:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.red)
                self.table.item(ix.row(), ix.column()).setText("r")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.red)
                red = 1
            elif i_charge == 3:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkGreen)
                self.table.item(ix.row(), ix.column()).setText("g")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkGreen)
                green = 1
            elif i_charge == 4:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkBlue)
                self.table.item(ix.row(), ix.column()).setText("b")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkBlue)
                blue = 1
            else:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkGray)
                self.table.item(ix.row(), ix.column()).setText("d")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkGray)
                gray = 1

    @pyqtSlot()
    def btn_chute(self):
        global yellow, red, green, blue, gray, file_grid
        i_chute = file_grid[14]
        for ix in self.table.selectedIndexes():
            # print('s r:{0},c:{1}'.format(ix.row(),ix.column()))
            if i_chute == 1:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.yellow)
                self.table.item(ix.row(), ix.column()).setText("y")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.yellow)
                yellow = 2
            elif i_chute == 2:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.red)
                self.table.item(ix.row(), ix.column()).setText("r")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.red)
                red = 2
            elif i_chute == 3:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkGreen)
                self.table.item(ix.row(), ix.column()).setText("g")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkGreen)
                green = 2
            elif i_chute == 4:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkBlue)
                self.table.item(ix.row(), ix.column()).setText("b")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkBlue)
                blue = 3
            else:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkGray)
                self.table.item(ix.row(), ix.column()).setText("d")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkGray)
                gray = 2

    @pyqtSlot()
    def btn_ws(self):
        global yellow, red, green, blue, gray, file_grid
        i_ws = file_grid[15]
        for ix in self.table.selectedIndexes():
            # print('s r:{0},c:{1}'.format(ix.row(),ix.column()))
            if i_ws == 1:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.yellow)
                self.table.item(ix.row(), ix.column()).setText("y")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.yellow)
                yellow = 3
            elif i_ws == 2:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.red)
                self.table.item(ix.row(), ix.column()).setText("r")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.red)
                red = 3
            elif i_ws == 3:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkGreen)
                self.table.item(ix.row(), ix.column()).setText("g")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkGreen)
                green = 3
            elif i_ws == 4:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkBlue)
                self.table.item(ix.row(), ix.column()).setText("b")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkBlue)
                blue = 3
            else:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkGray)
                self.table.item(ix.row(), ix.column()).setText("d")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkGray)
                gray = 3

    @pyqtSlot()
    def btn_buffer(self):
        global yellow, red, green, blue, gray, file_grid
        i_buf = file_grid[16]
        for ix in self.table.selectedIndexes():
            # print('s r:{0},c:{1}'.format(ix.row(),ix.column()))
            if i_buf == 1:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.yellow)
                self.table.item(ix.row(), ix.column()).setText("y")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.yellow)
                yellow = 4
            elif i_buf == 2:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.red)
                self.table.item(ix.row(), ix.column()).setText("r")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.red)
                red = 4
            elif i_buf == 3:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkGreen)
                self.table.item(ix.row(), ix.column()).setText("g")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkGreen)
                green = 4
            elif i_buf == 4:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkBlue)
                self.table.item(ix.row(), ix.column()).setText("b")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkBlue)
                blue = 4
            else:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkGray)
                self.table.item(ix.row(), ix.column()).setText("d")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkGray)
                gray = 4

    @pyqtSlot()
    def btn_block(self):
        global yellow, red, green, blue, gray, file_grid
        i_blk = file_grid[17]
        for ix in self.table.selectedIndexes():
            # print('s r:{0},c:{1}'.format(ix.row(),ix.column()))
            if i_blk == 1:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.yellow)
                self.table.item(ix.row(), ix.column()).setText("y")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.yellow)
                yellow = 5
            elif i_blk == 2:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.red)
                self.table.item(ix.row(), ix.column()).setText("r")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.red)
                red = 5
            elif i_blk == 3:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkGreen)
                self.table.item(ix.row(), ix.column()).setText("g")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkGreen)
                green = 5
            elif i_blk == 4:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkBlue)
                self.table.item(ix.row(), ix.column()).setText("b")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkBlue)
                blue = 5
            else:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkGray)
                self.table.item(ix.row(), ix.column()).setText("d")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkGray)
                gray = 5

    @pyqtSlot()
    def btn_trash(self):
        for ix in self.table.selectedIndexes():
            # print('s r:{0},c:{1}'.format(ix.row(),ix.column()))
            self.table.item(ix.row(), ix.column()).setBackground(Qt.white)
            self.table.item(ix.row(), ix.column()).setText("")

    @pyqtSlot()
    def btn_clear(self):
        # self.table.clear()
        # 색상 변경 위한 item 추가
        row_count = self.table.rowCount()
        col_count = self.table.columnCount()
        for i in range(row_count):
            for j in range(col_count):
                self.table.item(i, j).setBackground(Qt.white)
                self.table.item(i, j).setText("")
                # self.table.setItem(i, j, QTableWidgetItem())#

    @pyqtSlot()
    def btn_addcol(self):
        global temp_count_len
        temp_count_len += 1
        row_count = self.table.rowCount()
        col_count = self.table.columnCount()
        self.table.insertColumn(col_count)  # 새로운 행 count
        # 셀 색상 변경 위해 item 추가
        for i in range(row_count):
            self.table.setItem(i, col_count, QTableWidgetItem())

    @pyqtSlot()
    def btn_addrow(self):
        global temp_count_wid
        temp_count_wid += 1
        row_count = self.table.rowCount()
        col_count = self.table.columnCount()
        self.table.insertRow(row_count)
        # 셀 색상 변경 위해 item 추가
        for j in range(col_count):
            self.table.setItem(row_count, j, QTableWidgetItem())

    @pyqtSlot()
    def btn_delcol(self):
        global temp_count_len
        temp_count_len -= 1
        col_count = self.table.columnCount()
        self.table.removeColumn(col_count - 1)

    @pyqtSlot()
    def btn_delrow(self):
        global temp_count_wid
        temp_count_wid -= 1
        row_count = self.table.rowCount()
        self.table.removeRow(row_count - 1)

    # -saveMap button 함수: 맵 저장
    def btn_save_map(self):
        workbook = xlsxwriter.Workbook(filename)  # 지정 파일 이름
        worksheet1 = workbook.add_worksheet('NewSheet1')
        global yellow, red, blue, gray, green, temp_count_len, temp_count_wid, file_name, file_grid
        sql = "CALL deleteGrid(%s); CALL createGrid('s1', %s, %s, %s, %s, %s); CALL updateCellCnt('s1', %s, %s, %s, %s, %s, %s); CALL updateGridColor('s1', %s, %s, %s, %s, %s, %s);"
        cur.execute(sql, [file_name, file_name, temp_count_len, temp_count_wid, int(file_grid[4]), int(file_grid[5]),
                          file_name, int(file_grid[7]), int(file_grid[8]), int(file_grid[9]), int(file_grid[10]),
                          int(file_grid[11]), file_name, int(file_grid[13]), int(file_grid[14]), int(file_grid[15]),
                          int(file_grid[16]), int(file_grid[17])])

        for i in range(13, 18):
            if int(file_grid[i]) == 1:
                yellow = i - 12
            elif int(file_grid[i]) == 2:
                red = i - 12
            elif int(file_grid[i]) == 3:
                green = i - 12
            elif int(file_grid[i]) == 4:
                blue = i - 12
            else:
                gray = i - 12

        cnum = 1
        CSnum = 1
        CHnum = 1
        WSnum = 1
        BUFnum = 1

        for row in range(self.table.rowCount()):
            # rowData=[]
            for col in range(self.table.columnCount()):
                cell_num = str(file_name) + '_c' + str(cnum).zfill(4)
                cnum += 1
                sql = "CALL createCell('s1', %s, %s, %s, %s);"
                cur.execute(sql, [file_name, cell_num, str(row), str(col)])
                item = self.table.item(row, col)
                # worksheet1.write(row, col, item.text())
                format = workbook.add_format()
                if item.text() == "y":
                    sql = "CALL updateCellStatus('s1', %s, %s);"
                    cur.execute(sql, [cell_num, str(yellow)])
                    format.set_bg_color('yellow')
                    format.set_font_color('yellow')
                    worksheet1.write(row, col, 'y', format)
                if item.text() == "b":
                    sql = "CALL updateCellStatus('s1', %s, %s);"
                    cur.execute(sql, [cell_num, str(blue)])
                    format.set_bg_color('blue')
                    format.set_font_color('blue')
                    worksheet1.write(row, col, 'b', format)
                if item.text() == "g":
                    sql = "CALL updateCellStatus('s1', %s, %s);"
                    cur.execute(sql, [cell_num, str(green)])
                    format.set_bg_color('green')
                    format.set_font_color('green')
                    worksheet1.write(row, col, 'g', format)
                if item.text() == "r":
                    sql = "CALL updateCellStatus('s1', %s, %s);"
                    cur.execute(sql, [cell_num, str(red)])
                    format.set_bg_color('red')
                    format.set_font_color('red')
                    worksheet1.write(row, col, 'r', format)
                if item.text() == "d":
                    sql = "CALL updateCellStatus('s1', %s, %s);"
                    cur.execute(sql, [cell_num, str(gray)])
                    format.set_bg_color('gray')
                    format.set_font_color('gray')
                    worksheet1.write(row, col, 'd', format)
                # if item is not None:
                # rowData.append(item.text())
                # worksheet1.write(row,col,item.text())
                # else:
                # rowData.append('1')
                #   worksheet1.write(row, col, '1')
                sql = "SELECT CellStatus FROM cell " + "WHERE Cell_ID = %s && Simul_ID = 's1'"
                cur.execute(sql, [cell_num])
                Cstatus = cur.fetchone()
                if Cstatus[0] == 1:  # 충전
                    CS_num = str(file_name) + '_CS' + str(CSnum).zfill(4)
                    CSnum += 1
                    sql = "CALL createCS(%s, 's1', %s, %s, NULL);"
                    cur.execute(sql, [file_name, cell_num, CS_num])
                elif Cstatus[0] == 2:  # 슈트
                    CH_num = str(file_name) + '_CH' + str(CHnum).zfill(4)
                    CHnum += 1
                    sql = "CALL createChute(%s, 's1', %s, %s, NULL, NULL);"
                    cur.execute(sql, [file_name, cell_num, CH_num])
                elif Cstatus[0] == 3:  # 워크스테이션
                    WS_num = str(file_name) + '_WS' + str(WSnum).zfill(4)
                    WSnum += 1
                    sql = "CALL createWS(%s, 's1', %s, %s, NULL);"
                    cur.execute(sql, [file_name, cell_num, WS_num])
                elif Cstatus[0] == 4:  # 버퍼
                    BUF_num = str(file_name) + '_BUF' + str(BUFnum).zfill(4)
                    BUFnum += 1
                    sql = "CALL createBuffer(%s, 's1', %s, %s);"
                    cur.execute(sql, [file_name, cell_num, BUF_num])
        cur.execute("CALL deleteProject('p1');")
        workbook.close()
        self.close()


# 2.setGrid.ui
class secondwindow(QDialog, QWidget, form_secondwindow):
    def __init__(self):
        super(secondwindow, self).__init__()
        # self.initUi()
        self.setupUi(self)
        self.setWindowTitle("새 파일 만들기 - 그리드 설정")
        # self.setFixedSize(1000, 800)
        self.show()
        # self.cb1.activated[str].connect(self.onActivated
        self.cb1.activated[str].connect(self.selectone)
        self.cb2.activated[str].connect(self.selecttwo)
        self.cb3.activated[str].connect(self.selectthr)
        self.cb4.activated[str].connect(self.selectfour)
        self.gridNext.clicked.connect(self.btn_next_to_setattribute)  # gridNext button 클릭

    # def onActivated(self, text):
    #    size_len=text
    def selectone(self, text):
        global size_len
        size_len = text
        # print(size_len)

    def selecttwo(self, text):
        global size_wid
        size_wid = text
        # print(size_wid)

    def selectthr(self, text):
        global count_len
        count_len = text
        # print(count_len)

    def selectfour(self, text):
        global count_wid
        count_wid = text
        # print(count_wid)

    # -gridNext button 함수: setAttribute.ui로 창전환, DB저장
    def btn_next_to_setattribute(self, text):
        global size_len, size_wid, count_len, count_wid
        # 입력값 ->DB
        cur.execute("CALL deleteProject('p1');")
        sql = "CALL createProject('p1', NULL, NULL, NULL); CALL createSimul('p1', 's1'); CALL createGrid('s1', 'tempG', %s, %s, %s, %s);"
        self.hide()
        cur.execute(sql, [count_len, count_wid, size_len, size_wid])
        self.second = thirdwindow()
        self.second.exec_()

        # self.show()


# 3.setAttribute.ui
class thirdwindow(QDialog, QWidget, form_thirdwindow):
    def __init__(self):
        super(thirdwindow, self).__init__()
        # self.initUi()
        self.setupUi(self)
        self.setWindowTitle("새 파일 만들기 - 셀 설정")
        # self.setFixedSize(1000, 800)
        # i_charge=0
        global i_charge, i_chute, i_ws, i_buf, i_blk
        i_charge = 0
        i_chute = 0
        i_ws = 0
        i_buf = 0
        i_blk = 0
        self.cb1.activated[str].connect(self.selectone)
        self.cb2.activated[str].connect(self.selecttwo)
        self.cb3.activated[str].connect(self.selectthr)
        self.cb4.activated[str].connect(self.selectfour)
        self.cb5.activated[str].connect(self.selectfive)
        self.btn1.clicked.connect(self.btn_charge_color)
        self.btn2.clicked.connect(self.btn_chute_color)
        self.btn3.clicked.connect(self.btn_ws_color)
        self.btn4.clicked.connect(self.btn_buf_color)
        self.btn5.clicked.connect(self.btn_blk_color)
        self.attributeNext.clicked.connect(self.btn_next_to_map)  # attributeNext button 클릭
        self.show()

    def selectone(self, text):
        global count_charge
        count_charge = text
        # print(count_charge)

    def selecttwo(self, text):
        global count_chute
        count_chute = text
        # print(count_chute)

    def selectthr(self, text):
        global count_ws
        count_ws = text
        # print(count_ws)

    def selectfour(self, text):
        global count_buf
        count_buf = text
        # print(count_buf)

    def selectfive(self, text):
        global count_blk
        count_blk = text
        # print(count_blk)

    # darkGray, red, magenta, green, yellow, blue
    def btn_charge_color(self):
        # 버튼 클릭시 색상 변경 위한 변수(여러 색상)
        global i_charge
        global color_charge
        if i_charge == 5:
            i_charge = 0
        # print(i_charge)
        if i_charge == 0:
            self.btn1.setStyleSheet('background:yellow')
            color_charge = "yellow"
        if i_charge == 1:
            self.btn1.setStyleSheet('background:red')
            color_charge = "red"
        if i_charge == 2:
            self.btn1.setStyleSheet('background:green')
            color_charge = "green"
        if i_charge == 3:
            self.btn1.setStyleSheet('background:blue')
            color_charge = "blue"
        if i_charge == 4:
            self.btn1.setStyleSheet('background:darkGray')
            color_charge = "darkGray"
        i_charge = i_charge + 1

    def btn_chute_color(self):
        # 버튼 클릭시 색상 변경 위한 변수(여러 색상)
        global i_chute
        global color_chute
        if i_chute == 5:
            i_chute = 0
        # print(i_chute)
        if i_chute == 0:
            self.btn2.setStyleSheet('background:yellow')
            color_chute = "yellow"
        if i_chute == 1:
            self.btn2.setStyleSheet('background:red')
            color_chute = "red"
        if i_chute == 2:
            self.btn2.setStyleSheet('background:green')
            color_chute = "green"
        if i_chute == 3:
            self.btn2.setStyleSheet('background:blue')
            color_chute = "blue"
        if i_chute == 4:
            self.btn2.setStyleSheet('background:darkGray')
            color_chute = "darkGray"
        i_chute = i_chute + 1

    def btn_ws_color(self):
        # 버튼 클릭시 색상 변경 위한 변수(여러 색상)
        global i_ws
        global color_ws
        if i_ws == 5:
            i_ws = 0
        # print(i_ws)
        if i_ws == 0:
            self.btn3.setStyleSheet('background:yellow')
            color_ws = "yellow"
        if i_ws == 1:
            self.btn3.setStyleSheet('background:red')
            color_ws = "red"
        if i_ws == 2:
            self.btn3.setStyleSheet('background:green')
            color_ws = "green"
        if i_ws == 3:
            self.btn3.setStyleSheet('background:blue')
            color_ws = "blue"
        if i_ws == 4:
            self.btn3.setStyleSheet('background:darkGray')
            color_ws = "darkGray"
        i_ws = i_ws + 1

    def btn_buf_color(self):
        # 버튼 클릭시 색상 변경 위한 변수(여러 색상)
        global i_buf
        global color_buf
        if i_buf == 5:
            i_buf = 0
        # print(i_buf)
        if i_buf == 0:
            self.btn4.setStyleSheet('background:yellow')
            color_buf = "yellow"
        if i_buf == 1:
            self.btn4.setStyleSheet('background:red')
            color_buf = "red"
        if i_buf == 2:
            self.btn4.setStyleSheet('background:green')
            color_buf = "green"
        if i_buf == 3:
            self.btn4.setStyleSheet('background:blue')
            color_buf = "blue"
        if i_buf == 4:
            self.btn4.setStyleSheet('background:darkGray')
            color_buf = "darkGray"
        i_buf = i_buf + 1

    def btn_blk_color(self):
        # 버튼 클릭시 색상 변경 위한 변수(여러 색상)
        global i_blk
        global color_blk
        if i_blk == 5:
            i_blk = 0
        # print(i_blk)
        if i_blk == 0:
            self.btn5.setStyleSheet('background:yellow')
            color_blk = "yellow"
        if i_blk == 1:
            self.btn5.setStyleSheet('background:red')
            color_blk = "red"
        if i_blk == 2:
            self.btn5.setStyleSheet('background:green')
            color_blk = "green"
        if i_blk == 3:
            self.btn5.setStyleSheet('background:blue')
            color_blk = "blue"
        if i_blk == 4:
            self.btn5.setStyleSheet('background:darkGray')
            color_blk = "darkGray"
        i_blk = i_blk + 1

    # -attributeNext button 함수: createMap.ui로 창전환
    def btn_next_to_map(self):
        global count_charge, count_chute, count_ws, count_buf, count_blk
        global color_charge, color_chute, color_ws, color_buf, color_blk
        # !입력값 -> DB
        sql = "CALL updateCellCnt('s1', 'tempG', %s, %s, %s, %s, %s); CALL updateGridColor('s1', 'tempG', %s, %s, %s, %s, %s)"
        cur.execute(sql,
                    [count_charge, count_chute, count_ws, count_buf, count_blk, i_charge, i_chute, i_ws, i_buf, i_blk])
        self.hide()
        self.third = fourthwindow()
        self.third.exec_()


# 4.createMap.ui
class fourthwindow(QDialog, QWidget, form_fourthwindow):
    def __init__(self, parent=None):
        global temp_count_len, temp_count_wid, count_len, count_wid
        temp_count_len = int(count_len)
        temp_count_wid = int(count_wid)
        super(fourthwindow, self).__init__(parent)
        # self.initUi()
        # self.setupUi(self)
        self.setWindowTitle("새 파일 만들기 - 맵 그리기")
        self.setFixedSize(1000, 900)
        self.table = QTableWidget(parent)
        # self._mainwin=parent
        vbox = QVBoxLayout()
        vbox.addWidget(self.table)
        grid = QGridLayout()
        vbox.addLayout(grid)
        charge = QPushButton("충전")
        grid.addWidget(charge, 0, 0)
        chute = QPushButton("슈트")
        grid.addWidget(chute, 0, 1)
        ws = QPushButton("워크스테이션")
        grid.addWidget(ws, 0, 2)
        buffer = QPushButton("버퍼")
        grid.addWidget(buffer, 1, 0)
        block = QPushButton("블락")
        grid.addWidget(block, 1, 1)
        trash = QPushButton("삭제")
        grid.addWidget(trash, 0, 5)
        clear = QPushButton("초기화")
        grid.addWidget(clear, 1, 5)
        addrow = QPushButton("row추가")
        grid.addWidget(addrow, 0, 3)
        addcol = QPushButton("col추가")
        grid.addWidget(addcol, 0, 4)
        delrow = QPushButton("row삭제")
        grid.addWidget(delrow, 1, 3)
        delcol = QPushButton("col삭제")
        grid.addWidget(delcol, 1, 4)
        save = QPushButton("저장")
        grid.addWidget(save, 2, 6)
        self.setLayout(vbox)
        self.setGeometry(200, 200, 400, 500)
        # self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        # self.table.setSelectionBehavior(QAbstractItemView.SelectItems)
        # self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)  # edit 금지 모드,default
        # self.table.setEditTriggers(QAbstractItemView.AllEditTriggers)  # e
        self.table.setColumnCount(int(count_len))
        self.table.setRowCount(int(count_wid))
        # 반드시 item 생성해야 셀 색상 변경가능
        for i in range(int(count_wid)):
            for j in range(int(count_len)):
                self.table.setItem(i, j, QTableWidgetItem())
        self.table.resizeColumnsToContents()
        self.table.resizeRowsToContents()
        # self.table.selectionModel().selectionChanged.connect(self.on_selection)
        charge.clicked.connect(self.btn_charge)  # charge button 클릭
        chute.clicked.connect(self.btn_chute)  # chute button 클릭
        ws.clicked.connect(self.btn_ws)  # ws button 클릭
        buffer.clicked.connect(self.btn_buffer)  # buffer button 클릭
        block.clicked.connect(self.btn_block)  # block button 클릭
        trash.clicked.connect(self.btn_trash)  # trash button 클릭
        clear.clicked.connect(self.btn_clear)  # clear button 클릭
        addrow.clicked.connect(self.btn_addrow)  # addRow button 클릭
        addcol.clicked.connect(self.btn_addcol)  # addCol button 클릭
        delrow.clicked.connect(self.btn_delrow)  # delRow button 클릭
        delcol.clicked.connect(self.btn_delcol)  # delCol button 클릭
        save.clicked.connect(self.btn_save_map)  # saveMap button 클릭
        self.show()

    # def on_selection(self,selected):
    #    for ix in selected.indexes():
    #        print('select row: {0}, col: {1}'.format(ix.row(),ix.column()))
    @pyqtSlot()
    def btn_charge(self):
        global yellow, red, green, blue, gray
        for ix in self.table.selectedIndexes():
            # print('s r:{0},c:{1}'.format(ix.row(),ix.column()))
            if i_charge == 1:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.yellow)
                self.table.item(ix.row(), ix.column()).setText("y")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.yellow)
                yellow = 1
            elif i_charge == 2:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.red)
                self.table.item(ix.row(), ix.column()).setText("r")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.red)
                red = 1
            elif i_charge == 3:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkGreen)
                self.table.item(ix.row(), ix.column()).setText("g")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkGreen)
                green = 1
            elif i_charge == 4:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkBlue)
                self.table.item(ix.row(), ix.column()).setText("b")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkBlue)
                blue = 1
            else:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkGray)
                self.table.item(ix.row(), ix.column()).setText("d")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkGray)
                gray = 1

    @pyqtSlot()
    def btn_chute(self):
        global yellow, red, green, blue, gray
        for ix in self.table.selectedIndexes():
            # print('s r:{0},c:{1}'.format(ix.row(),ix.column()))
            if i_chute == 1:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.yellow)
                self.table.item(ix.row(), ix.column()).setText("y")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.yellow)
                yellow = 2
            elif i_chute == 2:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.red)
                self.table.item(ix.row(), ix.column()).setText("r")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.red)
                red = 2
            elif i_chute == 3:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkGreen)
                self.table.item(ix.row(), ix.column()).setText("g")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkGreen)
                green = 2
            elif i_chute == 4:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkBlue)
                self.table.item(ix.row(), ix.column()).setText("b")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkBlue)
                blue = 3
            else:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkGray)
                self.table.item(ix.row(), ix.column()).setText("d")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkGray)
                gray = 2

    @pyqtSlot()
    def btn_ws(self):
        global yellow, red, green, blue, gray
        for ix in self.table.selectedIndexes():
            # print('s r:{0},c:{1}'.format(ix.row(),ix.column()))
            if i_ws == 1:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.yellow)
                self.table.item(ix.row(), ix.column()).setText("y")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.yellow)
                yellow = 3
            elif i_ws == 2:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.red)
                self.table.item(ix.row(), ix.column()).setText("r")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.red)
                red = 3
            elif i_ws == 3:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkGreen)
                self.table.item(ix.row(), ix.column()).setText("g")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkGreen)
                green = 3
            elif i_ws == 4:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkBlue)
                self.table.item(ix.row(), ix.column()).setText("b")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkBlue)
                blue = 3
            else:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkGray)
                self.table.item(ix.row(), ix.column()).setText("d")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkGray)
                gray = 3

    @pyqtSlot()
    def btn_buffer(self):
        global yellow, red, green, blue, gray
        for ix in self.table.selectedIndexes():
            # print('s r:{0},c:{1}'.format(ix.row(),ix.column()))
            if i_buf == 1:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.yellow)
                self.table.item(ix.row(), ix.column()).setText("y")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.yellow)
                yellow = 4
            elif i_buf == 2:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.red)
                self.table.item(ix.row(), ix.column()).setText("r")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.red)
                red = 4
            elif i_buf == 3:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkGreen)
                self.table.item(ix.row(), ix.column()).setText("g")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkGreen)
                green = 4
            elif i_buf == 4:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkBlue)
                self.table.item(ix.row(), ix.column()).setText("b")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkBlue)
                blue = 4
            else:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkGray)
                self.table.item(ix.row(), ix.column()).setText("d")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkGray)
                gray = 4

    @pyqtSlot()
    def btn_block(self):
        global yellow, red, green, blue, gray
        for ix in self.table.selectedIndexes():
            # print('s r:{0},c:{1}'.format(ix.row(),ix.column()))
            if i_blk == 1:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.yellow)
                self.table.item(ix.row(), ix.column()).setText("y")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.yellow)
                yellow = 5
            elif i_blk == 2:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.red)
                self.table.item(ix.row(), ix.column()).setText("r")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.red)
                red = 5
            elif i_blk == 3:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkGreen)
                self.table.item(ix.row(), ix.column()).setText("g")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkGreen)
                green = 5
            elif i_blk == 4:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkBlue)
                self.table.item(ix.row(), ix.column()).setText("b")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkBlue)
                blue = 5
            else:
                self.table.item(ix.row(), ix.column()).setBackground(Qt.darkGray)
                self.table.item(ix.row(), ix.column()).setText("d")
                self.table.item(ix.row(), ix.column()).setForeground(Qt.darkGray)
                gray = 5

    @pyqtSlot()
    def btn_trash(self):
        for ix in self.table.selectedIndexes():
            # print('s r:{0},c:{1}'.format(ix.row(),ix.column()))
            self.table.item(ix.row(), ix.column()).setBackground(Qt.white)
            self.table.item(ix.row(), ix.column()).setText("")

    @pyqtSlot()
    def btn_clear(self):
        # self.table.clear()
        # 색상 변경 위한 item 추가
        row_count = self.table.rowCount()
        col_count = self.table.columnCount()
        for i in range(row_count):
            for j in range(col_count):
                self.table.item(i, j).setBackground(Qt.white)
                self.table.item(i, j).setText("")
                # self.table.setItem(i, j, QTableWidgetItem())#

    @pyqtSlot()
    def btn_addcol(self):
        global temp_count_len
        temp_count_len += 1
        row_count = self.table.rowCount()
        col_count = self.table.columnCount()
        self.table.insertColumn(col_count)  # 새로운 행 count
        # 셀 색상 변경 위해 item 추가
        for i in range(row_count):
            self.table.setItem(i, col_count, QTableWidgetItem())

    @pyqtSlot()
    def btn_addrow(self):
        global temp_count_wid
        temp_count_wid += 1
        row_count = self.table.rowCount()
        col_count = self.table.columnCount()
        self.table.insertRow(row_count)
        # 셀 색상 변경 위해 item 추가
        for j in range(col_count):
            self.table.setItem(row_count, j, QTableWidgetItem())

    @pyqtSlot()
    def btn_delcol(self):
        global temp_count_len
        temp_count_len -= 1
        col_count = self.table.columnCount()
        self.table.removeColumn(col_count - 1)

    @pyqtSlot()
    def btn_delrow(self):
        global temp_count_wid
        temp_count_wid -= 1
        row_count = self.table.rowCount()
        self.table.removeRow(row_count - 1)

    # -saveMap button 함수: 맵 저장
    def btn_save_map(self):
        global yellow, red, blue, gray, green, temp_count_len, temp_count_wid
        sql = "CALL updateGridSize('s1', 'tempG', %s, %s);"
        cur.execute(sql, [temp_count_len, temp_count_wid])
        # 저장할 파일 이름 지정위해 추가
        file = QFileDialog.getSaveFileName(self, '', '', 'xlsx Files(*.xlsx)')
        workbook = xlsxwriter.Workbook(file[0])  # 지정 파일 이름
        worksheet1 = workbook.add_worksheet('NewSheet1')
        # worksheet1.write
        cnum = 1
        CSnum = 1
        CHnum = 1
        WSnum = 1
        BUFnum = 1

        for row in range(self.table.rowCount()):
            # rowData=[]
            for col in range(self.table.columnCount()):
                cell_num = 'c' + str(cnum).zfill(4)
                cnum += 1
                sql = "CALL createCell('s1', 'tempG', %s, %s, %s);"
                cur.execute(sql, [cell_num, str(row), str(col)])
                item = self.table.item(row, col)
                # worksheet1.write(row, col, item.text())
                format = workbook.add_format()
                if item.text() == "y":
                    sql = "CALL updateCellStatus('s1', %s, %s);"
                    cur.execute(sql, [cell_num, str(yellow)])
                    format.set_bg_color('yellow')
                    format.set_font_color('yellow')
                    worksheet1.write(row, col, 'y', format)
                if item.text() == "b":
                    sql = "CALL updateCellStatus('s1', %s, %s);"
                    cur.execute(sql, [cell_num, str(blue)])
                    format.set_bg_color('blue')
                    format.set_font_color('blue')
                    worksheet1.write(row, col, 'b', format)
                if item.text() == "g":
                    sql = "CALL updateCellStatus('s1', %s, %s);"
                    cur.execute(sql, [cell_num, str(green)])
                    format.set_bg_color('green')
                    format.set_font_color('green')
                    worksheet1.write(row, col, 'g', format)
                if item.text() == "r":
                    sql = "CALL updateCellStatus('s1', %s, %s);"
                    cur.execute(sql, [cell_num, str(red)])
                    format.set_bg_color('red')
                    format.set_font_color('red')
                    worksheet1.write(row, col, 'r', format)
                if item.text() == "d":
                    sql = "CALL updateCellStatus('s1', %s, %s);"
                    cur.execute(sql, [cell_num, str(gray)])
                    format.set_bg_color('gray')
                    format.set_font_color('gray')
                    worksheet1.write(row, col, 'd', format)
                # if item is not None:
                # rowData.append(item.text())
                # worksheet1.write(row,col,item.text())
                # else:
                # rowData.append('1')
                #   worksheet1.write(row, col, '1')
                sql = "SELECT CellStatus FROM cell " + "WHERE Cell_ID = %s && Simul_ID = 's1';"
                cur.execute(sql, [cell_num])
                Cstatus = cur.fetchone()
                if Cstatus[0] == 1:  # 충전
                    CS_num = 'CS' + str(CSnum).zfill(4)
                    CSnum += 1
                    sql = "CALL createCS('tempG', 's1', %s, %s, NULL);"
                    cur.execute(sql, [cell_num, CS_num])
                elif Cstatus[0] == 2:  # 슈트
                    CH_num = 'CH' + str(CHnum).zfill(4)
                    CHnum += 1
                    sql = "CALL createChute('tempG', 's1', %s, %s, NULL, NULL);"
                    cur.execute(sql, [cell_num, CH_num])
                elif Cstatus[0] == 3:  # 워크스테이션
                    WS_num = 'WS' + str(WSnum).zfill(4)
                    WSnum += 1
                    sql = "CALL createWS('tempG', 's1', %s, %s, NULL);"
                    cur.execute(sql, [cell_num, WS_num])
                elif Cstatus[0] == 4:  # 버퍼
                    BUF_num = 'BUF' + str(BUFnum).zfill(4)
                    BUFnum += 1
                    sql = "CALL createBuffer('tempG', 's1', %s, %s);"
                    cur.execute(sql, [cell_num, BUF_num])

            # worksheet1.write_row(row,rowData)
        # 그리드 ID를 파일 이름으로 변경하기
        file_name = QFileInfo(file[0]).baseName()
        sql = "CALL updateGridName('s1', %s);"
        cur.execute(sql, [str(file_name)])

        # 셀 및 특수 셀 ID에 파일 이름을 연결짓기
        for i in range(1, cnum):
            cname = str(file_name) + '_c' + str(i).zfill(4)
            oldcname = 'c' + str(i).zfill(4)
            cur.execute("CALL updateCellName(%s, %s);", [oldcname, cname])
        for i in range(1, CSnum):
            CSname = str(file_name) + '_CS' + str(i).zfill(4)
            oldCSname = 'CS' + str(i).zfill(4)
            cur.execute("CALL updateCSName(%s, %s)", [oldCSname, CSname])
        for i in range(1, CHnum):
            CHname = str(file_name) + '_CH' + str(i).zfill(4)
            oldCHname = 'CH' + str(i).zfill(4)
            cur.execute("CALL updateCHName(%s, %s)", [oldCHname, CHname])
        for i in range(1, WSnum):
            WSname = str(file_name) + '_WS' + str(i).zfill(4)
            oldWSname = 'WS' + str(i).zfill(4)
            cur.execute("CALL updateWSName(%s, %s)", [oldWSname, WSname])
        for i in range(1, BUFnum):
            BUFname = str(file_name) + '_BUF' + str(i).zfill(4)
            oldBUFname = 'BUF' + str(i).zfill(4)
            cur.execute("CALL updateBufferName(%s, %s)", [oldBUFname, BUFname])

        cur.execute("CALL deleteProject('p1');")
        workbook.close()
        self.close()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle(QStyleFactory.create('Fusion'))
    myWindow = WindowClass()
    # fwin=fourthwindow()
    myWindow.show()
    # fwin.show()
    app.exec_()

conn.close()
