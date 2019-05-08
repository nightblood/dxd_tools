# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'mainwindow.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
# from PyQt5.Qt import QThread
from PyQt5.QtCore import *
from autofill_engine import Engine


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 558)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(MainWindow.sizePolicy().hasHeightForWidth())
        MainWindow.setSizePolicy(sizePolicy)
        MainWindow.setLayoutDirection(QtCore.Qt.LeftToRight)
        MainWindow.setAutoFillBackground(True)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayoutWidget_5 = QtWidgets.QWidget(self.centralwidget)
        self.verticalLayoutWidget_5.setGeometry(QtCore.QRect(0, 0, 791, 581))
        self.verticalLayoutWidget_5.setObjectName("verticalLayoutWidget_5")
        self.verticalLayout_5 = QtWidgets.QVBoxLayout(self.verticalLayoutWidget_5)
        self.verticalLayout_5.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_5.setObjectName("verticalLayout_5")
        self.label = QtWidgets.QLabel(self.verticalLayoutWidget_5)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label.sizePolicy().hasHeightForWidth())
        self.label.setSizePolicy(sizePolicy)
        self.label.setText("")
        self.label.setPixmap(QtGui.QPixmap("./res/logo.png"))
        self.label.setObjectName("label")
        self.verticalLayout_5.addWidget(self.label)
        self.frame = QtWidgets.QFrame(self.verticalLayoutWidget_5)
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.verticalLayoutWidget_2 = QtWidgets.QWidget(self.frame)
        self.verticalLayoutWidget_2.setGeometry(QtCore.QRect(-10, 0, 799, 221))
        self.verticalLayoutWidget_2.setObjectName("verticalLayoutWidget_2")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.verticalLayoutWidget_2)
        self.verticalLayout_2.setSizeConstraint(QtWidgets.QLayout.SetMinimumSize)
        self.verticalLayout_2.setContentsMargins(10, 10, 10, 10)
        self.verticalLayout_2.setSpacing(10)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.label_2 = QtWidgets.QLabel(self.verticalLayoutWidget_2)
        self.label_2.setStyleSheet("font: 75 22pt \"Aharoni\";")
        self.label_2.setObjectName("label_2")
        self.verticalLayout_2.addWidget(self.label_2, 0, QtCore.Qt.AlignHCenter)
        self.pushButton = QtWidgets.QPushButton(self.verticalLayoutWidget_2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton.sizePolicy().hasHeightForWidth())
        self.pushButton.setSizePolicy(sizePolicy)
        self.pushButton.setStyleSheet("font: 75 18pt \"Aharoni\";")
        self.pushButton.setObjectName("pushButton")
        self.verticalLayout_2.addWidget(self.pushButton, 0, QtCore.Qt.AlignHCenter)
        self.label_3 = QtWidgets.QLabel(self.verticalLayoutWidget_2)
        self.label_3.setObjectName("label_3")
        self.verticalLayout_2.addWidget(self.label_3, 0, QtCore.Qt.AlignHCenter)
        self.verticalLayout_5.addWidget(self.frame)
        self.frame_2 = QtWidgets.QFrame(self.verticalLayoutWidget_5)
        self.frame_2.setObjectName("frame_2")
        self.verticalLayoutWidget_3 = QtWidgets.QWidget(self.frame_2)
        self.verticalLayoutWidget_3.setGeometry(QtCore.QRect(110, 10, 451, 191))
        self.verticalLayoutWidget_3.setObjectName("verticalLayoutWidget_3")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout(self.verticalLayoutWidget_3)
        self.verticalLayout_3.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.label_4 = QtWidgets.QLabel(self.verticalLayoutWidget_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_4.sizePolicy().hasHeightForWidth())
        self.label_4.setSizePolicy(sizePolicy)
        self.label_4.setStyleSheet("font: 75 14pt \"Aharoni\";")
        self.label_4.setObjectName("label_4")
        self.verticalLayout_3.addWidget(self.label_4)
        self.radioButton_4 = QtWidgets.QRadioButton(self.verticalLayoutWidget_3)
        self.radioButton_4.setStyleSheet("\n"
"font: 75 14pt \"Aharoni\";")
        self.radioButton_4.setObjectName("radioButton_4")
        self.verticalLayout_3.addWidget(self.radioButton_4)
        self.radioButton_7 = QtWidgets.QRadioButton(self.verticalLayoutWidget_3)
        self.radioButton_7.setStyleSheet("\n"
"font: 75 14pt \"Aharoni\";")
        self.radioButton_7.setObjectName("radioButton_7")
        self.verticalLayout_3.addWidget(self.radioButton_7)
        self.radioButton_5 = QtWidgets.QRadioButton(self.verticalLayoutWidget_3)
        self.radioButton_5.setStyleSheet("font: 75 14pt \"Aharoni\";")
        self.radioButton_5.setObjectName("radioButton_5")
        self.verticalLayout_3.addWidget(self.radioButton_5)
        self.pushButton_2 = QtWidgets.QPushButton(self.verticalLayoutWidget_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_2.sizePolicy().hasHeightForWidth())
        self.pushButton_2.setSizePolicy(sizePolicy)
        self.pushButton_2.setStyleSheet("font: 75 12pt \"Aharoni\";")
        self.pushButton_2.setObjectName("pushButton_2")
        self.verticalLayout_3.addWidget(self.pushButton_2, 0, QtCore.Qt.AlignHCenter)
        self.verticalLayout_5.addWidget(self.frame_2)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 17))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        self.pushButton.clicked.connect(self.on_click_enter)
        self.pushButton_2.clicked.connect(self.on_click_exe)
        self.bg1 = QtWidgets.QButtonGroup(MainWindow)
        self.bg1.addButton(self.radioButton_4, 0)
        self.bg1.addButton(self.radioButton_5, 2)
        self.bg1.addButton(self.radioButton_7, 1)

        self.frame.show()
        self.frame_2.hide()
        self.mainwindow = MainWindow
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "大信贷自动填单系统"))
        self.label_2.setText(_translate("MainWindow", "大信贷自动填单系统"))
        self.pushButton.setText(_translate("MainWindow", "进入系统"))
        self.label_3.setText(_translate("MainWindow", "提示：成功登录后再点击【进入系统】"))
        self.label_4.setText(_translate("MainWindow", "功能列表："))
        self.radioButton_4.setText(_translate("MainWindow", "1. 财报录入"))
        self.radioButton_7.setText(_translate("MainWindow", "2. 新增普惠金融客户"))
        self.radioButton_5.setText(_translate("MainWindow", "3. 新增对公信贷客户"))
        self.pushButton_2.setText(_translate("MainWindow", "执行"))

    def on_click_enter(self):
        print('on_click_enter()....')
        try:
            # self.pushButton.setEnabled(False)
            self.thread = WorkThread(self.engine)
            self.thread.sig_enter_complite.connect(self.show_func)
            self.thread.finished.connect(self.log)
            self.thread.start()
        except Exception as e:
            print(e)
            # self.pushButton.setEnabled(True)

    def log(self):
        print('log()...')

    def show_func(self):
        print('show_func()...')
        self.frame_2.show()
        self.frame.hide()
        self.pushButton_2.setEnabled(True)

    def on_click_exe(self):
        print('on_click_exe()...')
        try:
            self.pushButton_2.setEnabled(False)

            self.thread = WorkThread(self.engine, self.bg1.checkedId())
            self.thread.sig_exe_complite.connect(self.hide_mask)
            self.thread.sig_show_alert.connect(self.alert_dialog)
            self.thread.start()
        except Exception as e:
            self.alert_dialog(e, title='错误')
            print(e)
            self.pushButton_2.setEnabled(True)

    def hide_mask(self):
        self.pushButton_2.setEnabled(True)

    def set_engine(self, engine):
        self.engine = engine
        engine.window = self

        # thread = WorkThread(self.engine, -2)
        # thread.start()
        # engine.alert_dialog = (self.alert_dialog)
        # engine.confirm_dialog = (self.confirm_dialog)

    def alert_dialog(self, msg, title='提示'):
        QtWidgets.QMessageBox.warning(self.mainwindow,
                                      title,
                                      msg,
                                      QtWidgets.QMessageBox.Yes)

    def confirm_dialog(self, msg, pos_callback, neg_callback, title='提示'):
        try:
            reply = QtWidgets.QMessageBox.question(self.mainwindow,
                                         title,
                                         msg,
                                         QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            if reply == QtWidgets.QMessageBox.Yes:
                pos_callback()
            elif reply == QtWidgets.QMessageBox.No:
                neg_callback()
        except Exception as e:
            print(e)

    def raw_confirm_dialog(self, msg, title='提示'):
        return QtWidgets.QMessageBox.question(self.mainwindow,
                                     title,
                                     msg,
                                     QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)

class WorkThread(QThread):
    sig_enter_complite = pyqtSignal(str)
    sig_exe_complite = pyqtSignal(str)
    sig_show_alert = pyqtSignal(str)

    def __init__(self, engine, func_id=-1):
        super().__init__()
        self.engine = engine
        self.func_id = func_id
        print('WorkThread', self)

    # def __del__(self):
    #     print('WorkThread', self)
    #     self.wait()

    def thread_click_enter(self):
        try:
            self.engine.click_enter()
            print('sig_enter_complite emit...')
            self.sig_enter_complite.emit('')
        except Exception as e:
            print(e)

    def thread_click_exe(self, index):
        try:
            res = self.engine.click_exe_func(index)
            print('click_exe', res)
            self.sig_exe_complite.emit('')
            if res.get('code') != 0:
                self.sig_show_alert.emit(res.get('msg'))
            else:
                self.sig_show_alert.emit('完成自动填单！')

        except Exception as e:
            print(e)

    def run(self):
        print('Thread running...')
        if self.func_id == -1:
            self.thread_click_enter()
        elif self.func_id == -2:
            self.engine.init_engine()
        else:
            self.thread_click_exe(self.func_id)
