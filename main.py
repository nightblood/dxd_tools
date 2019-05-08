# -*- coding: utf-8 -*-

import sys
from PyQt5.QtWidgets import QApplication, QMainWindow
from autofill_engine import Engine
from PyQt5 import QtWidgets, QtGui
import time
from main_window import MainWindow

import subprocess

if __name__ == '__main__':
    cmd = 'your command'
    res = subprocess.call(cmd, shell=True, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    # 对Qt部件的操作一般都要在创建Qt程序后才能进行

    app = QApplication(sys.argv)

    # 创建启动界面，支持png透明图片
    splash = QtWidgets.QSplashScreen(QtGui.QPixmap('./res/logo.png'))
    splash.show()
    # 可以显示启动信息
    # splash.showMessage('正在加载……')
    engine = Engine.get_instance()
    # engine.init_engine()
    # time.sleep(1)
    splash.close()
    # engine.init_engine()
    mainWindow = MainWindow(engine)
    mainWindow.show()
    sys.exit(app.exec_())
