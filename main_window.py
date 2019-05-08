# -*- coding: utf-8 -*-

import sys
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import excel_util
import os
from autofill_engine import Engine
from mask_layout import MaskWidget
import time


class MainWindow(QTabWidget):
    def __init__(self, engine, parent=None):
        super(MainWindow, self).__init__(parent)
        # MainWindow.resize(self, 760, 500)
        self.mainwindow = MainWindow
        MainWindow.setFixedSize(self, 760, 500)

        self.data = ''
        self.line_idx_0 = 0
        self.engine = engine
        self.mask = ''

        self.tab1 = QWidget()
        self.tab2 = QWidget()
        self.tab3 = QWidget()
        self.tab4 = QWidget()

        self.addTab(self.tab1, "Tab 1")
        self.addTab(self.tab2, "Tab 2")
        self.addTab(self.tab3, "Tab 3")

        self.tab1UI()
        self.tab2UI()
        self.tab3UI()

        self.setWindowTitle("大信贷自动填单")
        self.setGeometry(500, 500, 500, 500)

        # try:
        #     reply = QMessageBox.question(self,
        #                                  '提示',
        #                                  '点击yes 进入程序。所有功能都将在该应用打开的浏览器内操作，否则将无效。',
        #                                  QMessageBox.No | QMessageBox.Yes)
        #     if reply == QMessageBox.Yes:
        #         self.worker_engine = WorkerEngine()
        #         self.worker_engine.sig_engine_success.connect(self.set_engine)
        #         self.worker_engine.sig_engine_fail.connect(self.init_engine_fail)
        #         self.worker_engine.start()
        #     else:
        #         exit(1)
        # except Exception as e:
        #     print(e)


    def set_engine(self, msg):
        print('init engine success...' + msg)
        self.engine = Engine.get_instance()

    def init_engine_fail(self, msg):
        print(msg)

    def tab1UI(self):
        mainsp = QSplitter(Qt.Horizontal)
        leftsp = QSplitter(Qt.Vertical)
        rightsp = QSplitter(Qt.Vertical)

        mainsp.setFrameShape(QFrame.StyledPanel)
        # leftsp.setFrameShape(QFrame.StyledPanel)
        # rightsp.setFrameShape(QFrame.StyledPanel)

        desclabel = QLabel('提示：\n手动定位页面，\n再点击按钮进行\n自动填单。')
        btn0 = QPushButton('全自动填单')
        btn1 = QPushButton('资产负债表')
        btn2 = QPushButton('损益表')
        btn3 = QPushButton('现金流量表')
        btn4 = QPushButton('现金流量表附表')
        btn5 = QPushButton('生成财报模板')
        btn6 = QPushButton('导入财报')
        self.infoline_0 = QPlainTextEdit('财报自动填单\n这里是说明。。。')

        leftsp.addWidget(btn0)
        leftsp.addWidget(desclabel)
        leftsp.addWidget(btn1)
        leftsp.addWidget(btn2)
        leftsp.addWidget(btn3)
        leftsp.addWidget(btn4)
        leftsp.addWidget(btn5)
        leftsp.addWidget(btn6)
        rightsp.addWidget(self.infoline_0)

        btn1.setMaximumSize(250, 60)
        btn2.setMaximumSize(250, 60)
        btn3.setMaximumSize(250, 60)
        btn4.setMaximumSize(250, 60)
        btn5.setMaximumSize(250, 60)
        btn6.setMaximumSize(250, 60)

        btn0.setFont(QFont("Roman times", 10, QFont.Bold))
        btn1.setFont(QFont("Roman times", 10, QFont.Bold))
        btn2.setFont(QFont("Roman times", 10, QFont.Bold))
        btn3.setFont(QFont("Roman times", 10, QFont.Bold))
        btn4.setFont(QFont("Roman times", 10, QFont.Bold))
        btn5.setFont(QFont("Roman times", 10, QFont.Bold))
        btn6.setFont(QFont("Roman times", 10, QFont.Bold))

        icon_report = QIcon()
        icon_report.addPixmap(QPixmap("./res/icon_report.png"), QIcon.Normal, QIcon.Off)
        btn0.setIcon(icon_report)
        btn1.setIcon(icon_report)
        btn2.setIcon(icon_report)
        btn3.setIcon(icon_report)
        btn4.setIcon(icon_report)

        icon_import = QIcon()
        icon_import.addPixmap(QPixmap("./res/icon_import.png"), QIcon.Normal, QIcon.Off)
        btn6.setIcon(icon_import)

        icon_excel = QIcon()
        icon_excel.addPixmap(QPixmap("./res/icon_excel.png"), QIcon.Normal, QIcon.Off)
        btn5.setIcon(icon_excel)

        mainsp.setStretchFactor(0, 1)
        mainsp.setStretchFactor(1, 2)
        mainsp.addWidget(leftsp)
        mainsp.addWidget(rightsp)

        mainsp.setContentsMargins(10, 10, 10, 10)
        leftsp.setContentsMargins(20, 20, 20, 20)
        rightsp.setContentsMargins(20, 20, 20, 20)

        self.setTabText(0, "财报")
        layout = QHBoxLayout()
        layout.addWidget(mainsp)

        self.tab1.setLayout(layout)
        btn0.clicked.connect(self.on_click_autofill)
        btn1.clicked.connect(self.on_click_blance_sheet)
        btn2.clicked.connect(self.on_click_income_statement)
        btn3.clicked.connect(self.on_click_flow_statement)
        btn4.clicked.connect(self.on_click_flow_statement1)
        btn5.clicked.connect(self.on_click_templet_excel)
        btn6.clicked.connect(self.on_click_import_excel)

    def alert_dialog(self, msg):
        try:
            QMessageBox.question(self, '提示', msg,  QMessageBox.Yes)
        except Exception as e:
            print(e)

    def confirm_dialog(self, msg, pos_callback, neg_callback, title='提示'):
        try:
            reply = QMessageBox.question(self,
                                         title,
                                         msg,
                                         QMessageBox.No | QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                pos_callback()
            elif reply == QMessageBox.No:
                neg_callback()
        except Exception as e:
            print(e)

    def infoline_log_0(self, msg):
        self.infoline_0.appendPlainText(str(self.line_idx_0) + ': ' + msg)
        self.infoline_0.moveCursor(QTextCursor.End)
        self.line_idx_0 += 1

    def on_click_autofill(self):

        self.worker_engine = WorkerEngine()
        self.worker_engine.sig_engine_success.connect(self.set_engine)
        self.worker_engine.sig_engine_fail.connect(self.init_engine_fail)
        self.worker_engine.start()
        # try:
        #     reply = QMessageBox.question(self,
        #                                  '提示',
        #                                  '全自动的填单将自动定位页面进行填单，如果在定位过程出现异常时也可以通过下面半自动填单完成。点击yes进行填单。',
        #                                  QMessageBox.No | QMessageBox.Yes)
        #     if reply == QMessageBox.Yes:
        #         self.infoline_log_0('点击yes')
        #         pass
        # except Exception as e:
        #     print(e)

    def file_check(self, file):
        if file is None or len(file) == 0:
            return False
        if not file.endswith('.xlsx'):
            return False
        return os.path.exists(file)

    def on_click_import_excel(self):
        try:
            s = QFileDialog.getOpenFileName(self, "财报导入", "/", "Excel File(*.xlsx)")

            if self.file_check(s[0]):
                self.data = excel_util.get_business_report_data_from_excel(s[0])
                self.infoline_log_0('导入成功！！\n')
            else:
                self.infoline_log_0('文件选取异常，请重新选择 xlsx 格式的文件！！')
        except Exception as e:
            self.infoline_log_0('程序异常 %s！！' % e)
            print(e)

    def on_click_blance_sheet(self):

        try:
            if self.data is None or len(self.data) == 0:
                self.infoline_log_0('请先点击【导入财报】按钮，导入财报！！')
                self.alert_dialog('请先点击【导入财报】按钮，导入财报！！')
                return

            reply = QMessageBox.question(self,
                                         '提示',
                                         '请先将页面定位到需要填单的表。点击yes进行填单。',
                                         QMessageBox.No | QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                self.mask = MaskWidget(self)
                self.mask.show()
                self.worker_fill = WorkerFill(self.engine, self.data, 1)
                self.worker_fill.sig_fill_complite.connect(self.fill_complete)
                self.worker_fill.start()
        except Exception as e:
            print(e)
            self.infoline_log_0('程序异常 %s！！' % e)
            if self.mask is not None:
                self.mask.close()

    def fill_complete(self, msg):
        self.infoline_log_0('填单完成！！' + msg)
        if self.mask is not None:
            self.mask.close()

    def on_click_income_statement(self):
        """损益表"""
        try:
            reply = QMessageBox.question(self,
                                         '提示',
                                         '请先将页面定位到需要填单的表。点击yes进行填单。',
                                         QMessageBox.No | QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                self.mask = MaskWidget(self)
                self.mask.show()
                self.worker_fill = WorkerFill(self.engine, self.data, 2)
                self.worker_fill.sig_fill_complite.connect(self.fill_complete)
                self.worker_fill.start()
        except Exception as e:
            print(e)
            if self.mask is not None:
                self.mask.close()
            self.infoline_log_0('程序异常 %s！！' % e)

    def on_click_flow_statement(self):
        """现金流量表"""
        try:
            reply = QMessageBox.question(self,
                                         '提示',
                                         '现金流量表 开发中。',
                                         QMessageBox.No | QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                pass
        except Exception as e:
            print(e)

    def on_click_flow_statement1(self):
        """现金流量附属表"""
        try:
            reply = QMessageBox.question(self,
                                         '提示',
                                         '现金流量附属表 开发中。',
                                         QMessageBox.No | QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                pass
        except Exception as e:
            print(e)

    def on_click_templet_excel(self):
        try:
            file = r'C:\Users\Administrator\Desktop\财报.xlsx'
            if os.path.exists(file):
                reply = QMessageBox.question(self,
                                             '警告',
                                             '桌面已存在文件【财报.xlsx】！！是否覆盖该文件？',
                                             QMessageBox.No | QMessageBox.Yes)
                if reply == QMessageBox.Yes:
                    excel_util.create_excel()
                    self.infoline_log_0('已在桌面生成【财报.xlsx】。')
            else:
                excel_util.create_excel()
                self.infoline_log_0('已在桌面生成【财报.xlsx】。')
                QMessageBox.warning(self, '提示', '已在桌面生成【财报.xlsx】。', QMessageBox.Yes)
        except Exception as e:
            print(e)
            self.infoline_log_0('程序异常 %s！！' % e)

    def tab2UI(self):
        layout = QFormLayout()
        layout.addWidget(QLabel('开发中。。。'))
        self.tab2.setLayout(layout)
        self.setTabText(1, "普惠金融客户")


    def tab3UI(self):
        layout = QFormLayout()
        layout.addWidget(QLabel('开发中。。。'))
        self.tab3.setLayout(layout)
        self.setTabText(2, "普惠金融客户")


class WorkerFill(QThread):
    sig_fill_complite = pyqtSignal(str)

    def __init__(self, engine, data, func_id):
        super(WorkerFill, self).__init__()
        # super.__init__()
        self.engine = engine
        self.data = data
        self.func_id = func_id  # 0.全自动。1.资产负债表。2.损益表。3.流量表。4.流量附属表

    def run(self):
        try:
            if self.func_id == 0:
                self.engine.fill_all(self.data)
            elif self.func_id == 1:
                self.engine.fill_form(self.data[0])
            elif self.func_id == 2:
                self.engine.fill_form(self.data[1])
            elif self.func_id == 3:
                self.engine.fill_form(self.data[2])
            elif self.func_id == 4:
                self.engine.fill_form(self.data[3])
            self.sig_fill_complite.emit('')
        except Exception as e:
            print(e)
            self.sig_fill_complite.emit('异常：' + str(e))

    def do_something(self):
        time.sleep(3)
        raise Exception('i am not good...')


class WorkerEngine(QThread):
    sig_engine_success = pyqtSignal(str)
    sig_engine_fail = pyqtSignal(str)

    def __init__(self):
        super(WorkerEngine, self).__init__()

    def run(self):
        try:
            print('run init engine...')
            engine = Engine.get_instance()
            engine.init_engine()
            self.sig_engine_success.emit('')
        except Exception as e:
            print(e)
            self.sig_engine_fail.emit(str(e))



if __name__ == '__main__':
    app = QApplication(sys.argv)
    demo = MainWindow('')
    demo.show()
    sys.exit(app.exec_())
