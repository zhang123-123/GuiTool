# -*- coding:utf-8 -*-
"""
@Project Name : gui_tool
@File Name    : run.py
@Programmer   : XiaoPang
@Start Date   : 2021/3/8 11:11
@File Info    : 
"""
import sys
from tool import Ui_Form
from tool_oper import FileConversion
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QFileDialog


class MWindow(QWidget, Ui_Form):
    def __init__(self):
        super().__init__()
        self.file_conver = FileConversion()
        self.path = "D://"
        self.setupUi(self)
        # self.word_file.clicked.connect(self.file_conver.open_file(["doc", "docx"]))
        self.prepare()

    def prepare(self):
        """准备运行"""
        # self.word_file.clicked.connect(lambda: self.file_conver.open_file(["doc", "docx"]))
        # word --> pdf
        self.word_file.clicked.connect(lambda: self.open_file(["doc", "docx"]))
        self.pdf_save_path.clicked.connect(lambda: self.choose_path())
        self.word_to_pdf_Button.clicked.connect(lambda: self.file_conver.word_to_pdf())
        # pdf --> word
        self.pdf_file.clicked.connect(lambda: self.open_file(["pdf"]))
        self.word_save_path.clicked.connect(lambda: self.choose_path())
        self.pdf_to_word_Button.clicked.connect((lambda: self.file_conver.pdf_to_word()))

    def open_file(self, file_types=None):
        """
        打开文件，显示文件名
        :param file_types:
        :return:
        """
        self.file_conver.open_file(file_types)
        self.word_file_name.setText(self.file_conver.file_name)

    def choose_path(self):
        """
        选择路径，并打印路径
        :return:
        """
        self.file_conver.choose_path()
        self.pdf_save_path_text.setText(self.file_conver.fold_path)

    def file_con(self):
        """
        文件转换
        :return:
        """
        pass

    def process(self):
        pass


def main():
    app = QApplication(sys.argv)
    m = MWindow()
    m.show()
    # m.process()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
