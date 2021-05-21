# -*- coding:utf-8 -*-
"""
@Project Name : gui_tool
@File Name    : tool_oper.py
@Programmer   : XiaoPang
@Start Date   : 2021/3/8 14:32
@File Info    : 工具操作逻辑
"""
import os
import sys
import time
import comtypes.client
from tools.pdf_to_word import PdfToWord
from PyQt5.QtWidgets import QFileDialog, QWidget, QMessageBox


class FileConversion(object):
    """
    文件转换
    """

    def __init__(self):
        # 默认路径
        self.path = "D://"
        self.file_name = None
        self.file_path = None
        self.fold_path = None
        # self.ui = Ui_Form

    def open_file(self, file_types=None):
        type_list = []
        for type_ in file_types:
            type_list.append("Text Files (*.{})".format(type_))
        f = QFileDialog.getOpenFileName(None, "请选择要添加的文件", self.path, "{};;All Files (*)".format(";;".join(type_list)))
        if f[0]:
            self.file_path = f[0]
            self.file_name = f[0].split("/")[-1]
        else:
            print("添加文件错误")

    def choose_path(self):
        p = QFileDialog.getExistingDirectory(None, "请选择文件夹路径", self.path)
        if p != "":
            self.fold_path = p

    def word_to_pdf(self):
        """
        word 转 pdf
        :return:
        """
        result = QMessageBox.information(None, "提示框", "是否开始转换", QMessageBox.Yes | QMessageBox.No)
        # print(f"QMessageBox.Yes：{QMessageBox.Yes}, {type(QMessageBox.Yes)}")
        # print(f"QMessageBox.No：{QMessageBox.No}, {type(QMessageBox.No)}")
        if result != QMessageBox.Yes:
            return
        file_name, ext = None, None
        try:
            file_name, ext = self.file_name.split(".")
        except ValueError as e:
            print("异常：{}".format(e))
        except Exception as e:
            print("未知错误：{}".format(e))
        if ext not in ("doc", "docx"):
            return
        word_path = "{}".format(self.file_path)
        pdf_path = r"{}/{}.pdf".format(self.fold_path, file_name)
        word_path = "\\".join(word_path.split("/"))
        pdf_path = "\\".join(pdf_path.split("/"))
        start_ = time.time()
        word = comtypes.client.CreateObject("Word.Application")
        word.Visible = 0
        new_pdf = word.Documents.Open(word_path)
        new_pdf.SaveAs(pdf_path, FileFormat=17)
        new_pdf.Close()
        time_ = "{:.2f}".format(time.time() - start_)
        print("用时：{}s".format(time_))
        QMessageBox.information(None, "完成框", "转换完成，用时{}s.".format(time_))

    def pdf_to_word(self):
        result = QMessageBox.information(None, "提示框", "是否开始转换", QMessageBox.Yes | QMessageBox.No)
        # print(f"QMessageBox.Yes：{QMessageBox.Yes}, {type(QMessageBox.Yes)}")
        # print(f"QMessageBox.No：{QMessageBox.No}, {type(QMessageBox.No)}")
        if result != QMessageBox.Yes:
            return
        file_name, ext = None, None
        try:
            file_name, ext = self.file_name.split(".")
        except ValueError as e:
            print("异常：{}".format(e))
        except Exception as e:
            print("未知错误：{}".format(e))
        if ext not in ("pdf",):
            return

        pdf_path = "{}".format(self.file_path)
        word_path = r"{}/{}.doc".format(self.fold_path, file_name)
        print(pdf_path, word_path)
        pdf_path = "\\".join(pdf_path.split("/"))
        word_path = "\\".join(word_path.split("/"))
        print(pdf_path, word_path)



