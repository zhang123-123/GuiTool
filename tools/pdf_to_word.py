# -*- coding:utf-8 -*-
"""
@Project Name : gui_tool
@File Name    : pdf_to_word.py
@Programmer   : XiaoPang
@Start Date   : 2021/3/10 11:28
@File Info    : 
"""
import requests
import string
import time
import math
import random
import sys
import os
from requests_toolbelt import MultipartEncoder


class PdfToWord(object):
    """pdf 转 word"""

    def __init__(self, pdf_path, word_path=None):
        self.url = "https://pdf2doc.com"
        self.uid = None
        self.id_ = None
        self.pdf_path = pdf_path
        self.word_path = word_path
        self.prepare()

    def prepare(self):
        self.uid = self.get_uid()
        # self.id_ = self.get_id()
        # id_ ：可以是任意字符串。
        self.id_ = "o_1f0b39rqd9c013o8os1rh313s91"

    @staticmethod
    def __base32(x):
        result = ""
        while x > 0:
            result = string.printable[x % 32] + result
            x //= 32
        return result

    def get_id(self):
        uid = self.__base32(int(time.time() * 1000))  # Python equivalent of new Date().getTime().toString(32)
        for x in range(5):
            uid += self.__base32(int(math.floor(random.random() * 65535)))
        return "o_" + uid + self.__base32(1)

    @staticmethod
    def get_uid():
        chars = "0123456789abcdefghiklmnopqrstuvwxyz"
        result = ""
        for x in range(16):
            char = int(math.floor(random.random() * len(chars)))
            result += chars[char:char + 1]
        return result

    def process(self):
        """
        文件处理：
        1.上传；
        2.转换；
        3.查看转换状态；
        4.下载；
        :return:
        """
        # pdf 文件上传
        self.upload(self.pdf_path)
        # pdf 转换 Word
        self.convert()

    def upload(self, file_path):
        """
        pdf文件上传操作
        :param file_path:
        :return:
        """
        # file_name = file_path.split("/")[-1]
        file_name = "a.pdf"
        f = open(file_path, "rb")

        param = {
            "id": self.id_,
            "name": file_name,
            "file": (file_name, f, "application/pdf")
        }
        form_data = MultipartEncoder(fields=param)
        header = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.190 Safari/537.36",
            "Content-Type": form_data.content_type
        }
        # print(header)
        new_url = "{}/upload/{}".format(self.url, self.uid)
        print(new_url)
        response = requests.post(new_url, data=form_data, headers=header)
        print(response.text)
        f.close()

    def convert(self):
        """上传pdf 转换 word"""
        url = "{}/convert/{}/{}".format(self.url, self.uid, self.id_)
        print(url)
        response = requests.get(url)
        print(response.json())
        result = response.json()
        if result.get("status") == "success":
            while True:
                # 转换成功，查看转换状态，获取word文件名
                word_file_name = self.status()
                print(word_file_name)
                if word_file_name:
                    break
                #     # 下载word 文件
                #     self.download_(word_file_name)
                time.sleep(1)

    def status(self):
        url = "{}/status/{}/{}".format(self.url, self.uid, self.id_)
        print(url)
        response = requests.get(url)
        result = response.json()
        print(response.json())
        return result.get("convert_result")

    def download_(self, word_file):
        url = "{}/download/{}/{}/{}".format(self.url, self.uid, self.id_, word_file)
        response = requests.get(url)
        print(response.content)
        with open(self.word_path, "wb") as f:
            f.write(response.content)


# def main():
#     print("1111")
#     if len(sys.argv) < 2:
#         print("Usage : python3 {} [user_name]".format(sys.argv[0]))
#         return
#     pdf_list = sys.argv[1:]
#     print(pdf_list)
#     new_pdf_list = []
#     for pdf in pdf_list:
#         if not os.path.exists(pdf):
#             continue
#         file_name, ext = None, None
#         try:
#             file_name, ext = pdf.split(".")
#         except ValueError as e:
#             print("异常：{}".format(e))
#         except Exception as e:
#             print("未知错误：{}".format(e))
#         print(file_name, ext)
#         if ext not in ("pdf",):
#             continue
#         new_pdf_list.append(pdf)
#     p = PdfToWord()
#     p.process(new_pdf_list)


if __name__ == '__main__':
    pdf_path = "D:/python/test/test3/a.pdf"
    # pdf_path = "D:\\python\\test\\test3\\a.pdf"
    word_path = "C:/Users/张永浩/Desktop/a.doc"
    # word_path = r"C:\Users\张永浩\Desktop\a.doc"
    p = PdfToWord(pdf_path, word_path)
    p.process()
    print(f"pdf_path：{os.path.exists(pdf_path)}")
    print(f"word_path：{os.path.exists(word_path)}")
    # with open(pdf_path, "rb") as f:
    #     print(f.read())
