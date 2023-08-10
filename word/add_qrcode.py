#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2023/8/10 13:28
# @Author  : 杨江波
# @File    : add_qrcode.py
# @Software: PyCharm
import os

from docx import Document
from docxtpl import DocxTemplate
from qrcode import QRCode, constants


def gen_confirm_att(ORDER_CFM_DOC_ID):
    # 一、读取模板文件
    # 通过交易确认书编号查询对应合约类型

    """..........................................................."""
    file_path = "test.docx"
    document = Document(file_path)
    """..........................................................."""

    # 六、生成目标文件
    dest_temp_file = 'result.docx'
    document.save(dest_temp_file)
    # 生成新的二维码
    ORDER_CFM_DOC_ID_QRCODE_PATH = create_qrcode_handle(target_data=ORDER_CFM_DOC_ID)
    # 替换文档内的二维码
    replace_qrcode_handle(file_path=dest_temp_file, target_qrcode_path=ORDER_CFM_DOC_ID_QRCODE_PATH)

    return 0


def create_qrcode_handle(target_data: str):
    """
    为目标字符串创建二维码图片并保存
    :param target_data: 传入的目标字符串（CNTRID）
    :return: pic_path:二维码文件存放地址
    """
    # 实例Qrcode对象并设置图片尺寸
    qr = QRCode(version=1,
                error_correction=constants.ERROR_CORRECT_L,
                box_size=3,  # 图片边框像素设置
                border=3,  # 图片边框尺寸
                )
    qr.add_data(target_data)  # 传入目标数据
    qr.make(fit=True)  # 尺寸自适应
    pic_path = 'temp_{}.png'.format(target_data)  # 生成的二维码文件名称
    try:
        qr.make_image().save(pic_path)  # 生成临时二维码图像
    except Exception as e:
        print(str(e))
    else:
        return pic_path


def replace_qrcode_handle(file_path, target_qrcode_path) -> int:
    """
    替换Word文件中的二维码图片，替换后删除图片
    :param file_path: 原始Word模板文件
    :param target_qrcode_path: 准备用来替换的图片文件目录
    :return: int -1：不存在该二维码图片; -2:原始二维码图片消失，无法替换
    """
    if os.path.exists(target_qrcode_path) is False:
        return -1

    # Word文件内的原始二维码文件地址，如果被修改请及时同步模板文件内的新文件地址
    original_pic = "original_qrcode.png"
    if os.path.exists(original_pic) is False:
        return -2

    try:
        template_doc = DocxTemplate(file_path)
        template_doc.replace_media(src_file=original_pic,
                                   dst_file=target_qrcode_path)
        template_doc.save(file_path)
    except Exception as e:
        print(str(e))
    else:
        os.remove(target_qrcode_path)
        return 0


if __name__ == '__main__':
    test_code = "800762-CAN-FCF-2023081055"
    gen_confirm_att(test_code)
