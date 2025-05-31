# coding=utf-8
import logging

# from XLSXtoCSV_function import read_config

logging.basicConfig(level=logging.DEBUG,
                    format="%(asctime)s - %(levelname)s - %(message)s"
                    ,datefmt="%Y-%m-%d %H:%M:%S"
                    )

import XLSXtoCSV_function as convf_
import os
import shutil
import time
import configparser

msg=r'''=======================贝克欧SAP-立库文件转换,版本号0.0.1.241231_beta,联系人:XPC=======================
==*重要*==0410出库专用,发货订单自动录入0410,自动识别picking status,
无法识别的物料描述先库内替换，如无则整体替换,
读取应用程序所在目录下的config文件，查询待转换文件及转换后CSV文件的保存目录，执行转化并保存==='''

print(msg)
print('\n==============================XLSXtoCSV转换开始==============================')

yes_no_convert=input("请输入字母'y'或'Y'开始转换:")

# path_filename_purchase_origin= r'\\10.22.3.6\订单数据\92 SAP采购订单'
# path_filename_delivery_origin= r'\\10.22.3.6\订单数据\90 SAP销售订单'
# path_filename_manufacture_origin= r'\\10.22.3.6\订单数据\91 SAP生产订单'
# path_filename_purchase_converted= r'\\10.22.3.6\订单数据\92 SAP采购订单\1 conversion_record_purc'
# path_filename_delivery_converted= r'\\10.22.3.6\订单数据\90 SAP销售订单\1 conversion_record_sale'
# path_filename_manufacture_converted= r'\\10.22.3.6\订单数据\91 SAP生产订单\1 conversion_record_manu'
# path_filename_purchase_action= r'\\10.22.3.6\订单数据\采购订单'
# path_filename_delivery_action= r'\\10.22.3.6\订单数据\销售订单'
# path_filename_manufacture_action= r'\\10.22.3.6\订单数据\生产订单'
# print(os.path.dirname(__file__))

# 获取当前脚本所在的目录
# config_path = read_config()
config_path = r'C:\BEKOautowarehouse\config.ini'
# logging.debug(config_path)

# 创建配置解析器对象
config = configparser.ConfigParser()

# 读取配置文件
config.read(config_path,encoding='UTF-8')

# 获取数据库相关配置
path_filename_purchase_origin= config.get('database', 'path_filename_purchase_origin')
path_filename_delivery_origin= config.get('database', 'path_filename_delivery_origin')
path_filename_manufacture_origin= config.get('database', 'path_filename_delivery_origin')
path_filename_purchase_converted= config.get('database', 'path_filename_purchase_converted')
path_filename_delivery_converted= config.get('database', 'path_filename_delivery_converted')
path_filename_manufacture_converted= config.get('database', 'path_filename_manufacture_converted')
path_filename_purchase_action= config.get('database', 'path_filename_purchase_action')
path_filename_delivery_action= config.get('database', 'path_filename_delivery_action')
path_filename_manufacture_action= config.get('database', 'path_filename_manufacture_action')

# print(path_filename_purchase_origin)
filename_purchases=[]
filename_deliverys=[]

#确认是否转换
if yes_no_convert in ['y','Y']:
    try:
        filename_purchases=os.listdir(path_filename_purchase_origin)
        try:
            filename_purchases.remove('1 conversion_record_purc')
        except ValueError:
            print(r"'1 conversion_record_purc'文件夹不存在")
    except FileNotFoundError as e:
        print(e)

    try:
        filename_deliverys=os.listdir(path_filename_delivery_origin)
        try:
            filename_deliverys.remove('1 conversion_record_sale')
        except ValueError:
            print(r"'1 conversion_record_purc'文件夹不存在")
    except FileNotFoundError as e:
        print(e)

    # filename_manufctures=os.listdir(path_filename_manufacture_origin)
    print('采购文件夹下文件目录:',filename_purchases)
    print('发货文件夹下文件目录:',filename_deliverys,'\n')
    #检查有无待转换文件
    if len(filename_purchases):
        purchase_file_num=0
        path_filename_purchase_converting = convf_.create_folder(path_filename_purchase_converted)
        for filename_purchase in filename_purchases:
            #检查文件格式。如是目标文件，则移动到历史记录目录下
            if os.path.splitext(filename_purchase)[1] in ['.XLSX','.xlsx']:
                shutil.move(os.path.join(path_filename_purchase_origin, filename_purchase),
                            path_filename_purchase_converting)
                convf_.xlsx_to_csv_purchase_multi(path_filename_purchase_converting, filename_purchase)
                purchase_file_num+=1
        print(purchase_file_num, ' 个采购文件转换(非成功，请复核)','\n')
        convf_.csvfiles_copy(path_filename_purchase_converting,path_filename_purchase_action)
    else:
        print('\n****无可转换的采购文件存在****\n')

    # 仅为防止创建相同名字的文件夹
    time.sleep(0.5)
    # 如果存在待转换文件
    if len(filename_deliverys):
        delivery_file_num=0
        path_filename_delivery_converting = convf_.create_folder(path_filename_delivery_converted)
        for filename_delivery in filename_deliverys:
            #检查文件格式。如是目标文件，则移动到历史记录目录下
            if os.path.splitext(filename_delivery)[1] in ['.XLSX','.xlsx']:
                shutil.move(os.path.join(path_filename_delivery_origin, filename_delivery),
                            path_filename_delivery_converting)
                convf_.xlsx_to_csv_delivery_multi(path_filename_delivery_converting, filename_delivery)
                delivery_file_num+=1
        print(delivery_file_num, ' 个发货文件转换(非成功，请复核)','\n')
        convf_.csvfiles_copy(path_filename_delivery_converting, path_filename_delivery_action)
    else:
        print('\n****无可转换的发货文件存在****\n')

    print('==============================XLSXtoCSV转换结束==============================\n')
    # 仅为防止创建相同名字的文件夹
    time.sleep(0.5)
else:
    print('\n选择了不转换\n')
    time.sleep(0.5)
input('请输入任意键退出')


