# coding=utf-8
import logging
logging.basicConfig(level=logging.DEBUG,
                    format="%(asctime)s - %(levelname)s - %(message)s"
                    ,datefmt="%Y-%m-%d %H:%M:%S"
                    )

import XLSXtoCSV_function as convf_
import os
import shutil
import time

print('========贝克欧SAP-立库文件转换,测试版7,20220815========')
print(r'==*重要*==0410出库专用,发货订单自动录入0410,自动识别picking status,'
      r'无法识别的物料描述先库内替换，如无则整体替换,'
      r'读取C:\BEKOautowarehouse\目录ed下文件，转至C:\BEKOautowarehouse\目录ing下，并转换成CSV文件，========')
print('\n========XLSXtoCSV转换开始========')

yes_no_convert=input("请输入字母'y'开始转换:")

path_filename_purchase_origin= r'C:\BEKOautowarehouse\purchase'
path_filename_delivery_origin= r'C:\BEKOautowarehouse\delivery'
path_filename_manufacture_origin= r'C:\BEKOautowarehouse\manufacture'
path_filename_purchase_converted= r'C:\BEKOautowarehouse\purchase_converted'
path_filename_delivery_converted= r'C:\BEKOautowarehouse\delivery_converted'

#确认是否转换
if yes_no_convert in ['y','Y']:
    filename_purchases=os.listdir(path_filename_purchase_origin)
    filename_deliverys=os.listdir(path_filename_delivery_origin)
    filename_manufctures=os.listdir(path_filename_manufacture_origin)
    print('采购文件夹下文件目录:',filename_purchases)
    print('发货文件夹下文件目录:',filename_deliverys,'\n')
    #检查有无文件
    if len(filename_purchases):
        purchase_file_num=0
        path_filename_purchase_converting = convf_.create_folder(path_filename_purchase_converted)
        for filename_purchase in filename_purchases:
            #检查文件格式
            if os.path.splitext(filename_purchase)[1] in ['.XLSX','.xlsx']:
                shutil.move(os.path.join(path_filename_purchase_origin, filename_purchase),
                            path_filename_purchase_converting)
                convf_.xlsx_to_csv_purchase_multi(path_filename_purchase_converting, filename_purchase)
                purchase_file_num+=1
        print(purchase_file_num, ' 采购文件转换(非成功，请复核)','\n')
    else:
        print('****无可转换的采购文件存在****\n')

    # 仅为防止创建相同名字的文件夹
    time.sleep(0.5)

    if len(filename_deliverys):
        delivery_file_num=0
        path_filename_delivery_converting = convf_.create_folder(path_filename_delivery_converted)
        for filename_delivery in filename_deliverys:
            #检查文件格式
            if os.path.splitext(filename_delivery)[1] in ['.XLSX','.xlsx']:
                shutil.move(os.path.join(path_filename_delivery_origin, filename_delivery),
                            path_filename_delivery_converting)
                convf_.xlsx_to_csv_delivery_multi(path_filename_delivery_converting, filename_delivery)
                delivery_file_num+=1
        print(delivery_file_num, ' 发货文件转换(非成功，请复核)','\n')
    else:
        print('========无可转换的发货文件存在========\n')

    print('========XLSXtoCSV转换结束========')
    # 仅为防止创建相同名字的文件夹
    time.sleep(0.5)

    input('请输入任意键退出')


