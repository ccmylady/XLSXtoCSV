# coding=utf-8
import XLSXtoCSV_function as convf_
import os

yes_no_convert=input("Please input 'y' if convert:")

filename_purchase_path='D:\BEKOautowarehouse\purchase'
filename_delivery_path='D:\BEKOautowarehouse\delivery'
filename_manufacture_path='D:\BEKOautowarehouse\manufacture'

#确认是否转换
if yes_no_convert=='y':
    filename_purchases=os.listdir(filename_purchase_path)
    filename_deliverys=os.listdir(filename_delivery_path)
    filename_manufctures=os.listdir(filename_manufacture_path)
    print(filename_purchases)
    print(filename_deliverys)
    #检查有无文件
    if len(filename_purchases):
        purchase_file_num=0
        for filename_purchase in filename_purchases:
            #检查文件格式
            if os.path.splitext(filename_purchase)[1]=='.XLSX':
                convf_.xlsx_to_csv_purchase_multi(filename_purchase_path, filename_purchase)
                purchase_file_num+=1
        print(purchase_file_num, ' purchase files convert')
    else:
        print('no purchase files exist')

    if len(filename_deliverys):
        delivery_file_num=0
        for filename_delivery in filename_deliverys:
            #检查文件格式
            if os.path.splitext(filename_delivery)[1]=='.XLSX':
                convf_.xlsx_to_csv_delivery_multi(filename_delivery_path, filename_delivery)
                delivery_file_num+=1
        print(delivery_file_num, ' delivery files convert')
    else:
        print('no delivery files exist')
