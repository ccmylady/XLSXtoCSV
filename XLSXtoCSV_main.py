# coding=utf-8
import XLSXtoCSV_function as convf_
import os

print('贝克欧SAP-立库文件转换,测试版2,20201023')
yes_no_convert=input("请输入字母'y'开始转换:")

filename_purchase_path='D:\BEKOautowarehouse\purchase'
filename_delivery_path='D:\BEKOautowarehouse\delivery'
filename_manufacture_path='D:\BEKOautowarehouse\manufacture'

#确认是否转换
if yes_no_convert in ['y','Y']:
    filename_purchases=os.listdir(filename_purchase_path)
    filename_deliverys=os.listdir(filename_delivery_path)
    filename_manufctures=os.listdir(filename_manufacture_path)
    print('采购文件夹下文件目录:',filename_purchases)
    print('发货文件夹下文件目录:',filename_deliverys,'\n')
    #检查有无文件
    if len(filename_purchases):
        purchase_file_num=0
        for filename_purchase in filename_purchases:
            #检查文件格式
            if os.path.splitext(filename_purchase)[1] in ['.XLSX','.xlsx']:
                convf_.xlsx_to_csv_purchase_multi(filename_purchase_path, filename_purchase)
                purchase_file_num+=1
        print(purchase_file_num, ' 采购文件转换(非成功，请复核)','\n')
    else:
        print('无可转换的采购文件存在','\n')

    if len(filename_deliverys):
        delivery_file_num=0
        for filename_delivery in filename_deliverys:
            #检查文件格式
            if os.path.splitext(filename_delivery)[1] in ['.XLSX','.xlsx']:
                convf_.xlsx_to_csv_delivery_multi(filename_delivery_path, filename_delivery)
                delivery_file_num+=1
        print(delivery_file_num, ' 发货文件转换(非成功，请复核)','\n')
    else:
        print('无可转换的发货文件存在','\n')

    input('请输入任意键退出')
