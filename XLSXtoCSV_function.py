# coding=utf-8
import xlrd
from xlrd import xldate_as_tuple
import csv
import codecs
from datetime import datetime
import os

def xlsx_to_csv_purchase(filename_purchase_path,filename_purchase):
    """转换xlsx采购文件至csv文件"""
    #打开并读取xlsx文件
    filename_purchase_read=os.path.join(filename_purchase_path,filename_purchase)
    try:
        workbook_purchase = xlrd.open_workbook(filename_purchase_read)
    except xlrd.biffh.XLRDError:
        print('unsupported format')
    else:
        table_purchase = workbook_purchase.sheet_by_index(0)

        #准备csv文件名
        purchase_order_name=os.path.splitext(filename_purchase)[0]
        filename_purchase_write=os.path.join(filename_purchase_path,purchase_order_name+'.csv')
        goods_owner_reminder="Please input goods owner '0410/0410/others?' for order "+purchase_order_name+': '
        goods_owner=input(goods_owner_reminder)

        #写入csv文件
        with codecs.open(filename_purchase_write,'w',encoding='utf-8') as f:
            write = csv.writer(f)
            serial_number=0
            for row_num in range(table_purchase.nrows):
                row_value_sel=[]

                if row_num==0:
                    row_value_sel=[r'Vendor number\供应商编号',r'Vendor name\供应商名称',r'Purchasing Document/采购订单号',
                                   r'Material\物料代码',r'Short Text\描述',r'Plant\仓库编号','Order Quantity/数量','Order Unit']
                else:
                    row_value = table_purchase.row_values(row_num)
                    row_value_sel.append(row_value[0][0:6])
                    row_value_sel.append(row_value[0][6:].strip())
                    row_value_sel.append(row_value[1])
                    row_value_sel.append(row_value[2])
                    row_value_sel.append(row_value[3])
                    row_value_sel.append(goods_owner)
                    row_value_sel.append(row_value[4])
                    row_value_sel.append(row_value[5])

                print(row_num,row_value_sel)
                write.writerow(row_value_sel)

def xlsx_to_csv_delivery_multi(filename_delivery_path, filename_delivery):
    """转换xlsx发货文件至csv文件"""
    #打开并读取xlsx文件
    filename_delivery_read=os.path.join(filename_delivery_path,filename_delivery)
    try:
        workbook_delivery = xlrd.open_workbook(filename_delivery_read)
    except xlrd.biffh.XLRDError:
        print('unsupported format')
    else:
        table_delivery = workbook_delivery.sheet_by_index(0)

        delivery_order_names=[]
        for row_num in range(table_delivery.nrows-1):
            delivery_order_names.append(table_delivery.row_values(row_num+1)[0])
        #print(delivery_order_names)
        delivery_order_names_new=list(set(delivery_order_names))
        delivery_order_names_new.sort(key=delivery_order_names.index)
        #print(delivery_order_names_new)

        for delivery_order_name in delivery_order_names_new:
            filename_delivery_write=os.path.join(filename_delivery_path,delivery_order_name+'.csv')
            goods_owner_reminder="Please input goods owner '0410/0410/others?' for order "+delivery_order_name+': '
            goods_owner=input(goods_owner_reminder)
            with codecs.open(filename_delivery_write,'w',encoding='utf-8') as f:
                write = csv.writer(f)
                serial_number=0
                for row_num in range(table_delivery.nrows):
                    row_value_sel=[]
                    row_value = table_delivery.row_values(row_num)
                    if row_num==0:
                        row_value_sel=['发运单号','发货日期','客户代码','客户名称','销售订单号',
                                       '物料编号','货主','Item','物料描述','数量']
                        print(row_num,row_value_sel)
                        write.writerow(row_value_sel)
                    elif row_value[0]==delivery_order_name:
                        row_value_sel.append(row_value[0])
                        #日期转换，首行除外
                        goods_issue_date=datetime(*xldate_as_tuple(row_value[4],0))
                        row_value_sel.append(goods_issue_date.strftime('%Y-%m-%d'))
                        row_value_sel.append(row_value[10])
                        row_value_sel.append(row_value[12])
                        row_value_sel.append(row_value[13])
                        row_value_sel.append(row_value[22])
                        row_value_sel.append(goods_owner)
                        serial_number+=10
                        row_value_sel.append(str(serial_number))
                        row_value_sel.append(row_value[24])
                        row_value_sel.append(row_value[25])

                        print(row_num,row_value_sel)
                        write.writerow(row_value_sel)


if __name__ == '__main__':
    xlsx_to_csv_delivery_multi('E:\STUDY\python\XLSXtoCSV\BEKOautowarehouse\delivery', 'export.XLSX')
    xlsx_to_csv_purchase('E:\STUDY\python\XLSXtoCSV\BEKOautowarehouse\purchase', '4500226596.XLSX')
