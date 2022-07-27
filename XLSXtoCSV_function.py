# coding=utf-8
import xlrd
from xlrd import xldate_as_tuple
import csv
import codecs
from datetime import datetime
import os


def xlsx_to_csv_purchase_multi(filename_purchase_path, filename_purchase):
    """转换xlsx采购文件至csv文件"""
    print('===========采购入库xlsx文件转csv文件程序开始===========')
    # 打开并读取xlsx文件
    filename_purchase_read = os.path.join(filename_purchase_path, filename_purchase)
    table_purchase_title_req_all = ['Vendor/supplying plant','Purchasing Document','Material','Short Text','Plant',
                                    'Order Quantity','Order Unit']

    try:
        #打开文件
        workbook_purchase = xlrd.open_workbook(filename_purchase_read)

        table_purchase = workbook_purchase.sheet_by_index(0)

        # 根据所需内容检索表头位置
        table_purchase_title = table_purchase.row_values(0)
        #print(table_purchase.row_values(0))
        table_purchase_title_req_position = []
        for table_purchase_title_req in table_purchase_title_req_all:
            table_purchase_title_req_position.append(table_purchase_title.index(table_purchase_title_req))
        #print(table_purchase_title_req_position)

    except xlrd.biffh.XLRDError as error:
        print('打开采购文件',filename_purchase,'时发生错误:',error)

    except ValueError as error:
        print('读取采购文件',filename_purchase,'时发生错误:',error)

    else:
        #读取采购订单编号
        purchase_order_names = []
        for row_num in range(table_purchase.nrows - 1):
            purchase_order_names.append(table_purchase.row_values(row_num + 1)[table_purchase_title_req_position[1]])
        # print(purchase_order_names)
        purchase_order_names_new = list(set(purchase_order_names))
        purchase_order_names_new.sort(key=purchase_order_names.index)
        # print(purchase_order_names_new)


        for purchase_order_name in purchase_order_names_new:
            filename_purchase_write = os.path.join(filename_purchase_path, purchase_order_name + '.csv')
            #goods_owner_reminder = "Please input goods owner '0410/0410/others?' for order " + purchase_order_name + ': '
            #goods_owner = input(goods_owner_reminder)
            with codecs.open(filename_purchase_write, 'w', encoding='gbk') as f:
                write = csv.writer(f)
                serial_number = 0
                for row_num in range(table_purchase.nrows):
                    row_value_sel = []
                    row_value = table_purchase.row_values(row_num)
                    if row_num == 0:
                        row_value_sel = ['Vendor number/供应商编号','Vendor name/供应商名称',
                                         'Purchasing Document/采购订单号','Material/物料代码','Short Text/描述',
                                         'Plant/仓库编号','Order Quantity/数量','Order Unit/单位']
                        print(row_num, row_value_sel)
                        write.writerow(row_value_sel)
                    elif row_value[table_purchase_title_req_position[1]] == purchase_order_name:
                        row_value = table_purchase.row_values(row_num)
                        row_value_sel.append(row_value[table_purchase_title_req_position[0]][0:11].strip())
                        row_value_sel.append(row_value[table_purchase_title_req_position[0]][11:].strip())
                        row_value_sel.append(row_value[table_purchase_title_req_position[1]])
                        row_value_sel.append(row_value[table_purchase_title_req_position[2]])
                        row_value_sel.append(row_value[table_purchase_title_req_position[3]])
                        #row_value_sel.append(goods_owner)
                        row_value_sel.append(row_value[table_purchase_title_req_position[4]])
                        row_value_sel.append(row_value[table_purchase_title_req_position[5]])
                        row_value_sel.append(row_value[table_purchase_title_req_position[6]])

                        print(row_num, row_value_sel)
                        write.writerow(row_value_sel)
        print('===========采购入库xlsx文件转csv文件程序结束===========\n')

def xlsx_to_csv_delivery_multi(filename_delivery_path, filename_delivery):
    """转换xlsx发货文件至csv文件"""
    print('===========销售出库xlsx文件转csv文件程序开始===========')
    #打开并读取xlsx文件
    filename_delivery_read=os.path.join(filename_delivery_path,filename_delivery)
    table_delivery_title_req_all = ['Delivery', 'Goods Issue Date','Sold-to party', 'Name of sold-to party',
                                    'Reference document','Material', 'Item','Description','Delivery quantity',
                                    'Purchase order number']
    table_delivery_title_req_Picking_status='Picking status'

    try:
        workbook_delivery = xlrd.open_workbook(filename_delivery_read)

        table_delivery = workbook_delivery.sheet_by_index(0)

        # 根据内容检索表头位置
        table_delivery_title = table_delivery.row_values(0)
        #print(table_delivery.row_values(0))
        table_delivery_title_req_position = []
        for table_delivery_title_req in table_delivery_title_req_all:
            table_delivery_title_req_position.append(table_delivery_title.index(table_delivery_title_req))

        # 当存在picking status时，按picking status判断，否则默认全部转。
        if table_delivery_title_req_Picking_status in table_delivery_title:
            print('****文件中存在picking status，将以其值作为依据****')
            picking_status_onoff=True
            picking_status_position=table_delivery_title.index(table_delivery_title_req_Picking_status)
        #print(table_delivery_title_req_position)



    except xlrd.biffh.XLRDError as error:
        print('打开发货文件',filename_delivery,'时发生错误:',error)

    except ValueError as error:
        print('读取发货文件',filename_delivery,'时发生错误:',error)

    else:
        # 读取发货订单编号
        delivery_order_names=[]
        for row_num in range(table_delivery.nrows-1):
            delivery_order_names.append(table_delivery.row_values(row_num+1)[table_delivery_title_req_position[0]])
        #print(delivery_order_names)
        delivery_order_names_new=list(set(delivery_order_names))
        delivery_order_names_new.sort(key=delivery_order_names.index)
        #print(delivery_order_names_new)

        # goods_owner_reminder = "Please input goods owner '0410/0410/others?' for order " + 'delivery_order_names' + ': '
        # goods_owner = input(goods_owner_reminder)
        goods_owner = "0410"
        count_of_UnicodeEncodeError=0

        for delivery_order_name in delivery_order_names_new:
            filename_delivery_write=os.path.join(filename_delivery_path,delivery_order_name+'.csv')
            # goods_owner_reminder="Please input goods owner '0410/0410/others?' for order "+delivery_order_name+': '
            # goods_owner=input(goods_owner_reminder)
            # 'GB18030'
            with codecs.open(filename_delivery_write,'w',encoding='GBK') as f:
                write = csv.writer(f)
                serial_number=0
                for row_num in range(table_delivery.nrows):
                    row_value_sel=[]
                    row_value = table_delivery.row_values(row_num)
                    if row_num==0:
                        row_value_sel=['发运单号','发货日期','客户代码','客户名称','销售订单号',
                                       '物料编号','货主','物项','物料描述','数量','参考号']
                        print(row_num,row_value_sel)
                        write.writerow(row_value_sel)
                    elif row_value[table_delivery_title_req_position[0]]==delivery_order_name:
                        if picking_status_onoff and row_value[picking_status_position] != 'C':
                            print('{} **{} 发运单中物料 {} {} 被识别为忽略代码，无需出库'.format(
                                row_num,
                                row_value[table_delivery_title_req_position[0]],
                                row_value[table_delivery_title_req_position[5]],
                                row_value[table_delivery_title_req_position[7]]))
                            continue
                        else:
                            row_value_sel.append(row_value[table_delivery_title_req_position[0]])
                            #日期转换，首行除外
                            goods_issue_date=datetime(*xldate_as_tuple(row_value[table_delivery_title_req_position[1]],0))
                            row_value_sel.append(goods_issue_date.strftime('%Y-%m-%d'))
                            row_value_sel.append(row_value[table_delivery_title_req_position[2]])
                            row_value_sel.append(row_value[table_delivery_title_req_position[3]])
                            row_value_sel.append(row_value[table_delivery_title_req_position[4]])
                            row_value_sel.append(row_value[table_delivery_title_req_position[5]])
                            row_value_sel.append(goods_owner)
                            #serial_number+=10
                            #row_value_sel.append(str(serial_number))
                            row_value_sel.append(row_value[table_delivery_title_req_position[6]])
                            row_value_sel.append(row_value[table_delivery_title_req_position[7]])
                            row_value_sel.append(row_value[table_delivery_title_req_position[8]])
                            row_value_sel.append(row_value[table_delivery_title_req_position[9]])

                            print(row_num,row_value_sel)
                            try:
                                write.writerow(row_value_sel)
                            except UnicodeEncodeError:
                                row_value_sel[8]='?物料描述存在未知字符?'
                                write.writerow(row_value_sel)
                                count_of_UnicodeEncodeError+=1
                                print('***物料名称描述中存在未知字符?,已替换，计数{}***'.format(count_of_UnicodeEncodeError))

        print('===========销售出库xlsx文件转csv文件程序结束===========\n')

if __name__ == '__main__':
    xlsx_to_csv_delivery_multi(r'E:\STUDYing\PYTHON\Example\BEKOautowarehouse\delivery', 'export.XLSX')
    xlsx_to_csv_purchase_multi(r'E:\STUDYing\PYTHON\Example\BEKOautowarehouse\purchase', 'export2.XLSX')
