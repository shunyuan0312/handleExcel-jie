import os
from handle_excel import HandleExcel
import time
import xlrd


dic = {'1':'A','2':'B','3':'C','4':'D','5':'E','6':'F','7':'G','8':'H','9':'I','10':'J','11':'K','12':'L','13':'M','14':'N','15':'O','16':'P','17':'Q','18':'R','19':'S','20':'T','21':'U','22':'V','23':'W','24':'X','25':'Y','26':'Z'}

def down_to_fifteen(eighteen_card):
    return eighteen_card[0:6] + eighteen_card[8:17]

def update_tableC():
    
    for i in range(len(c_id_cards_list)):
        print("处理第%s条"%i)
        # 如果C表中该用户的身份证可以在A表身份证列查到
        if c_id_cards_list[i] in a_id_cards_list:
            a_index = a_id_cards_list.index(c_id_cards_list[i])
            # 判断表c的代发账号是否在表b的三列中存在
            row_num = False
            if c_issue_account_list[i] in b_account_list:
                row_num = b_account_list.index(c_issue_account_list[i]) + 1 + 1  # B表第一行为标题
            if c_issue_account_list[i] in b_card_list:
                row_num = b_card_list.index(c_issue_account_list[i]) + 1 + 1
            if c_issue_account_list[i] in b_old_account_list:
                row_num = b_old_account_list.index(c_issue_account_list[i]) + 1 + 1
            # 更新表c的I-L列
            a_sht.Range('C' + str(a_index + a_id_cards_row) + ':' + 'E' + str(a_index + a_id_cards_row)).Copy()
            c_sht.Paste(Destination=c_sht.Range('I' + str(i + c_issue_account_row)))
            a_sht.Range('G' + str(a_index + a_id_cards_row)).Copy()
            c_sht.Paste(Destination=c_sht.Range('L' + str(i + c_issue_account_row)))
            # 如果C表的账号在B表三列中存在
            if row_num:  
                account = b_table.getCell('Sheet1', row_num, account_idx + 1)
                card = b_table.getCell('Sheet1', row_num, card_idx + 1)
                old_account = b_table.getCell('Sheet1', row_num, old_account_idx + 1)
                all_account = [account, card, old_account]
                # 如果A表账户在B表三个值中存在
                if a_issue_account_list[a_index] in all_account:
                    # C表H列，“一套”
                    c_table.setCell('Sheet1', i+c_issue_account_row, 8, "一套")
                    # C表N列：“OK”
                    c_table.setCell('Sheet1', i+c_issue_account_row, 14, "OK")
                else:
                    # 非一套，更新表c代发账户，M列为C表的原代发账户
                    # C表H列，"非一套，已将原代发账户替换为一卡通账户"
                    c_table.setCell('Sheet1', i+c_issue_account_row, 8, "非一套，已将原代发账户替换为一卡通账户")
                    # C表N列：“OK”
                    c_table.setCell('Sheet1', i+c_issue_account_row, 14, "OK")
                    # C表M列，保留原代发账户数据
                    c_sht.Range(dic.get(str(c_issue_account_col)) + str(i + c_issue_account_row)).Copy()
                    c_sht.Paste(Destination=c_sht.Range('M' + str(i + c_issue_account_row)))
                    # 替换C表D列的代发账户为A表的一卡通账户
                    a_sht.Range(dic.get(str(a_issue_account_col)) + str(a_index + a_issue_account_row)).Copy()
                    c_sht.Paste(Destination=c_sht.Range(dic.get(str(c_issue_account_col)) + str(i + c_issue_account_row)))                
            else:
                if c_issue_account_list[i] == a_issue_account_list[a_index]:
                    # C表H列，“一套”
                    c_table.setCell('Sheet1', i+c_issue_account_row, 8, "一套")
                    # C表N列：“OK”
                    c_table.setCell('Sheet1', i+c_issue_account_row, 14, "OK")
                else:
                    # 非一套，更新表c代发账户，更新c表的I-L列，M列为C表的原代发账户'
                    # C表H列，"非一套，已将原代发账户替换为一卡通账户"
                    c_table.setCell('Sheet1', i+c_issue_account_row, 8, "非一套，已将原代发账户替换为一卡通账户")
                    # C表N列：“OK”
                    c_table.setCell('Sheet1', i+c_issue_account_row, 14, "OK")
                    # C表M列，保留原代发账户数据
                    c_sht.Range(dic.get(str(c_issue_account_col)) + str(i + c_issue_account_row)).Copy()
                    c_sht.Paste(Destination=c_sht.Range('M' + str(i + c_issue_account_row)))
                    # 替换C表D列的代发账户为A表的一卡通账户
                    a_sht.Range(dic.get(str(a_issue_account_col)) + str(a_index + a_issue_account_row)).Copy()
                    c_sht.Paste(Destination=c_sht.Range(dic.get(str(c_issue_account_col)) + str(i + c_issue_account_row)))
        else:
            # 判断C表的代发账号是否在B表的三列（账号、卡号、旧帐号）数据中存在
            row_num = False
            if c_issue_account_list[i] in b_account_list:
                row_num = b_account_list.index(c_issue_account_list[i]) + 1 + 1
            if c_issue_account_list[i] in b_card_list:
                row_num = b_card_list.index(c_issue_account_list[i]) + 1 + 1
            if c_issue_account_list[i] in b_old_account_list:
                row_num = b_old_account_list.index(c_issue_account_list[i]) + 1 + 1
            # 如果存在 422723196110091213、42272319700917121x
            if row_num:
                # 判断C表和B表的身份证是否一致
                b_id = b_table.getCell('Sheet1', row_num, id_cards_idx + 1)
                c_id = c_id_cards_list[i]
                # 复制b表的身份证
                b_sht.Range(dic.get(str(id_cards_idx+1)) + str(row_num)).Copy()
                c_sht.Paste(Destination=c_sht.Range(dic.get('15') + str(i+c_issue_account_row)))
                if b_id == c_id:
                    # 一致，C表N列“OK”，O列“取B表身份证号”
                    c_table.setCell('Sheet1', i+c_issue_account_row, 14, "OK")
                    # c_sht.Paste(Destination=c_sht.Range(dic.get('15') + str(i+c_issue_account_row)))
                elif b_id[-1].isalpha():
                    if b_id.upper() == c_id:
                        c_table.setCell('Sheet1', i+c_issue_account_row, 14, "代发账户数据异常，建议重新提供代发账户")
                        # c_sht.Paste(Destination=c_sht.Range(dic.get('15') + str(i+c_issue_account_row)))
                        c_table.setCell('Sheet1', i+c_issue_account_row, 16, "身份证末位为小写x")
                else:
                    # C表N列，数据异常；O列“取B表身份证号”
                    c_table.setCell('Sheet1', i+c_issue_account_row, 14, "代发账户数据异常，建议重新提供代发账户")
                    # c_sht.Paste(Destination=c_sht.Range(dic.get('15') + str(i+c_issue_account_row)))
                    # 不一致，判断是否是18和15位新旧身份证问题
                    if down_to_fifteen(c_id) == b_id:
                        # 身份证未升位，P列“旧身份证未升位”
                        c_table.setCell('Sheet1', i+c_issue_account_row, 16, "旧身份证未升位")
                    else:
                        # 身份证完全不一致；P列“身份证完全不一致”
                        c_table.setCell('Sheet1', i+c_issue_account_row, 16, "身份证号完全不一致")
            # 如果不存在
            else:
                # C表N列“暂未提取到数据”
                c_table.setCell('Sheet1', i+c_issue_account_row, 14, "暂未提取到数据")
                   
    a_table.close()
    b_table.close()
    c_table.close()


if __name__ == '__main__':
    print("开始执行！")
    start = time.perf_counter()
    a_file_name = "表A-手工数据母表.xlsx"
    a_sht_name = '母表'
    b_file_name = '表B-后台导出数据.xlsx'
    b_sht_name = 'Sheet1'
    c_file_name = "表C-待处理数据.xlsx"
    c_sht_name = 'Sheet1'


    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    a_file_path = os.path.join(base_dir, a_file_name)
    a_table = HandleExcel(a_file_path)
    a_sht = a_table.xlsx_book.Worksheets(a_sht_name)

    
    b_file_path = os.path.join(base_dir, b_file_name)
    b_table = HandleExcel(b_file_path)
    b_sht = b_table.xlsx_book.Worksheets(b_sht_name)


    c_file_path = os.path.join(base_dir, c_file_name)
    c_table = HandleExcel(c_file_path)
    c_sht = c_table.xlsx_book.Worksheets(c_sht_name)

    a_id_cards_row,  a_id_cards_col = 2, 3
    a_id_cards_list = a_table.get_col_list(a_sht_name, a_id_cards_row, a_id_cards_col)
    a_issue_account_row, a_issue_account_col = 2, 4
    a_issue_account_list = a_table.get_col_list(a_sht_name, a_issue_account_row, a_issue_account_col)

    c_id_cards_row, c_id_cards_col = 3, 3
    c_id_cards_list = c_table.get_col_list(c_sht_name, c_id_cards_row, c_id_cards_col)
    c_issue_account_row, c_issue_account_col = 3, 4
    c_issue_account_list = c_table.get_col_list(c_sht_name, c_issue_account_row, c_issue_account_col)

    # B表
    title_list = b_table.get_row_list(b_sht_name, 1, 1)

    account_idx = title_list.index('账号')
    card_idx = title_list.index('卡号')
    old_account_idx = title_list.index('旧账号')
    id_cards_idx = title_list.index('证件号码')

    b_account_list = b_table.get_col_list(b_sht_name, 2, account_idx+1)
    b_card_list = b_table.get_col_list(b_sht_name, 2, card_idx+1)
    b_old_account_list = b_table.get_col_list(b_sht_name, 2, old_account_idx+1)

    update_tableC()
    end = time.perf_counter()
    print("执行成功！")
    print('Running time: %s Seconds' %(end-start))