import time
import datetime
# localtime = time.localtime(time.time())
# print(localtime)

# local = time.strftime("%Y%m%d%H%M%S", time.localtime())
# print(local)
# c_new_file_name = "test.xlsx"
# c_new_file_name = c_new_file_name.split('.')[0] + local + '.' + c_new_file_name.split('.')[1]
# print(c_new_file_name)
# a = time.localtime()
# print(a.tm_year, a.tm_mon, a.tm_mday)

# list1 = [x for x in range(100000)]
# print(max(list1))
# del_A_sht1_row_num_lists = [31, '测试单位31', '422721195309051211', 123, '李应想', 31, '测试项目31', datetime.datetime(2022, 1, 31, 0, 0), datetime.date(2022, 4, 19), '正常']
# for row_num in range(len(del_A_sht1_row_num_lists),0,-1):
#     print(row_num)
# d = ["b1","d2","d3"]
# lis = [["a1","a2","a3"],["b1","b2","b3"],["c1","c2","c3"]]
# lis = [["a1","a2","a3"],["b1","b2","b3"],["c1","c2","c3"],["d1","d2","d3"]]

# new = [list(t) for t in set(tuple(a) for a in lis)]
# print(new)
# a = False
# for i in lis:
#     if i[0] == d[0]:
#         a = False
#     else:
#         a = True


    
# #         lis.append("d1")
# lis.append(d)
# print(lis)

a = ['6224120061467123', '6224122014160123', '', '', '6224122027674123', '6230550008452123', '', '', '', '', '6224122014142123', '6224122013957123', '6224120064238123', '', '6224121173224123', '', '6224122022129123', '6210130005263123', '6210130005276123', '6210137281175123', '6210130993032123', '6210134951132123', '', '', '6210134094329123', '6210135642550123', '', '6224120081371123', '6224122013833123', '', '6230530008452123', '6224120008452123', '6224120026060123', '6210135642542123']
if a[2] == '':
    print(1)