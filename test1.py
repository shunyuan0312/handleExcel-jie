import time
import datetime
# localtime = time.localtime(time.time())
# print(localtime)

# local = time.strftime("%Y%m%d%H%M%S", time.localtime())
# print(local)
# c_new_file_name = "test.xlsx"
# c_new_file_name = c_new_file_name.split('.')[0] + local + '.' + c_new_file_name.split('.')[1]
# print(c_new_file_name)
a = time.localtime()
print(a.tm_year, a.tm_mon, a.tm_mday)