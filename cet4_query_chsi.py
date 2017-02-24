import urllib.request
import threading
import xlrd

#初始数据
######################
data_file = 'data.xlsx'
data_file = '162四六级考生安排161205.xlsx'
out_file = 'out.cvs'
xff = 1
max_rows = 0
threading_max = 5
threads = []
now = 1
now_write = 1
lock = threading.Lock()
######################

def query_function():
    global rsh
    global xff
    global now
    global now_write
    global rows
    global max_rows
    while True:
        #print(threading.current_thread().name,1)
        lock.acquire()
        if now == (max_rows or rows):
            break
        now_query = now
        now += 1
        lock.release()
        zkzh = rsh.cell_value(now, column_data[0]).replace(' ','')
        xm = rsh.cell_value(now, column_data[1]).replace(' ','')
        num = rsh.cell_value(now, column_data[2]).replace(' ','')
        #print(threading.current_thread().name,2)
        lock.acquire()
        xff, total_score = query_data(zkzh, xm, xff, 0)
        lock.release()
        while now_query != now_write:
            #print(threading.current_thread().name,3)
            pass
        fp.write('%s,%s,%s,%s\n' % (zkzh, xm, num, total_score))
        print('%d/%d,%s 当前已完成:%.2f%%' % (now_write, rows - 1, xm, now_write * 100 / (rows - 1)))
        now_write += 1
def query_data(zkzh, xm, xff, err_num):
    #查询
    url = 'http://www.chsi.com.cn/cet/query?'
    query = {'zkzh': zkzh, 'xm': xm}
    headers = {'Referer':'http://www.chsi.com.cn/cet/','X-Forwarded-For':'127.0.0.' + str(xff)}
    query = urllib.parse.urlencode(query)
    req = urllib.request.Request(url + query,None,headers)
    for i in range(4):
        try:
            text = urllib.request.urlopen(req).read().decode('utf8')
            break
        except Exception as e:
            if i:
                print('正在进行第' + str(i) + '次重试')
            print ('哎呀，你的网络炸了呢')
    else:
        print ('人家才不会和网络差的人一起玩呢，哼')
        exit()
    #处理数据
    search_text = ['<span class="colorRed">', '</span>']
    total_score_start = text.find(search_text[0])
    total_score_start += len(search_text[0])
    total_score_end = text.find(search_text[1], total_score_start)
    try:
        total_score = str(int(text[total_score_start:total_score_end]))
        ret = (xff, total_score)
    except Exception as e:
        print(xm + '查询出错了呢,人家再帮你试一次吧')
        if err_num < 3:
            ret = query_data(zkzh, xm, xff + 1, err_num + 1)
        else:
            print('连续错误那么多次，人家不查这人了，哼～')
            return (0, 'error')
    return ret


print('读取表格数据...')
try:
    rbook = xlrd.open_workbook(data_file)
except FileNotFoundError as e:
    print ('连文件都没有，你让人家怎么查嘛，根本找不到' + data_file + '的说')
    exit()

rsh = rbook.sheet_by_index(0)
rows = rsh.nrows
column = rsh.ncols
header = rsh.row_values(0)
column_data = [None] * 3
for i in range(column):
    header_text = str(header[i])
    if '准考证' in header_text:
        column_data[0] = i
    elif '姓名' in header_text:
        column_data[1] = i
    elif '学号' in header_text:
        column_data[2] = i
print('读取完成...')


print('正在进行查询...')
fp = open(out_file,'w')
fp.write('准考证号,姓名,学号,四级成绩\n')
while now < (max_rows or rows):
    if len(threads) < threading_max:
        threads.append(threading.Thread(target=query_function))
        threads[len(threads) - 1].start()

for i in threads:
    i.join()
