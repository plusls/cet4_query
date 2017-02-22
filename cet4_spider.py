import urllib.request
import xlrd
#import xlwt
def solve_data(text):
    search_text = ['<span class="colorRed">', '</span>']
    score_start = text.find(search_text[0])
    score_start += len(search_text[0])
    score_end = text.find(search_text[1], score_start)
    score = str(int(text[score_start:score_end]))
    return score
def get_data(zkzh,xm):
    url = 'http://www.chsi.com.cn/cet/query?'
    query = {'zkzh': zkzh, 'xm': xm}
    headers = {'Referer':'http://www.chsi.com.cn/cet/','X-Forwarded-For':'127.0.0.1'}
    query = urllib.parse.urlencode(query)
    req = urllib.request.Request(url + query,None,headers)
    text = urllib.request.urlopen(req).read().decode('utf8')
    return solve_data(text)

	
filename = '计院.xlsx'
print('读取表格数据...')
rbook = xlrd.open_workbook(filename)
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
rows = 8
fp = open('out.txt','w')
fp.write('准考证号,姓名,学号,四级成绩\n')
for i in range(1, rows):
    zkzh = rsh.cell_value(i, column_data[0])
    xm = rsh.cell_value(i, column_data[1])
    num = rsh.cell_value(i, column_data[2])
    score = get_data(zkzh, xm)
    fp.write('%s,%s,%s,%s\n' % (zkzh, xm, num, score))
    print('%d/%d,当前已完成:%.2f%%' % (i, rows - 1, i * 100 / (rows - 1)))