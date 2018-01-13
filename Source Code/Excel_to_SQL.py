import xlrd, xlwt, MySQLdb, sqlite3, getpass
from datetime import datetime, timedelta

x = 0
y = 0
z = 0
original_sheet_names = ['Instructions','user_tab','order_tab']

def xldate_to_datetime(xldate):
    tempdate = datetime(1900, 1, 1)
    deltadays = timedelta(days=int(xldate)-2)
    thetime = tempdate + deltadays
    return thetime.strftime('%Y-%m-%d')

print('*** Shopee Pre-Interview Questions to Result Converter ***')
while x < 1:
    try:
        book = xlrd.open_workbook('SQL Test (1).xlsx')
        x = x + 1
        print('')
        print('Reading SQL Test (1).xlsx . . .')
    except:
        print('SQL Test (1).xlsx is not found')
        print('SQL Test (1).xlsx must be placed in the same directory as SQL Test (1)-Answers.exe')
        userinput = input('Place the SQL Test (1).xlsx file in the same directory and hit enter. To quit the program, type "quit": ')
        if len(userinput) < 1: continue
        if userinput.lower() == "quit": quit()
        if len(userinput) > 1: continue
while y < 1:
    book = xlrd.open_workbook('SQL Test (1).xlsx')
    sheet_names = book.sheet_names()
    if 'Instructions' and 'user_tab' and 'order_tab' in sheet_names:
        print('Sheet names:', sheet_names)
        y = y + 1
    else:
        print('Error, this copy of SQL Test (1).xlsx is not genuie')
        user_input = input('Place the original SQL Test (1).xlsx and hit enter. To quit the program type "quit": ')
        print('')
        if len(user_input) < 1: continue
        if user_input.lower() == "quit": quit()
        if len(user_input) > 1: continue
while z < 1:
    print('')
    username = input('MySQL username: ')
    if len(username) < 1: username = 'root'
    password = getpass.getpass('MySQL password (no asterisks nor characters will be shown): ')
    if len(password) < 1: password = 'root'
    try:
        database = MySQLdb.connect (host = 'localhost', user = username, passwd = password)
        cur = database.cursor()
        cur.execute('''DROP DATABASE IF EXISTS Shopee_SQL_Test''')
        cur.execute('''CREATE DATABASE Shopee_SQL_Test''')
        cur.execute('''USE Shopee_SQL_Test''')
        cur.execute('''CREATE TABLE IF NOT EXISTS user_tab (userid INTEGER UNIQUE, register_time DATE, country TEXT)''')
        cur.execute('''CREATE TABLE IF NOT EXISTS order_tab (orderid INTEGER UNIQUE, userid INTEGER, itemid INTEGER, gmv FLOAT, order_time DATE)''')
        print('MySQL connection is succesful . . .')
        print('Converting data set to MySQL . . .')
        z = z + 1
    except:
        print('MySQL connection failed . . .')
        userinput = input('Connect to MySQL (Y/N)?: ')
        if userinput.lower() == 'n': z = z + 1
        if userinput.lower() == 'y': continue
        else: continue

print('Converting data set to sqlite . . .')
sqlitedb = sqlite3.connect('Shopee_SQL_Test.sqlite')
sqlitecur = sqlitedb.cursor()
sqlitecur.execute('CREATE TABLE IF NOT EXISTS user_tab (userid INTEGER UNIQUE, register_time DATETIME, country TEXT)')
sqlitecur.execute('CREATE TABLE IF NOT EXISTS order_tab (orderid INTEGER UNIQUE, userid INTEGER, itemid INTEGER, gmv FLOAT, order_time DATETIME)')

user_tab = book.sheet_by_index(1)
for row in range(1, user_tab.nrows):
    userid = int(user_tab.cell(row,0).value)
    register_time = xldate_to_datetime(user_tab.cell(row,1).value)
    country = str(user_tab.cell(row,2).value)
    sqlitecur.execute('INSERT OR IGNORE INTO user_tab VALUES(?, ?, ?)',(userid, register_time, country))
    try: cur.execute('''INSERT IGNORE INTO user_tab VALUES(%s, %s,%s)''',(userid, register_time, country))
    except: pass
sqlitedb.commit()
try: database.commit()
except: pass

order_tab = book.sheet_by_index(2)
for row in range(1, order_tab.nrows):
    orderid = int(order_tab.cell(row,0).value)
    userid = int(order_tab.cell(row,1).value)
    itemid = int(order_tab.cell(row,2).value)
    gmv = float(order_tab.cell(row,3).value)
    order_time = xldate_to_datetime(order_tab.cell(row,4).value)
    sqlitecur.execute('INSERT OR IGNORE INTO order_tab VALUES (?, ?, ?, ?, ?)', (orderid, userid, itemid, gmv, order_time))
    try: cur.execute('''INSERT IGNORE INTO order_tab VALUES(%s, %s, %s, %s, %s)''',(orderid, userid, itemid, gmv, order_time))
    except: pass
sqlitedb.commit()
try: database.commit()
except: pass

sqlitedb.close()
try: database.close()
except: pass

print('')
print(str(user_tab.nrows),'rows from',original_sheet_names[1],'and', str(order_tab.nrows),'rows from', original_sheet_names[2],'are converted')
print('*.xlsx to MySQL and *.sqlite conversion is completed . . .')

print('''
Questions:
1) Write an SQL statement to count the number of users per country (5 marks)
2) Write an SQL statement to count the number of orders per country (10 marks)
3) Write an SQL statement to find the first order date of each user (10 marks)
4) Write an SQL statement to find the number of users who made their first order in each country, each day (25 marks)
5) Write an SQL statement to find the first order GMV of each user. If there is a tie, use the order with the lower orderid (30 marks)
6) Find out what is wrong with the sample data (20 marks)''')
print('Getting the correct answers . . .')

book.release_resources()
sqlitedb = sqlite3.connect('file:Shopee_SQL_Test.sqlite?mode=ro', uri=True)
sqlitecur = sqlitedb.cursor()

answer_1 = ('''SELECT country, COUNT(DISTINCT userid) AS NumberOfUsers
                    FROM user_tab
                    GROUP BY country''')
answer_2 = ('''SELECT user_tab.country AS Country, COUNT(orderid) AS NumberOfOrders
                    FROM user_tab JOIN order_tab
                    ON user_tab.userid=order_tab.userid
                    GROUP BY user_tab.country''')
answer_3 = ('''SELECT userid, MIN(order_time) AS FirstOrderTime
                    FROM order_tab
                    GROUP BY userid''')
answer_4 = ('''SELECT u.country AS Country, COUNT(u.userid) AS NumOfUsers
                    FROM user_tab u JOIN order_tab o ON u.userid = o.userid
                    WHERE o.order_time IN (SELECT MIN(order_time)
                                            FROM order_tab o JOIN user_tab u
                                            ON o.userid=u.userid GROUP BY u.country)
                    GROUP BY u.country''')
answer_5 = ('''SELECT x.userid AS UserID, x.gmv AS GMV, x.orderid AS OrderID
FROM (SELECT a.userid, a.gmv, a.orderid, COUNT(*) AS FirstOrderGMV
		FROM order_tab a JOIN order_tab b
		ON a.userid = b.userid
		WHERE a.gmv < b.gmv OR (a.gmv = b.gmv AND a.orderid >= b.orderid)
		GROUP BY a.userid, a.gmv, a.orderid
		HAVING FirstOrderGMV = 1) AS x''')
answer_6 = ('''
1. Some values in order_time column are earlier than the values found in register_time column.
2. There are duplicates in country column in user_tab table.
3. There should be another table that assigns country's id to each corresponding country's name''')

answerbook = xlwt.Workbook(encoding='utf-8')
sheet1 = answerbook.add_sheet('Answer - 1')
sheet2 = answerbook.add_sheet('Answer - 2')
sheet3 = answerbook.add_sheet('Answer - 3')
sheet4 = answerbook.add_sheet('Answer - 4')
sheet5 = answerbook.add_sheet('Answer - 5')
sheet6 = answerbook.add_sheet('Answer - 6')

style = xlwt.XFStyle()
font = xlwt.Font()
font.bold = True
style.font = font

#1
sheet1.write(0,0, 'Country', style=style)
sheet1.write(0,1, 'NumberOfUsers', style=style)
sheet1.write(0,4, 'SQL Query: ', style=style)
for r, row in enumerate(answer_1.replace('\t','').split('\n')):
    sheet1.write(r+1,4, row)
sqlitecur.execute(answer_1)
for r, row in enumerate(sqlitecur):
    sheet1.write(r+1, 0, row[0])
    sheet1.write(r+1, 1, row[1])

#2
sheet2.write(0,0, 'Country', style=style)
sheet2.write(0,1, 'NumberOfOrders', style=style)
sheet2.write(0,4, 'SQL Query: ', style=style)
for r, row in enumerate(answer_2.replace('\t','').split('\n')):
    sheet2.write(r+1, 4, row)
sqlitecur.execute(answer_2)
for r,row in enumerate(sqlitecur):
    sheet2.write(r+1, 0, row[0])
    sheet2.write(r+1, 1, row[1])

#3
sheet3.write(0,0, 'UserID', style=style)
sheet3.write(0,1, 'FirstOrderTime', style=style)
sheet3.write(0,4, 'SQL Query: ', style=style)
for r, row in enumerate(answer_3.replace('\t','').split('\n')):
    sheet3.write(r+1, 4, row)
sqlitecur.execute(answer_3)
for r,row in enumerate(sqlitecur):
    sheet3.write(r+1, 0, row[0])
    sheet3.write(r+1, 1, row[1])

#4
sheet4.write(0,0, 'Country', style=style)
sheet4.write(0,1, 'NumOfUsers', style=style)
sheet4.write(0,4, 'SQL Query: ', style=style)
for r, row in enumerate(answer_4.replace('\t','').split('\n')):
    sheet4.write(r+1, 4, row)
sqlitecur.execute(answer_4)
for r,row in enumerate(sqlitecur):
    sheet4.write(r+1, 0, row[0])
    sheet4.write(r+1, 1, row[1])

#5
sheet5.write(0,0, 'UserID', style=style)
sheet5.write(0,1, 'GMV', style=style)
sheet5.write(0,2, 'OrderID', style=style)
sheet5.write(0,4, 'SQL Query: ', style=style)
for r, row in enumerate(answer_5.replace('\t','').split('\n')):
    sheet5.write(r+1, 4, row)
sqlitecur.execute(answer_5)
for r,row in enumerate(sqlitecur):
    sheet5.write(r+1, 0, row[0])
    sheet5.write(r+1, 1, row[1])
    sheet5.write(r+1, 2, row[2])

sheet6.write(0,0, 'Answer: ', style=style)
for r, row in enumerate(answer_6.split('\n')):
    sheet6.write(r+1,0, row)
print('Exporting answers . . .')
print('Done . . .')
answerbook.save('Shopee_SQL_Test_Answers.xls')
user_input=input('Please open Shopee_SQL_Test_Answers.xls and press "Enter" to quit the program')
if len(user_input) > 0: quit()
else: quit()
