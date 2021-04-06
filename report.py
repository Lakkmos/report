import pyodbc,datetime,smtplib, xlwt
from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase

users = {}
delta = {}
log = open('log.txt', 'a')

try:
    f = open('conf.txt', 'r')
    f = f.read()
    if f[0] == str(1):
        temp = f[f.find('deep_day'):]
        temp = temp[9:temp.find('\n')]
        V = int(temp)
        time_d2 = str(datetime.datetime.today().strftime('%d.%m.%y 00:00'))
        time_d1 = datetime.datetime.now() - datetime.timedelta(days=V)
        time_d1 = str(time_d1.strftime('%d.%m.%y 00:00'))
    elif f[0] == str(2):
        time_d1 = f[f.find('time1:'):]
        time_d1 = time_d1[6:time_d1.find('\n')]
        time_d2 = f[f.find('time2:'):]
        time_d2 = time_d2[6:time_d2.find('\n')]
    else:
        time_d2 = str(datetime.datetime.today().strftime('%d.%m.%y 00:00'))
        time_d1 = datetime.datetime.now() - datetime.timedelta(days=1)
        time_d1 = str(time_d1.strftime('%d.%m.%y 00:00'))
    send = f[f.find('send'):]
    send = send[5:send.find('\n')]
    addr_to = f[f.find('addr_to'):]
    addr_to = addr_to[8:addr_to.find('\n')]
    addr_from = f[f.find('addr_from'):]
    addr_from = addr_from[10:addr_from.find('\n')]
    password = f[f.find('password'):]
    password = password[9:password.find('\n')]
    serv_sql = f[f.find('|')+1:]
    serv_sql = serv_sql[:serv_sql.find('|')]
    log.write(datetime.datetime.today().strftime('%Y.%m.%d %H:%M') +' --- '+'Использованы настройки из файла'+'\n')
except:
    time_d2 = str(datetime.datetime.today().strftime('%d.%m.%y 00:00'))
    time_d1 = datetime.datetime.now() - datetime.timedelta(days=1)
    time_d1 = str(time_d1.strftime('%d.%m.%y 00:00'))
    send = yes
    serv_sql = "DRIVER={SQL Server}; SERVER=\SQLSERVER2012; DATABASE=; UID=; PWD="
    log.write(datetime.datetime.today().strftime('%Y.%m.%d %H:%M') + ' --- ' + 'Использованы настройки по умолчанию. Отчет за сутки и ярославский сервер' + '\n')
log.write(datetime.datetime.today().strftime('%Y.%m.%d %H:%M') + ' --- ' + serv_sql + '\n')
cnxn = pyodbc.connect(serv_sql)
cursor = cnxn.cursor()
cursor2 = cnxn.cursor()


vir1 = "SELECT DeviceTime, Remark, HozOrgan FROM [dbo].[pLogData] WHERE TimeVal BETWEEN (convert(datetime,'"+ time_d1 +"')) AND (convert(datetime,'"+ time_d2 +"'))ORDER BY DeviceTime ASC"

vir2 = "SELECT Owner, OwnerName FROM [dbo].[pMark]"

for row in cursor.execute(vir2):
    users[row[0]] = row[1]
log.write(datetime.datetime.today().strftime('%Y.%m.%d %H:%M') +' --- '+'Подключение к серверу'+'\n')
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('1', cell_overwrite_ok=True)

i=1
same1=0
same2=0
same3=0
worksheet.write(0, 0, 'ФИО')
worksheet.write(0, 1, 'Дата')
worksheet.write(0, 2, 'Время')
worksheet.write(0, 3, 'Датчик')
worksheet.write(0, 4, 'ИД оператора')
for row in cursor.execute(vir1):

    if row[2]!=0 and row[1]!=None and row[0]!=None :
        try:
            user = users[row[2]]
        except:
            user = 'Не найден'
        #print(user, row[0].strftime('%Y.%m.%d %H:%M'),row[1:])
        if ((same1 == row[0]) and (same2 == row[2]) and (same3 == row[1])):
            i=i
        else:
            worksheet.write(i, 0, user)
            worksheet.write(i, 1, row[0].strftime('%Y.%m.%d'))
            worksheet.write(i, 2, row[0].strftime('%H:%M'))
            same1 = row[0]
            same2 = row[2]
            same3 = row[1]
            if user not in delta.keys():
                delta[user] = []
            delta[user].append(row[0])
            worksheet.write(i, 3, row[1])
            worksheet.write(i, 4, row[2])
            i = i + 1
worksheet.col(0).width = 256 * 46
worksheet.col(1).width = 256 * 11
worksheet.col(4).width = 256 * 8
worksheet.col(3).width = 256 * 48
log.write(datetime.datetime.today().strftime('%Y.%m.%d %H:%M') +' --- '+'Сформирован первый лист'+'\n')

i=1
worksheet = workbook.add_sheet('2', cell_overwrite_ok=True)
worksheet.write(0, 0, 'ФИО')
worksheet.write(0, 1, 'Первый выход')
worksheet.write(0, 2, 'Последний вход')
for d in delta:
    worksheet.write(i, 0, d)
    delta[d].sort()
    worksheet.write(i, 1, delta[d][0].strftime('%Y.%m.%d %H:%M'))
    worksheet.write(i, 2, delta[d][-1].strftime('%Y.%m.%d %H:%M'))
    i = i + 1
worksheet.col(0).width = 256 * 46
worksheet.col(1).width = 256 * 15
worksheet.col(2).width = 256 * 15
log.write(datetime.datetime.today().strftime('%Y.%m.%d %H:%M') +' --- '+'Сформирован второй лист'+'\n')

try:
    workbook.save('Report.xls')

    if send == 'yes':
        msg = MIMEMultipart()
        msg['From'] = addr_from
        msg['To'] = addr_to
        msg['Subject'] = "Отчет по входу выходу "+ str(time_d1)+ ' - '+str(time_d2)
        body = "_"
        msg.attach(MIMEText(body, 'plain'))
        filename = "Report.xls"
        attachment = open("Report.xls", "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        msg.attach(part)
        print(addr_from, password)
        #try:
        server = smtplib.SMTP('smtp.yandex.ru', 587)
        server.starttls()
        server.login(addr_from, password)  # Получаем доступ
        text = msg.as_string()
        server.sendmail(addr_from, addr_to.split(','), text)
        server.quit()
        log.write(datetime.datetime.today().strftime('%Y.%m.%d %H:%M') + ' --- ' + 'Отправлены письма удачно' + '\n')
        # except:
        log.write(datetime.datetime.today().strftime(
            '%Y.%m.%d %H:%M') + ' --- ' + 'Письма не отправлены, ошибка подключения/отправки' + '\n')
    else:
        log.write(datetime.datetime.today().strftime('%Y.%m.%d %H:%M') + ' --- ' + 'Письма не отправлены, согласно настройкам' + '\n')
except:
    print('Не получилось совсем')

log.write(datetime.datetime.today().strftime('%Y.%m.%d %H:%M') +' --- '+'Скрипт отработал правильно'+'\n'+'\n')
