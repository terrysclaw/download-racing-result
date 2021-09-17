import datetime
from time import sleep
from selenium import webdriver
import mysql.connector


config = {
    'user': 'hkjc',
    'password': 'Fr5cypE3D%',
    'host': 'mysql.ixx.cc',
    'database': 'hkjc',
    'raise_on_warnings': True,
    'buffered': True,
}


cnx = mysql.connector.connect(**config)
cursor = cnx.cursor()


driver = webdriver.Chrome("chromedriver.exe")

driver.get("http://win.on.cc/ftpdata/marksix/2020.txt")

sleep(3)

content = driver.find_element_by_xpath("/html/body/pre")


lines = content.text.split("\n")

for line in lines:
    columns = line.split(chr(32))

    draw_no = columns[0][2:4] + "-" + columns[0][4:7]


    try:
        add_res = ("INSERT INTO MarkSix(DrawNo) VALUES(%(draw_no)s)")
        data = {
            'draw_no': draw_no,
        }
        cursor.execute(add_res, data)
        cnx.commit()
    except:
        print(draw_no, "existed!")

    
	
    update_res = ("UPDATE MarkSix SET "
        "Date = %(draw_date)s, "
        "DrawnNo1 = %(drawn_no1)s, "
        "DrawnNo2 = %(drawn_no2)s, "
        "DrawnNo3 = %(drawn_no3)s, "
        "DrawnNo4 = %(drawn_no4)s, "
        "DrawnNo5 = %(drawn_no5)s, "
        "DrawnNo6 = %(drawn_no6)s, "
        "ExtraNum = %(extra_no)s "
        "WHERE DrawNo = %(draw_no)s")
        
    data = {
        'draw_date': datetime.datetime.strptime(columns[1], '%Y/%m/%d'),
        'drawn_no1': int(columns[2]),
        'drawn_no2': int(columns[3]),
        'drawn_no3': int(columns[4]),
        'drawn_no4': int(columns[5]),
        'drawn_no5': int(columns[6]),
        'drawn_no6': int(columns[7]),
        'extra_no': int(columns[8]),
        'draw_no': draw_no,
    }

    cursor.execute(update_res, data)
    cnx.commit()	
        
cursor.close()
cnx.close()

sleep(3)
driver.quit()
