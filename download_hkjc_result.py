import requests
from bs4 import BeautifulSoup
import xlsxwriter
from datetime import date, timedelta


year = 2024
# export result to excel
# Create a new Excel file
workbook = xlsxwriter.Workbook(f'raceresult_{year}.xlsx')

# Add a worksheet
worksheet = workbook.add_worksheet()

# Add a bold format to use to highlight cells
bold = workbook.add_format({'bold': True})

# Write the header row
worksheet.write('A1', '日期', bold)
worksheet.write('B1', '總場次', bold)
worksheet.write('C1', '馬場', bold)
worksheet.write('D1', '場次', bold)
worksheet.write('E1', '班次', bold)
worksheet.write('F1', '路程', bold)
worksheet.write('G1', '賽事名稱', bold)
worksheet.write('H1', '賽道', bold)
worksheet.write('I1', '賽事時間1', bold)
worksheet.write('J1', '賽事時間2', bold)
worksheet.write('K1', '賽事時間3', bold)
worksheet.write('L1', '賽事時間4', bold)
worksheet.write('M1', '賽事時間5', bold)
worksheet.write('N1', '賽事時間6', bold)
worksheet.write('O1', '名次', bold)
worksheet.write('P1', '馬號', bold)
worksheet.write('Q1', '馬名', bold)
worksheet.write('R1', '布號', bold)
worksheet.write('S1', '騎師', bold)
worksheet.write('T1', '練馬師', bold)
worksheet.write('U1', '實際負磅', bold)
worksheet.write('V1', '排位體重', bold)
worksheet.write('W1', '檔位', bold)
worksheet.write('X1', '頭馬距離', bold)
worksheet.write('Y1', '沿途走位', bold)
worksheet.write('Z1', '完成時間', bold)
worksheet.write('AA1', '獨贏賠率', bold)

# Start from the first cell below the headers
row = 1
col = 0


## create date range from 2023/09/01 to 2024/07/14
date_from = date(2024, 11, 3)
date_to = date(2024, 11, 3)


# Loop through each date
for date in [date_from + timedelta(days=x) for x in range((date_to-date_from).days + 1)]:
    # Convert date
    race_date = date.strftime('%Y/%m/%d')

    # URL to scrape HKJC racing results
    for race_course in ['ST', 'HV']:
        
        for race_no in range(1, 12):

            print(race_date, race_course, race_no)

            # URL to scrape HKJC racing results
            url = f'https://racing.hkjc.com/racing/information/Chinese/Racing/LocalResults.aspx?RaceDate={race_date}&Racecourse={race_course}&RaceNo={race_no}' 

            # Make a GET request to the URL
            response = requests.get(url)

            try:

                # Parse the HTML using BeautifulSoup
                soup = BeautifulSoup(response.text, 'html.parser')


                race_div = soup.find('div', class_='race_tab')
                tr = race_div.find_all('tr')

                # Extract race_no, race_index
                race_index = tr[0].find_all('td')[0].text.split('(')[1].replace(')', '')

                # Extract class, distance, and rating
                race_class = tr[2].find_all('td')[0].text.split('-')[0]

                # Safely extract distance
                distance_parts = tr[2].find_all('td')[0].text.split('-')
                if len(distance_parts) > 1:
                    distance = distance_parts[1].split('米')[0].strip()
                else:
                    distance = ''

                # Safely extract rating
                rating_parts = distance_parts[1].split('(') if len(distance_parts) > 1 else []
                if len(rating_parts) > 1:
                    rating = rating_parts[1].replace(')', '').strip()
                else:
                    rating = ''

                condition = tr[2].find_all('td')[2].text


                # Extract class, distance, and rating
                cup = tr[3].find_all('td')[0].text

                band = tr[3].find_all('td')[2].text

                time_1 = tr[4].find_all('td')[2].text.replace('(', '').replace(')', '')
                time_2 = tr[4].find_all('td')[3].text.replace('(', '').replace(')', '')
                time_3 = tr[4].find_all('td')[4].text.replace('(', '').replace(')', '')
                if len(tr[4].find_all('td'))>5:
                    time_4 = tr[4].find_all('td')[5].text.replace('(', '').replace(')', '')
                else:
                    time_4 = None
                if len(tr[4].find_all('td'))>6:
                    time_5 = tr[4].find_all('td')[6].text.replace('(', '').replace(')', '')
                else:
                    time_5 = None
                if len(tr[4].find_all('td'))>7:
                    time_6 = tr[4].find_all('td')[7].text.replace('(', '').replace(')', '')
                else:
                    time_6 = None


                split_time_1 = tr[5].find_all('td')[2].text.replace('(', '').replace(')', '')
                split_time_2 = tr[5].find_all('td')[3].text.replace('(', '').replace(')', '')
                split_time_3 = tr[5].find_all('td')[4].text.replace('(', '').replace(')', '')
                if len(tr[5].find_all('td'))>5:
                    split_time_4 = tr[5].find_all('td')[5].text.replace('(', '').replace(')', '')
                else:
                    split_time_4 = None
                if len(tr[5].find_all('td'))>6:
                    split_time_5 = tr[5].find_all('td')[6].text.replace('(', '').replace(')', '')
                else:
                    split_time_5 = None
                if len(tr[5].find_all('td'))>7:
                    split_time_6 = tr[5].find_all('td')[7].text.replace('(', '').replace(')', '')
                else:
                    split_time_6 = None
                # print(split_time_1, split_time_2, split_time_3, split_time_4, split_time_5, split_time_6)


                # Find the table containing the race results
                results_div = soup.find('div', class_='performance')


                # Loop through each row in the table
                for tr in results_div.find_all('tr')[1:]:
                    
                    # Get data from each cell
                    cells = tr.find_all('td')
                    
                    worksheet.write(row, col, race_date)
                    worksheet.write(row, col + 1, f"{year}{race_index.zfill(3)}")
                    worksheet.write(row, col + 2, race_course)
                    worksheet.write(row, col + 3, int(race_no))
                    worksheet.write(row, col + 4, race_class)
                    worksheet.write(row, col + 5, int(distance))
                    worksheet.write(row, col + 6, cup)
                    worksheet.write(row, col + 7, band)
                    worksheet.write(row, col + 8, time_1)
                    worksheet.write(row, col + 9, time_2)
                    worksheet.write(row, col + 10, time_3)
                    worksheet.write(row, col + 11, time_4)
                    worksheet.write(row, col + 12, time_5)
                    worksheet.write(row, col + 13, time_6)               
                    
                    position = cells[0].text.replace('\r', '').replace('\n', '')
                    try:     
                        worksheet.write(row, col + 14, int(position))
                    except:
                        worksheet.write(row, col + 14, None)
                    
                    horse_no = cells[1].text.replace('\r', '').replace('\n', '')
                    try:
                        worksheet.write(row, col + 15, int(horse_no))
                    except:
                        worksheet.write(row, col + 15, None)

                    worksheet.write(row, col + 16, cells[2].text.replace('\xa0', '').split('(')[0])
                    worksheet.write(row, col + 17, cells[2].text.replace('\xa0', '').split('(')[1].replace(')', ''))
                    worksheet.write(row, col + 18, cells[3].text.replace('\r', '').replace('\n', ''))
                    worksheet.write(row, col + 19, cells[4].text.replace('\r', '').replace('\n', ''))
                    
                    
                    actual_weight = cells[5].text.replace('\r', '').replace('\n', '')
                    try:
                        worksheet.write(row, col + 20, int(actual_weight))
                    except:
                        worksheet.write(row, col + 20, None)
                    

                    horse_weight = cells[6].text.replace('\r', '').replace('\n', '')
                    try:
                        worksheet.write(row, col + 21, int(horse_weight))
                    except:
                        worksheet.write(row, col + 21, None)
                    
                    draw_no = cells[7].text.replace('\r', '').replace('\n', '')
                    try:
                        worksheet.write(row, col + 22, int(draw_no))
                    except:
                        worksheet.write(row, col + 22, None)
                    
                    worksheet.write(row, col + 23, cells[8].text.replace('\r', '').replace('\n', ''))
                    worksheet.write(row, col + 24, cells[9].text.replace('\r', '').replace('\n', ''))
                    worksheet.write(row, col + 25, cells[10].text.replace('\r', '').replace('\n', ''))
                    
                    odds = cells[11].text.replace('\r', '').replace('\n', '')
                    try:
                        worksheet.write(row, col + 26, float(odds))
                    except:
                        worksheet.write(row, col + 26, None)
                    row += 1


            except:            
                pass



# Close the Excel file
workbook.close()

print('Excel file created successfully!')