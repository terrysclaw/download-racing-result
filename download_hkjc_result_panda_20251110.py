import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import date
import time
from typing import Dict, List, Optional, Tuple
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class HKJCResultDownloader:
    def __init__(self, year: int):
        self.year = year
        self.base_url = "https://racing.hkjc.com/racing/information/Chinese/Racing"
        self.df = self._load_existing_data()
        
    def _load_existing_data(self) -> pd.DataFrame:
        """Load existing Excel file or create new DataFrame."""
        load_file = f'raceresult_{self.year}.xlsx'
        try:
            df = pd.read_excel(load_file)
            logger.info(f"Loaded existing file: {load_file}")
            # 23.04_x000D_
            for i in range(1, 7):
                col_name = f'第 {i} 段'
                if col_name in df.columns:
                    df[col_name] = df[col_name].astype(str).str.replace('nan', '').str.strip()
            return df
        except FileNotFoundError:
            logger.info(f"No existing file found. Creating new DataFrame.")
            return pd.DataFrame({})
    
    def _get_race_dates(self) -> List[Dict]:
        """Get race dates based on year."""
        dates_2024 = [
            {'date': date(2024, 9, 8), 'course': 'ST'},
            {'date': date(2024, 9, 11), 'course': 'HV'},
            # ... (keep all your existing dates)
            {'date': date(2025, 7, 16), 'course': 'HV'},
        ]
        
        dates_2025 = [
            # {'date': date(2026, 2, 1), 'course': 'ST'},
            # {'date': date(2026, 2, 4), 'course': 'HV'},
            # {'date': date(2026, 2, 8), 'course': 'ST'},
            # {'date': date(2026, 2, 11), 'course': 'HV'},
            # {'date': date(2026, 2, 14), 'course': 'ST'},
            {'date': date(2026, 2, 19), 'course': 'ST'},
            # Add more 2025 dates as needed
        ]
        
        return dates_2024 if self.year == 2024 else dates_2025
    
    def _initialize_data_structure(self) -> Dict[str, List]:
        """Initialize the data structure for race results."""
        return {
            '日期': [], '總場次': [], '馬場': [], '場次': [], '班次': [], '路程': [],
            '場地狀況': [], '賽事名稱': [], '泥草': [], '賽道': [],
            '賽事時間1': [], '賽事時間2': [], '賽事時間3': [], '賽事時間4': [], '賽事時間5': [], '賽事時間6': [],
            '名次': [], '馬號': [], '馬名': [], '布號': [], '騎師': [], '練馬師': [],
            '實際負磅': [], '排位體重': [], '檔位': [], '頭馬距離': [], '沿途走位': [], '完成時間': [], '獨贏賠率': [],
            '第 1 段': [], '第 2 段': [], '第 3 段': [], '第 4 段': [], '第 5 段': [], '第 6 段': [],
        }
    
    def _safe_int_convert(self, value: str) -> Optional[int]:
        """Safely convert string to integer."""
        try:
            return int(value.replace('\r', '').replace('\n', '').strip())
        except (ValueError, AttributeError):
            return None
    
    def _safe_float_convert(self, value: str) -> Optional[float]:
        """Safely convert string to float."""
        try:
            return float(value.replace('\r', '').replace('\n', '').strip())
        except (ValueError, AttributeError):
            return None
    
    def _clean_text(self, text: str) -> str:
        """Clean text by removing unwanted characters."""
        return text.replace('\r', '').replace('\n', '').replace('\xa0', '').strip()
    
    def _extract_race_info(self, soup: BeautifulSoup) -> Tuple[str, str, str, str, str, str, List[str], List[str]]:
        """Extract race information from soup."""
        race_div = soup.find('div', class_='race_tab')
        tr = race_div.find_all('tr')
        
        # Extract race index
        race_index = tr[0].find_all('td')[0].text.split('(')[1].replace(')', '')
        
        # Extract class and distance
        race_info = tr[2].find_all('td')[0].text
        race_class = race_info.split('-')[0]
        
        distance_parts = race_info.split('-')
        distance = distance_parts[1].split('米')[0].strip() if len(distance_parts) > 1 else ''
        
        # Extract condition, cup, and band
        condition = tr[2].find_all('td')[2].text
        cup = tr[3].find_all('td')[0].text
        band = tr[3].find_all('td')[2].text
        
        # Extract race times
        time_cells = tr[4].find_all('td')[2:]
        times = [time_cells[i].text.replace('(', '').replace(')', '') if i < len(time_cells) else None 
                for i in range(6)]
        
        # Extract split times (not used in current implementation but kept for future use)
        split_time_cells = tr[5].find_all('td')[2:]
        split_times = [split_time_cells[i].text.replace('(', '').replace(')', '') if i < len(split_time_cells) else None 
                      for i in range(6)]
        
        return race_index, race_class, distance, condition, cup, band, times, split_times
    
    def _extract_horse_data(self, cells: List, race_data: Dict, data: Dict[str, List]) -> None:
        """Extract individual horse data from table cells."""
        # Basic race information
        for key, value in race_data.items():
            data[key].append(value)
        
        # Horse-specific data
        data['名次'].append(self._safe_int_convert(cells[0].text))
        data['馬號'].append(self._safe_int_convert(cells[1].text))
        
        # Horse name and cloth number
        horse_info = self._clean_text(cells[2].text)
        if '(' in horse_info:
            data['馬名'].append(horse_info.split('(')[0])
            data['布號'].append(horse_info.split('(')[1].replace(')', ''))
        else:
            data['馬名'].append(horse_info)
            data['布號'].append('')
        
        data['騎師'].append(self._clean_text(cells[3].text))
        data['練馬師'].append(self._clean_text(cells[4].text))
        data['實際負磅'].append(self._safe_int_convert(cells[5].text))
        data['排位體重'].append(self._safe_int_convert(cells[6].text))
        data['檔位'].append(self._safe_int_convert(cells[7].text))
        data['頭馬距離'].append(self._clean_text(cells[8].text))
        data['沿途走位'].append(self._clean_text(cells[9].text))
        data['完成時間'].append(self._clean_text(cells[10].text))
        data['獨贏賠率'].append(self._safe_float_convert(cells[11].text))
        
        # Initialize sectional times (will be filled later)
        for i in range(1, 7):
            data[f'第 {i} 段'].append('')
    
    def _get_sectional_times(self, race_date: date, race_no: int, data: Dict[str, List]) -> None:
        """Fetch and populate sectional times."""
        url = f"{self.base_url}/DisplaySectionalTime.aspx?RaceDate={race_date.strftime('%d/%m/%Y')}&RaceNo={race_no}"
        
        try:
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.text, 'html.parser')
            race_table = soup.find('table', class_='table_bd')
            
            if not race_table:
                logger.warning(f"Sectional time table not found for {race_date} Race {race_no}")
                return
            
            # Extract sectional time data
            tbody = race_table.find('tbody')
            if not tbody:
                return
                
            for row in tbody.find_all('tr'):
                cells = [cell.text.strip() for cell in row.find_all('td')]
                if len(cells) < 3:
                    continue
                    
                try:
                    cloth_number = cells[2].split('(')[1].replace(')', '')
                except IndexError:
                    continue
                
                # Find matching horse in data
                for index, item in enumerate(data['布號']):
                    if item == cloth_number:
                        # Update sectional times
                        for section in range(1, 7):
                            if section + 2 < len(cells):
                                try:
                                    sectional_time = cells[section + 2].split('\n')[1] if '\n' in cells[section + 2] else cells[section + 2]
                                    # 23.04_x000D_n -> 23.04
                                    sectional_time = sectional_time.replace('_x000D_n', '').strip()
                                    data[f'第 {section} 段'][index] = sectional_time
                                except (IndexError, AttributeError):
                                    pass
                        break
                        
        except requests.RequestException as e:
            logger.error(f"Error fetching sectional time for {race_date} Race {race_no}: {e}")
        except Exception as e:
            logger.error(f"Unexpected error processing sectional time for {race_date} Race {race_no}: {e}")
    
    def _scrape_race_results(self, race_date: str, race_course: str, race_no: int) -> Optional[Dict[str, List]]:
        """Scrape race results for a specific race."""
        url = f"{self.base_url}/LocalResults.aspx?RaceDate={race_date}&Racecourse={race_course}&RaceNo={race_no}"
        
        try:
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Extract race information
            race_index, race_class, distance, condition, cup, band, times, _ = self._extract_race_info(soup)
            
            # Determine track type
            if '草地' in band:
                track_type = '草地'
                track = band.split(' - ')[1] if ' - ' in band else None
            else:
                track_type = '泥地'
                track = None
            
            # Prepare race data template
            race_data = {
                '日期': race_date,
                '總場次': f"{self.year}{race_index.zfill(3)}",
                '馬場': race_course,
                '場次': int(race_no),
                '班次': race_class,
                '路程': self._safe_int_convert(distance),
                '場地狀況': condition,
                '賽事名稱': cup,
                '泥草': track_type,
                '賽道': track,
                **{f'賽事時間{i+1}': times[i] for i in range(6)}
            }
            
            # Initialize data structure
            data = self._initialize_data_structure()
            
            # Extract horse results
            results_div = soup.find('div', class_='performance')
            if not results_div:
                logger.warning(f"No performance data found for {race_date} Race {race_no}")
                return None
            
            for tr in results_div.find_all('tr')[1:]:  # Skip header row
                cells = tr.find_all('td')
                if len(cells) >= 12:  # Ensure we have enough columns
                    self._extract_horse_data(cells, race_data, data)
            
            return data
            
        except requests.RequestException as e:
            logger.error(f"Error fetching race results for {race_date} Race {race_no}: {e}")
            return None
        except Exception as e:
            logger.error(f"Unexpected error processing race {race_date} Race {race_no}: {e}")
            return None
    
    def download_all_results(self) -> None:
        """Download all race results for the specified year."""
        race_dates = self._get_race_dates()
        
        for date_info in race_dates:
            race_date = date_info['date'].strftime('%Y/%m/%d')
            race_course = date_info['course']
            
            for race_no in range(1, 12):
                logger.info(f"Processing {race_date} {race_course} Race {race_no}")
                
                # Scrape race results
                race_data = self._scrape_race_results(race_date, race_course, race_no)
                
                if race_data:
                    # Get sectional times
                    self._get_sectional_times(date_info['date'], race_no, race_data)
                    
                    # Append to main dataframe
                    self.df = pd.concat([self.df, pd.DataFrame(race_data)], ignore_index=True)
                
                # Add small delay to be respectful to the server
                time.sleep(1)
    
    def save_results(self) -> None:
        """Save results to Excel file."""
        output_file = f'raceresult_{self.year}.xlsx'
        self.df.to_excel(output_file, index=False)
        logger.info(f"Results saved to {output_file}")

def main():
    year = 2025
    downloader = HKJCResultDownloader(year)
    downloader.download_all_results()
    downloader.save_results()

if __name__ == "__main__":
    main()