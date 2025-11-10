import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import date
import numpy as np
import logging
from typing import Dict, List, Optional, Tuple
from pathlib import Path

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class HKJCRacecardDownloader:
    def __init__(self, race_date: date, race_course: str):
        self.race_date = race_date
        self.race_course = race_course
        self.base_url = "https://racing.hkjc.com/racing/information/Chinese/Racing"
        self.df_results = self._load_historical_data()
        
    def _load_historical_data(self) -> pd.DataFrame:
        """Load historical race results."""
        try:
            df_2024 = pd.read_excel('raceresult_2024.xlsx')
            df_2025 = pd.read_excel('raceresult_2025.xlsx')
            return pd.concat([df_2024, df_2025], ignore_index=True)
        except FileNotFoundError as e:
            logger.error(f"Historical data file not found: {e}")
            return pd.DataFrame()
    
    def _initialize_data_structure(self) -> Dict[str, List]:
        """Initialize the data structure for racecard."""
        return {
            '日期': [], '馬場': [], '場次': [], '泥草': [], '賽道': [], '班次': [], '路程': [],
            '馬匹編號': [], '6次近績': [], '馬名': [], '網址': [], '烙號': [], '負磅': [],
            '騎師': [], '可能超磅': [], '檔位': [], '練馬師': [], '國際評分': [], '評分': [],
            '評分+/-': [], '排位體重': [], '排位體重+/-': [], '最佳時間': [], '馬齡': [],
            '分齡讓磅': [], '性別': [], '今季獎金': [], '優先參賽次序': [], '上賽距今日數': [],
            '配備': [], '馬主': [], '父系': [], '母系': [], '進口類別': [],
        }
    
    def _safe_int_convert(self, value: str) -> Optional[int]:
        """Safely convert string to integer."""
        try:
            return int(value.strip())
        except (ValueError, AttributeError):
            return None
    
    def _clean_text(self, text: str) -> str:
        """Clean text by removing unwanted characters."""
        return text.replace('\r', '').replace('\n', '').replace('\xa0', '').strip()
    
    def _extract_race_info(self, soup: BeautifulSoup) -> Tuple[Optional[str], Optional[str], Optional[int]]:
        """Extract race information from soup."""
        race_div = soup.find('div', class_='f_fs13')
        
        if not race_div:
            return None, None, None
            
        band = None
        race_class = None
        distance = None
        
        for col in race_div.text.split(','):
            if '賽道' in col:
                band = col.strip()
            if '班' in col:
                race_class = col.strip()
            if '米' in col:
                try:
                    distance = self._safe_int_convert(col.split('米')[0])
                except:
                    distance = None
                    
        return band, race_class, distance
    
    def _extract_horse_data(self, cells: List, race_info: Dict, data: Dict[str, List]) -> None:
        """Extract individual horse data from table cells."""
        # Add race information
        for key, value in race_info.items():
            data[key].append(value)
        
        # Extract horse-specific data with safe conversion
        data['馬匹編號'].append(self._safe_int_convert(cells[0].text))
        data['6次近績'].append(self._clean_text(cells[1].text) if len(cells) > 1 else None)
        
        # Horse name and URL
        if len(cells) > 3:
            data['馬名'].append(self._clean_text(cells[3].text))
            try:
                data['網址'].append("https://racing.hkjc.com" + cells[3].find('a')['href'])
            except:
                data['網址'].append(None)
        else:
            data['馬名'].append(None)
            data['網址'].append(None)
        
        # Continue with other fields
        fields_mapping = [
            ('烙號', 4), ('負磅', 5), ('騎師', 6), ('可能超磅', 7), ('檔位', 8),
            ('練馬師', 9), ('國際評分', 10), ('評分', 11), ('評分+/-', 12),
            ('排位體重', 13), ('排位體重+/-', 14), ('最佳時間', 15), ('馬齡', 16),
            ('分齡讓磅', 17), ('性別', 18), ('今季獎金', 19), ('優先參賽次序', 20),
            ('上賽距今日數', 21), ('配備', 22), ('馬主', 23), ('父系', 24),
            ('母系', 25), ('進口類別', 26)
        ]
        
        for field_name, index in fields_mapping:
            if index < len(cells):
                if field_name in ['負磅', '檔位', '國際評分', '評分', '排位體重', '馬齡', '分齡讓磅']:
                    data[field_name].append(self._safe_int_convert(cells[index].text))
                else:
                    data[field_name].append(self._clean_text(cells[index].text))
            else:
                data[field_name].append(None)
    
    def _get_time_adjustment_factors(self, venue: str, surface: str, distance: int) -> Dict[str, float]:
        """Get time adjustment factors based on venue, surface, and distance."""
        factors = {'weight_early': 0, 'weight_late': 0, 'gate': 0}
        
        if venue == 'HV':
            distance_factors = {
                1000: {'weight_early': 0.015, 'weight_late': 0.056, 'gate': 0.02},
                1200: {'weight_early': 0.017, 'weight_late': 0.056, 'gate': 0.033},
                1650: {'weight_early': 0.019, 'weight_late': 0.056, 'gate': 0.043},
                1800: {'weight_early': 0.021, 'weight_late': 0.056, 'gate': 0.048},
                2200: {'weight_early': 0.023, 'weight_late': 0.056, 'gate': 0.053}
            }
        else:  # ST
            if surface == '草地':
                distance_factors = {
                    1000: {'weight_early': 0.015, 'weight_late': 0.056, 'gate': -0.019},
                    1200: {'weight_early': 0.017, 'weight_late': 0.056, 'gate': 0.028},
                    1400: {'weight_early': 0.018, 'weight_late': 0.056, 'gate': 0.033},
                    1600: {'weight_early': 0.019, 'weight_late': 0.056, 'gate': 0.038},
                    1800: {'weight_early': 0.021, 'weight_late': 0.056, 'gate': 0.043},
                    2000: {'weight_early': 0.022, 'weight_late': 0.056, 'gate': 0.048},
                    2400: {'weight_early': 0.024, 'weight_late': 0.056, 'gate': 0.053}
                }
            else:  # 泥地
                distance_factors = {
                    1200: {'weight_early': 0.017, 'weight_late': 0.056, 'gate': 0.028},
                    1650: {'weight_early': 0.019, 'weight_late': 0.056, 'gate': 0.043},
                    1800: {'weight_early': 0.021, 'weight_late': 0.056, 'gate': 0.043}
                }
        
        return distance_factors.get(distance, factors)
    
    def _calculate_time_adjustment(self, row: pd.Series) -> float:
        """Calculate time adjustment based on weight and gate changes."""
        if pd.isna(row['上次負磅 +/-']) or pd.isna(row['上次檔位 +/-']):
            return 0
        
        factors = self._get_time_adjustment_factors(
            row['馬場'], row['泥草'], row['路程']
        )
        
        distance = row['路程']
        weight_diff = row['上次負磅 +/-']
        gate_diff = row['上次檔位 +/-']
        
        # Calculate weight adjustment (early and late portions)
        early_portion = distance - 400 if distance > 400 else 0
        late_portion = min(distance, 400)
        
        time_adjustment = 0
        time_adjustment += (weight_diff * factors['weight_early'] / distance * early_portion)
        time_adjustment += (weight_diff * factors['weight_late'] / distance * late_portion)
        time_adjustment += gate_diff * factors['gate']
        
        return time_adjustment
    
    def _apply_ground_adjustment(self, time_value: float, ground_condition: str, venue: str) -> float:
        """Apply ground condition adjustment to time."""
        if ground_condition == '好地至快地':
            multiplier = 1.006 if venue == 'HV' else 1.004
            return time_value * multiplier
        elif ground_condition == '好地至黏地':
            return time_value / 1.012
        return time_value
    
    def _calculate_sectional_adjustments(self, row: pd.Series) -> Optional[float]:
        """Calculate adjusted last 800m time."""
        if pd.isna(row['上次最後 800']) or pd.isna(row['上次負磅 +/-']) or pd.isna(row['上次檔位 +/-']):
            return None
        
        time_1 = row['上次彎位 400'] if pd.notna(row['上次彎位 400']) else 0
        time_2 = row['上次最後 400'] if pd.notna(row['上次最後 400']) else 0
        
        factors = self._get_time_adjustment_factors(
            row['馬場'], row['泥草'], row['路程']
        )
        
        distance = row['路程']
        weight_diff = row['上次負磅 +/-']
        gate_diff = row['上次檔位 +/-']
        
        # Adjust each 400m section
        time_1 += (weight_diff * factors['weight_early'] / distance * 400)
        time_1 += gate_diff * factors['gate'] / distance * 400
        time_2 += (weight_diff * factors['weight_late'] / distance * 400)
        
        adjusted_800 = time_1 + time_2
        
        # Apply ground condition adjustment
        if row['上次泥草'] == '草地':
            adjusted_800 = self._apply_ground_adjustment(
                adjusted_800, row['上次場地狀況'], row['上次馬場']
            )
        
        return round(adjusted_800, 3)
    
    def _process_historical_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """Process and add historical performance data."""
        # Initialize columns
        historical_columns = [
            '上次總場次', '上次日期', '上次馬場', '上次泥草', '上次場次', '上次班次', '上次路程',
            '上次場地狀況', '上次名次', '上次騎師', '上次賠率', '上次負磅', '上次負磅 +/-',
            '上次檔位', '上次檔位 +/-', '上次賽事時間1', '上次賽事時間2', '上次賽事時間3',
            '上次賽事時間4', '上次賽事時間5', '上次賽事時間6', '上次賽事時間', '上次完成時間',
            '上次第 1 段', '上次第 2 段', '上次第 3 段', '上次第 4 段', '上次第 5 段', '上次第 6 段',
            '上次最後 400', '上次最後 800', '上次頭段完成時間', '上次頭段', '上次彎位 400',
            '上次調整基數', '上次調整後最後 800', '上次調整後完成時間', '上次調整後完成秒速'
        ]
        
        for col in historical_columns:
            df[col] = None
        
        for index, row in df.iterrows():
            # Find last matching race
            if row['泥草'] == '草地':
                last_match = self.df_results[
                    (self.df_results['布號'] == row['烙號']) &
                    (self.df_results['馬場'] == row['馬場']) &
                    (self.df_results['泥草'] == row['泥草']) &
                    (~self.df_results['場地狀況'].isin(['軟地', '黏地'])) &
                    (self.df_results['路程'] == row['路程']) &
                    (self.df_results['獨贏賠率'] > 0)
                ].tail(1)
            else:
                last_match = self.df_results[
                    (self.df_results['布號'] == row['烙號']) &
                    (self.df_results['馬場'] == row['馬場']) &
                    (self.df_results['泥草'] == row['泥草']) &
                    (self.df_results['路程'] == row['路程']) &
                    (self.df_results['獨贏賠率'] > 0)
                ].tail(1)
            
            if not last_match.empty:
                self._populate_historical_data(df, index, row, last_match.iloc[0])
        
        return df
    
    def _populate_historical_data(self, df: pd.DataFrame, index: int, row: pd.Series, last_match: pd.Series) -> None:
        """Populate historical data for a horse."""
        # Basic historical data
        df.at[index, '上次總場次'] = last_match['總場次']
        df.at[index, '上次日期'] = last_match['日期']
        df.at[index, '上次馬場'] = last_match['馬場']
        df.at[index, '上次泥草'] = last_match['泥草']
        df.at[index, '上次場次'] = last_match['場次']
        df.at[index, '上次班次'] = last_match['班次']
        df.at[index, '上次路程'] = last_match['路程']
        df.at[index, '上次場地狀況'] = last_match['場地狀況']
        df.at[index, '上次名次'] = last_match['名次']
        df.at[index, '上次騎師'] = last_match['騎師']
        df.at[index, '上次賠率'] = last_match['獨贏賠率']
        df.at[index, '上次負磅'] = last_match['實際負磅']
        df.at[index, '上次檔位'] = last_match['檔位']
        
        # Calculate differences
        if pd.notna(last_match['實際負磅']):
            df.at[index, '上次負磅 +/-'] = row['負磅'] - last_match['實際負磅']
        
        if pd.notna(last_match['檔位']):
            df.at[index, '上次檔位 +/-'] = row['檔位'] - last_match['檔位']
        
        # Race times
        for i in range(1, 7):
            df.at[index, f'上次賽事時間{i}'] = last_match[f'賽事時間{i}']
            df.at[index, f'上次第 {i} 段'] = last_match[f'第 {i} 段']
        
        # Set appropriate race time based on distance
        distance = row['路程']
        if 1000 <= distance <= 1200:
            df.at[index, '上次賽事時間'] = last_match['賽事時間3']
        elif 1200 < distance <= 1650:
            df.at[index, '上次賽事時間'] = last_match['賽事時間4']
        elif 1650 < distance <= 2000:
            df.at[index, '上次賽事時間'] = last_match['賽事時間5']
        elif distance > 2000:
            df.at[index, '上次賽事時間'] = last_match['賽事時間6']
        
        df.at[index, '上次完成時間'] = last_match['完成時間']
        
        # Calculate sectional times
        self._calculate_sectional_times(df, index, last_match)
        
        # Calculate time adjustments
        if pd.notna(df.at[index, '上次完成時間']) and df.at[index, '上次完成時間'] != '---':
            self._calculate_adjusted_times(df, index, row)
    
    def _calculate_sectional_times(self, df: pd.DataFrame, index: int, last_match: pd.Series) -> None:
        """Calculate sectional time breakdowns."""
        # Find the last available sectional segment
        available_sections = []
        for i in range(1, 7):
            if pd.notna(last_match.get(f'第 {i} 段')):
                available_sections.append(i)
        
        if not available_sections:
            return
        
        last_section = max(available_sections)
        
        # Calculate sectional times based on available data
        if last_section >= 6:  # 6 sections available
            df.at[index, '上次頭段完成時間'] = last_match.get('賽事時間5')
            df.at[index, '上次頭段'] = round(sum(last_match.get(f'第 {j} 段', 0) for j in range(1, 5) if pd.notna(last_match.get(f'第 {j} 段'))), 2)
            df.at[index, '上次彎位 400'] = round(last_match.get('第 5 段', 0), 2)
            df.at[index, '上次最後 400'] = round(last_match.get('第 6 段', 0), 2)
            df.at[index, '上次最後 800'] = round((last_match.get('第 5 段', 0) + last_match.get('第 6 段', 0)), 2)
            
        elif last_section >= 5:  # 5 sections available
            df.at[index, '上次頭段完成時間'] = last_match.get('賽事時間4')
            df.at[index, '上次頭段'] = round(sum(last_match.get(f'第 {j} 段', 0) for j in range(1, 4) if pd.notna(last_match.get(f'第 {j} 段'))), 2)
            df.at[index, '上次彎位 400'] = round(last_match.get('第 4 段', 0), 2)
            df.at[index, '上次最後 400'] = round(last_match.get('第 5 段', 0), 2)
            df.at[index, '上次最後 800'] = round((last_match.get('第 4 段', 0) + last_match.get('第 5 段', 0)), 2)
            
        elif last_section >= 4:  # 4 sections available
            df.at[index, '上次頭段完成時間'] = last_match.get('賽事時間3')
            df.at[index, '上次頭段'] = round(sum(last_match.get(f'第 {j} 段', 0) for j in range(1, 3) if pd.notna(last_match.get(f'第 {j} 段'))), 2)
            df.at[index, '上次彎位 400'] = round(last_match.get('第 3 段', 0), 2)
            df.at[index, '上次最後 400'] = round(last_match.get('第 4 段', 0), 2)
            df.at[index, '上次最後 800'] = round((last_match.get('第 3 段', 0) + last_match.get('第 4 段', 0)), 2)
            
        elif last_section >= 3:  # 3 sections available
            df.at[index, '上次頭段完成時間'] = last_match.get('賽事時間2')
            df.at[index, '上次頭段'] = round(last_match.get('第 1 段', 0) + last_match.get('第 2 段', 0), 2) if pd.notna(last_match.get('第 1 段')) and pd.notna(last_match.get('第 2 段')) else round(last_match.get('第 1 段', 0), 2)
            df.at[index, '上次彎位 400'] = round(last_match.get('第 2 段', 0), 2)
            df.at[index, '上次最後 400'] = round(last_match.get('第 3 段', 0), 2)
            df.at[index, '上次最後 800'] = round((last_match.get('第 2 段', 0) + last_match.get('第 3 段', 0)), 2)
            
        elif last_section >= 2:  # Only 2 sections available
            df.at[index, '上次頭段完成時間'] = last_match.get('賽事時間1')
            df.at[index, '上次頭段'] = round(last_match.get('第 1 段', 0), 2)
            df.at[index, '上次彎位 400'] = round(last_match.get('第 1 段', 0), 2)
            df.at[index, '上次最後 400'] = round(last_match.get('第 2 段', 0), 2)
            df.at[index, '上次最後 800'] = round((last_match.get('第 1 段', 0) + last_match.get('第 2 段', 0)), 2)
            
        else:  # Only 1 section available
            df.at[index, '上次頭段完成時間'] = None
            df.at[index, '上次頭段'] = None
            df.at[index, '上次彎位 400'] = None
            df.at[index, '上次最後 400'] = round(last_match.get('第 1 段', 0), 2)
            df.at[index, '上次最後 800'] = None
        
        # Set adjustment base (調整基數) - this seems to be related to time adjustment
        df.at[index, '上次調整基數'] = self._calculate_time_adjustment(df.loc[index]) if hasattr(self, '_calculate_time_adjustment') else 0
    
    def _calculate_adjusted_times(self, df: pd.DataFrame, index: int, row: pd.Series) -> None:
        """Calculate time-adjusted performance metrics."""
        finish_time_str = df.at[index, '上次完成時間']
        
        try:
            # Convert finish time to seconds
            time_parts = finish_time_str.split(':')
            finish_seconds = int(time_parts[0]) * 60 + float(time_parts[1])
            
            # Apply time adjustment
            time_adjustment = self._calculate_time_adjustment(df.loc[index])
            adjusted_seconds = finish_seconds + time_adjustment
            
            # Apply ground condition adjustment
            if df.at[index, '上次泥草'] == '草地':
                adjusted_seconds = self._apply_ground_adjustment(
                    adjusted_seconds, df.at[index, '上次場地狀況'], df.at[index, '上次馬場']
                )
            
            # Store results
            df.at[index, '上次調整後完成秒速'] = adjusted_seconds
            df.at[index, '上次調整後完成時間'] = f"{int(adjusted_seconds // 60)}:{round(adjusted_seconds % 60, 2)}"
            
            # Calculate adjusted last 800m
            df.at[index, '上次調整後最後 800'] = self._calculate_sectional_adjustments(df.loc[index])
            
        except (ValueError, IndexError) as e:
            logger.warning(f"Error processing finish time for index {index}: {e}")
    
    def _scrape_race_data(self, race_no: int) -> Optional[pd.DataFrame]:
        """Scrape race data for a specific race."""
        url = f"{self.base_url}/RaceCard.aspx?RaceDate={self.race_date.strftime('%Y/%m/%d')}&Racecourse={self.race_course}&RaceNo={race_no}"
        
        try:
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Extract race information
            band, race_class, distance = self._extract_race_info(soup)
            
            race_info = {
                '日期': self.race_date,
                '馬場': self.race_course,
                '場次': race_no,
                '泥草': '草地' if band else '泥地',
                '賽道': band,
                '班次': race_class,
                '路程': distance
            }
            
            # Find racecard table
            racecard_table = soup.find('table', id='racecardlist')
            if not racecard_table:
                logger.warning(f"No racecard table found for Race {race_no}")
                return None
            
            # Extract horse data
            data = self._initialize_data_structure()
            tbody = racecard_table.find('tbody')
            
            if tbody:
                for tr in tbody.find_all('tr')[2:]:  # Skip header rows
                    cells = tr.find_all('td')
                    if len(cells) >= 4:  # Ensure minimum required cells
                        self._extract_horse_data(cells, race_info, data)
            
            df = pd.DataFrame(data)
            
            if not df.empty:
                df = self._process_historical_data(df)
            
            return df
            
        except requests.RequestException as e:
            logger.error(f"Error fetching race {race_no}: {e}")
            return None
        except Exception as e:
            logger.error(f"Unexpected error processing race {race_no}: {e}")
            return None
    
    def _create_terry_sheet(self, df: pd.DataFrame) -> pd.DataFrame:
        """Create Terry AI formatted sheet."""
        # First, ensure all required columns exist in the DataFrame
        required_columns = [
            '場次', '班次', '路程', '馬匹編號', '馬名', '騎師', '練馬師', '評分',
            '上次日期', '上賽距今日數', '上次場次', '上次班次', '上次場地狀況',
            '負磅', '上次負磅', '上次負磅 +/-', '檔位', '上次檔位', '上次檔位 +/-',
            '最佳時間', '上次賽事時間', '上次完成時間',
            '上次調整後完成時間', '上次調整後完成秒速', '上次頭段完成時間', '上次頭段',
            '上次彎位 400', '上次最後 400', '上次最後 800', '上次調整後最後 800',
            '網址'
        ]
        
        # Add missing columns that will be calculated
        calculated_columns = ['上次完成時間 - 上次賽事時間', '賽事結果連結', '賽事重溫連結']
        
        # Create a copy of the DataFrame with only existing columns
        existing_columns = [col for col in required_columns if col in df.columns]
        df_terry = df[existing_columns].copy()
        
        # Add calculated columns
        for col in calculated_columns:
            df_terry[col] = None
        
        race_date_ts = pd.Timestamp(self.race_date)
        
        for index, row in df_terry.iterrows():
            if pd.notna(row.get('上次日期')):
                last_date = pd.to_datetime(row['上次日期'])
                df_terry.at[index, '上賽距今日數'] = (race_date_ts - last_date).days
                
                # Calculate time difference
                if pd.notna(row.get('上次完成時間')) and pd.notna(row.get('上次賽事時間')):
                    try:
                        finish_time_str = str(row['上次完成時間'])
                        race_time_str = str(row['上次賽事時間'])
                        
                        # Parse finish time
                        if ':' in finish_time_str:
                            time_parts = finish_time_str.split(':')
                            finish_time = int(time_parts[0]) * 60 + float(time_parts[1])
                        else:
                            finish_time = float(finish_time_str)
                        
                        # Parse race time based on distance
                        if row.get('路程', 0) >= 1200 and ':' in race_time_str:
                            time_parts = race_time_str.split(':')
                            race_time = int(time_parts[0]) * 60 + float(time_parts[1])
                        else:
                            race_time = float(race_time_str) if race_time_str != 'nan' else 0
                        
                        df_terry.at[index, '上次完成時間 - 上次賽事時間'] = finish_time - race_time
                    except (ValueError, IndexError, TypeError):
                        df_terry.at[index, '上次完成時間 - 上次賽事時間'] = None
                
                # Add result and replay links
                if pd.notna(row.get('上次場次')):
                    try:
                        race_no = int(row['上次場次'])
                        df_terry.at[index, '賽事結果連結'] = f"https://racing.hkjc.com/racing/information/chinese/Racing/LocalResults.aspx?RaceDate={last_date.strftime('%Y/%m/%d')}&RaceNo={race_no}"
                        df_terry.at[index, '賽事重溫連結'] = f"https://racing.hkjc.com/racing/video/play.asp?type=replay-full&date={last_date.strftime('%Y%m%d')}&no={race_no:02d}&lang=chi"
                    except (ValueError, TypeError):
                        pass
        
        return df_terry
    
    def _create_alfred_sheet(self, df: pd.DataFrame) -> pd.DataFrame:
        """Create Alfred AI formatted sheet."""
        required_columns = [
            '場次', '班次', '路程', '馬匹編號', '馬名', '騎師', '評分',
            '上次日期', '上賽距今日數', '上次班次', '上次場地狀況',
            '負磅', '上次負磅', '上次負磅 +/-', '檔位', '上次檔位', '上次檔位 +/-',
            '最佳時間', '上次完成時間', '上次頭段完成時間', '上次彎位 400',
            '上次最後 400', '上次最後 800', '網址', '上次場次'
        ]
        
        # Add missing columns that will be calculated
        calculated_columns = ['賽事結果連結', '賽事重溫連結']
        
        # Create a copy of the DataFrame with only existing columns
        existing_columns = [col for col in required_columns if col in df.columns]
        df_alfred = df[existing_columns].copy()
        
        # Add calculated columns
        for col in calculated_columns:
            df_alfred[col] = None
        
        race_date_ts = pd.Timestamp(self.race_date)
        
        for index, row in df_alfred.iterrows():
            if pd.notna(row.get('上次日期')):
                last_date = pd.to_datetime(row['上次日期'])
                df_alfred.at[index, '上賽距今日數'] = (race_date_ts - last_date).days
                
                # Add result and replay links
                if pd.notna(row.get('上次場次')):
                    try:
                        race_no = int(row['上次場次'])
                        df_alfred.at[index, '賽事結果連結'] = f"https://racing.hkjc.com/racing/information/chinese/Racing/LocalResults.aspx?RaceDate={last_date.strftime('%Y/%m/%d')}&RaceNo={race_no}"
                        df_alfred.at[index, '賽事重溫連結'] = f"https://racing.hkjc.com/racing/video/play.asp?type=replay-full&date={last_date.strftime('%Y%m%d')}&no={race_no:02d}&lang=chi"
                    except (ValueError, TypeError):
                        pass
        
        return df_alfred
    
    def _create_combined_files(self, date_str: str) -> None:
        """Create combined Excel files for different formats."""
        # Standard combined file
        standard_sheets_created = 0
        with pd.ExcelWriter(f"Racecard_{date_str}.xlsx", engine='openpyxl') as writer:
            for race_no in range(1, 12):
                try:
                    df = pd.read_excel(f"Racecard_{date_str}_{race_no}.xlsx")
                    if not df.empty:
                        df.to_excel(writer, sheet_name=f"Race_{race_no}", index=False)
                        standard_sheets_created += 1
                except FileNotFoundError:
                    logger.warning(f"File Racecard_{date_str}_{race_no}.xlsx not found")
                except Exception as e:
                    logger.error(f"Error adding Race {race_no} to standard file: {e}")
        
        if standard_sheets_created == 0:
            logger.warning(f"No sheets created for standard file Racecard_{date_str}.xlsx")
        
        # Terry format
        terry_sheets_created = 0
        with pd.ExcelWriter(f"Racecard_{date_str}_Terry.xlsx", engine='openpyxl') as writer:
            for race_no in range(1, 12):
                try:
                    df = pd.read_excel(f"Racecard_{date_str}_{race_no}.xlsx")
                    if not df.empty:
                        df_terry = self._create_terry_sheet(df)
                        if not df_terry.empty:
                            df_terry.to_excel(writer, sheet_name=f"Race_{race_no}", index=False)
                            terry_sheets_created += 1
                except FileNotFoundError:
                    logger.warning(f"File Racecard_{date_str}_{race_no}.xlsx not found for Terry format")
                except Exception as e:
                    logger.error(f"Error creating Terry sheet for Race {race_no}: {e}")
            
            # Create a dummy sheet if no sheets were created
            if terry_sheets_created == 0:
                dummy_df = pd.DataFrame({'Message': ['No race data available']})
                dummy_df.to_excel(writer, sheet_name="No_Data", index=False)
                logger.warning("Created dummy sheet for Terry format due to no valid data")
        
        # Alfred format
        alfred_sheets_created = 0
        with pd.ExcelWriter(f"Racecard_{date_str}_Alfred.xlsx", engine='openpyxl') as writer:
            for race_no in range(1, 12):
                try:
                    df = pd.read_excel(f"Racecard_{date_str}_{race_no}.xlsx")
                    if not df.empty:
                        df_alfred = self._create_alfred_sheet(df)
                        if not df_alfred.empty:
                            df_alfred.to_excel(writer, sheet_name=f"Alfred_{race_no}", index=False)
                            alfred_sheets_created += 1
                except FileNotFoundError:
                    logger.warning(f"File Racecard_{date_str}_{race_no}.xlsx not found for Alfred format")
                except Exception as e:
                    logger.error(f"Error creating Alfred sheet for Race {race_no}: {e}")
            
            # Create a dummy sheet if no sheets were created
            if alfred_sheets_created == 0:
                dummy_df = pd.DataFrame({'Message': ['No race data available']})
                dummy_df.to_excel(writer, sheet_name="No_Data", index=False)
                logger.warning("Created dummy sheet for Alfred format due to no valid data")
        
        logger.info(f"Created combined files: Standard({standard_sheets_created} sheets), Terry({terry_sheets_created} sheets), Alfred({alfred_sheets_created} sheets)")
    
    def download_all_racecards(self) -> None:
        """Download all racecards for the configured race date and course."""
        date_str = self.race_date.strftime('%Y%m%d')
        logger.info(f"Starting racecard download for {self.race_date} at {self.race_course}")
        
        successful_downloads = 0
        individual_files = []  # Track individual files created
        
        # Download individual race cards (typically 1-11 races)
        for race_no in range(1, 12):
            logger.info(f"Processing {self.race_date} {self.race_course} Race {race_no}")
            
            df = self._scrape_race_data(race_no)
            
            if df is not None and not df.empty:
                # Save individual race file
                filename = f"Racecard_{date_str}_{race_no}.xlsx"
                try:
                    df.to_excel(filename, index=False)
                    logger.info(f"Saved {filename}")
                    successful_downloads += 1
                    individual_files.append(filename)  # Track the file
                except Exception as e:
                    logger.error(f"Error saving {filename}: {e}")
            else:
                logger.warning(f"No data found for Race {race_no}")
        
        # Create combined files in different formats
        if successful_downloads > 0:
            try:
                self._create_combined_files(date_str)
                
                # Remove individual race files after successful combined file creation
                self._cleanup_individual_files(individual_files)
                
            except Exception as e:
                logger.error(f"Error creating combined files: {e}")
        else:
            logger.warning("No successful downloads, skipping combined file creation")
        
        logger.info(f"Racecard download completed. Successfully processed {successful_downloads} races.")
    
    def _cleanup_individual_files(self, individual_files: List[str]) -> None:
        """Remove individual race files after combined files are created."""
        removed_count = 0
        
        for filename in individual_files:
            try:
                file_path = Path(filename)
                if file_path.exists():
                    file_path.unlink()  # Delete the file
                    logger.info(f"Removed individual file: {filename}")
                    removed_count += 1
                else:
                    logger.warning(f"File not found for cleanup: {filename}")
            except Exception as e:
                logger.error(f"Error removing file {filename}: {e}")
        
        logger.info(f"Cleanup completed. Removed {removed_count} individual race files.")
    
def main():
    # Configuration
    race_date = date(2025, 11, 12)
    race_course = "HV"  # ST / HV
    
    # Create downloader and process
    downloader = HKJCRacecardDownloader(race_date, race_course)
    downloader.download_all_racecards()
    
    logger.info("Racecard download completed!")

if __name__ == "__main__":
    main()