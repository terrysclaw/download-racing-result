"""
Hong Kong Horse Racing Analysis Module
計算馬匹勝率、騎師因子、練馬師影響等投注分析指標
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Optional, Tuple
import logging

logger = logging.getLogger(__name__)


class RacingAnalyzer:
    """香港賽馬分析器 - 計算勝率、騎師因子、一致性評分"""
    
    def __init__(self, historical_df: Optional[pd.DataFrame] = None):
        """
        初始化分析器
        
        Args:
            historical_df: 歷史比賽資料DataFrame
        """
        if historical_df is None:
            self.historical_df = pd.DataFrame()
        else:
            self.historical_df = historical_df
        self.jockey_stats = None
        self.trainer_stats = None
        self.combo_stats = None
        
        # 初始化統計數據
        if not self.historical_df.empty:
            self.build_jockey_stats()
            self.build_trainer_stats()
            self.build_trainer_jockey_combo_stats()
    
    # ==================== 勝率計算模型 ====================
    
    def calculate_win_probability(self, rating: int, race_class: str, field_size: int = 10) -> float:
        """
        根據評分計算馬匹的理論勝率 (Kelly Criterion)
        
        Args:
            rating: 馬匹評分 (例如: 72)
            race_class: 比賽班次 (例如: "第 1 班")
            field_size: 比賽馬匹數 (預設10匹)
            
        Returns:
            勝率 (0.0 - 1.0)
        
        Examples:
            >>> analyzer.calculate_win_probability(80, "第 1 班", 10)
            0.35  # 約35%勝率
        """
        try:
            # 評分轉換為勝率 (基於香港賽馬歷史統計)
            # 評分越高，勝率越高
            
            # 標準化：假設70分為場地平均
            rating_diff = rating - 70
            
            # 基礎勝率：場地平均
            base_probability = 1.0 / max(field_size, 3)
            
            # 評分優勢
            if rating_diff <= -15:
                # 評分較弱
                return base_probability * 0.4
            elif rating_diff < -5:
                # 評分略弱
                return base_probability * 0.7
            elif rating_diff <= 5:
                # 評分接近
                return base_probability * 1.0
            elif rating_diff < 15:
                # 評分略優
                return base_probability * 1.4
            elif rating_diff < 25:
                # 評分優
                return base_probability * 1.8
            else:
                # 評分遠優
                return base_probability * 2.2
            
        except Exception as e:
            logger.warning(f"Error calculating win probability for rating {rating}: {e}")
            return 1.0 / max(field_size, 3)
    
    def calculate_rating_advantage(self, horse_rating: int, field_ratings: list) -> float:
        """
        計算馬匹相對於同場對手的評分優勢
        
        Args:
            horse_rating: 該馬評分
            field_ratings: 同場所有馬評分列表
            
        Returns:
            評分優勢幅度 (正數表示優勢，負數表示劣勢)
        """
        try:
            if not field_ratings:
                return 0.0
            
            field_ratings = [r for r in field_ratings if pd.notna(r) and r != 0]
            if not field_ratings:
                return 0.0
            
            avg_rating = np.mean(field_ratings)
            return horse_rating - avg_rating
            
        except Exception as e:
            logger.warning(f"Error calculating rating advantage: {e}")
            return 0.0
    
    # ==================== 騎師因子 ====================
    
    def build_jockey_stats(self) -> pd.DataFrame:
        """
        從歷史資料構建騎師統計表
        
        Returns:
            騎師統計DataFrame (騎師名稱, 勝場, 總場次, 勝率, 平均賠率等)
        """
        try:
            if self.historical_df.empty:
                logger.warning("No historical data available for jockey stats")
                return pd.DataFrame()
            
            # 假設欄位名稱：騎師, 名次, 獨贏賠率, 馬場, 路程, 泥草
            df = self.historical_df.copy()
            
            # 識別名次欄位 (可能是 '名次', 'finish_position' 等)
            position_col = None
            for col in ['名次', '排名', 'finish_position', 'position']:
                if col in df.columns:
                    position_col = col
                    break
            
            if position_col is None:
                logger.warning("Cannot find position column in historical data")
                return pd.DataFrame()
            
            jockey_col = None
            for col in ['騎師', 'jockey', '騎手']:
                if col in df.columns:
                    jockey_col = col
                    break
            
            if jockey_col is None:
                logger.warning("Cannot find jockey column in historical data")
                return pd.DataFrame()
            
            # 清理資料：移除 NaN 騎師
            df = df[df[jockey_col].notna()].copy()
            df = df[df[jockey_col] != '']
            
            # 計算名次 (轉換為數字，1 = 勝利)
            df['is_win'] = df[position_col].apply(
                lambda x: 1 if (pd.notna(x) and (x == 1 or str(x).strip() == '1')) else 0
            )
            
            # 按騎師分組統計
            jockey_stats = df.groupby(jockey_col).agg({
                'is_win': ['sum', 'count'],
            }).reset_index()
            
            jockey_stats.columns = ['騎師', '勝場', '總場次']
            jockey_stats['勝率'] = jockey_stats['勝場'] / jockey_stats['總場次']
            
            # 計算平均賠率
            odds_col = None
            for col in ['獨贏賠率', '賠率', '獨贏', 'odds']:
                if col in df.columns:
                    odds_col = col
                    break
            
            if odds_col:
                avg_odds = df.groupby(jockey_col)[odds_col].mean()
                jockey_stats['平均賠率'] = jockey_stats['騎師'].map(avg_odds)
            
            # 計算場地和距離的專項統計
            if '馬場' in df.columns:
                venue_stats = df.groupby([jockey_col, '馬場'])['is_win'].agg(['sum', 'count']).reset_index()
                venue_stats.columns = ['騎師', '馬場', '勝場', '總場次']
                venue_stats['勝率'] = venue_stats['勝場'] / venue_stats['總場次']
                jockey_stats['馬場_stats'] = jockey_stats['騎師'].apply(
                    lambda j: venue_stats[venue_stats['騎師'] == j].to_dict('records')
                )
            
            jockey_stats = jockey_stats.sort_values('勝場', ascending=False)
            self.jockey_stats = jockey_stats
            
            logger.info(f"Built jockey stats for {len(jockey_stats)} jockeys")
            return jockey_stats
            
        except Exception as e:
            logger.error(f"Error building jockey stats: {e}")
            return pd.DataFrame()
    
    def get_jockey_factor(self, jockey_name: str, venue: Optional[str] = None, distance: Optional[int] = None) -> float:
        """
        獲取騎師因子 (相對於平均騎師的倍數)
        
        Args:
            jockey_name: 騎師名稱
            venue: 馬場 (HV/ST) - 可選，用於特定馬場統計
            distance: 距離 - 可選，用於特定距離統計
            
        Returns:
            騎師因子 (1.0 = 平均, >1.0 = 優於平均, <1.0 = 弱於平均)
        
        Examples:
            >>> analyzer.get_jockey_factor("蔡明紹")
            1.15  # 比平均騎師強15%
        """
        try:
            if self.jockey_stats is None or self.jockey_stats.empty:
                logger.warning("Jockey stats not available")
                return 1.0
            
            # 尋找騎師
            jockey_data = self.jockey_stats[self.jockey_stats['騎師'] == jockey_name]
            
            if jockey_data.empty:
                # 騎師不在記錄中，返回平均
                logger.debug(f"Jockey {jockey_name} not found in stats")
                return 1.0
            
            # 計算勝率相對於整體平均的倍數
            overall_avg_wr = self.jockey_stats['勝率'].mean()
            jockey_wr = jockey_data['勝率'].iloc[0]
            
            if overall_avg_wr > 0:
                factor = jockey_wr / overall_avg_wr
                return max(0.5, min(2.0, factor))  # 限制在 0.5 - 2.0 之間
            
            return 1.0
            
        except Exception as e:
            logger.warning(f"Error getting jockey factor for {jockey_name}: {e}")
            return 1.0
    
    def build_trainer_stats(self) -> pd.DataFrame:
        """
        從歷史資料構建練馬師統計表
        
        Returns:
            練馬師統計DataFrame
        """
        try:
            if self.historical_df.empty:
                logger.warning("No historical data available for trainer stats")
                return pd.DataFrame()
            
            df = self.historical_df.copy()
            
            # 識別練馬師欄位
            trainer_col = None
            for col in ['練馬師', 'trainer', '教練']:
                if col in df.columns:
                    trainer_col = col
                    break
            
            if trainer_col is None:
                logger.warning("Cannot find trainer column in historical data")
                return pd.DataFrame()
            
            # 識別名次欄位
            position_col = None
            for col in ['名次', '排名', 'finish_position', 'position']:
                if col in df.columns:
                    position_col = col
                    break
            
            if position_col is None:
                return pd.DataFrame()
            
            # 清理資料
            df = df[df[trainer_col].notna()].copy()
            df = df[df[trainer_col] != '']
            
            # 計算勝利
            df['is_win'] = df[position_col].apply(
                lambda x: 1 if (pd.notna(x) and (x == 1 or str(x).strip() == '1')) else 0
            )
            
            # 按練馬師分組統計
            trainer_stats = df.groupby(trainer_col).agg({
                'is_win': ['sum', 'count'],
            }).reset_index()
            
            trainer_stats.columns = ['練馬師', '勝場', '總場次']
            trainer_stats['勝率'] = trainer_stats['勝場'] / trainer_stats['總場次']
            trainer_stats = trainer_stats.sort_values('勝場', ascending=False)
            
            self.trainer_stats = trainer_stats
            logger.info(f"Built trainer stats for {len(trainer_stats)} trainers")
            return trainer_stats
            
        except Exception as e:
            logger.error(f"Error building trainer stats: {e}")
            return pd.DataFrame()

    def build_trainer_jockey_combo_stats(self) -> pd.DataFrame:
        """Build trainer-jockey combo statistics from historical data."""
        try:
            if self.historical_df.empty:
                logger.warning("No historical data available for trainer-jockey combo stats")
                return pd.DataFrame()

            df = self.historical_df.copy()

            trainer_col = next((c for c in ['練馬師', 'trainer', '教練'] if c in df.columns), None)
            jockey_col = next((c for c in ['騎師', 'jockey', '騎手'] if c in df.columns), None)
            position_col = next((c for c in ['名次', '排名', 'finish_position', 'position'] if c in df.columns), None)

            if not trainer_col or not jockey_col or not position_col:
                logger.warning("Cannot build combo stats due to missing trainer/jockey/position columns")
                return pd.DataFrame()

            df = df[df[trainer_col].notna() & (df[trainer_col] != '')]
            df = df[df[jockey_col].notna() & (df[jockey_col] != '')]
            if df.empty:
                return pd.DataFrame()

            df['is_win'] = df[position_col].apply(
                lambda x: 1 if (pd.notna(x) and (x == 1 or str(x).strip() == '1')) else 0
            )

            combo_stats = df.groupby([trainer_col, jockey_col]).agg({'is_win': ['sum', 'count']}).reset_index()
            combo_stats.columns = ['練馬師', '騎師', '勝場', '總場次']
            combo_stats['勝率'] = combo_stats['勝場'] / combo_stats['總場次']

            self.combo_stats = combo_stats.sort_values(['勝場', '總場次'], ascending=False)
            logger.info(f"Built trainer-jockey combo stats for {len(self.combo_stats)} combos")
            return self.combo_stats

        except Exception as e:
            logger.error(f"Error building trainer-jockey combo stats: {e}")
            return pd.DataFrame()

    def get_trainer_jockey_combo_factor(self, trainer_name: str, jockey_name: str) -> float:
        """Get trainer-jockey combo factor relative to historical average win rate."""
        try:
            if not trainer_name or not jockey_name:
                return 1.0

            if self.combo_stats is None or self.combo_stats.empty:
                return 1.0

            combo_data = self.combo_stats[
                (self.combo_stats['練馬師'] == trainer_name) &
                (self.combo_stats['騎師'] == jockey_name)
            ]

            if combo_data.empty:
                return 1.0

            sample_size = int(combo_data['總場次'].iloc[0])
            combo_wr = float(combo_data['勝率'].iloc[0])
            overall_wr = float(self.combo_stats['勝率'].mean())

            if overall_wr <= 0:
                return 1.0

            # Blend toward neutral when sample size is small.
            reliability = min(1.0, sample_size / 20.0)
            raw_factor = combo_wr / overall_wr
            blended_factor = 1.0 + (raw_factor - 1.0) * reliability
            return max(0.7, min(1.5, blended_factor))

        except Exception as e:
            logger.warning(f"Error getting trainer-jockey combo factor for {trainer_name}/{jockey_name}: {e}")
            return 1.0

    def classify_pace_style(self, head_section_times: List[float]) -> str:
        """Classify pace style by average head section time of recent runs."""
        try:
            times = [float(t) for t in head_section_times if pd.notna(t)]
            if not times:
                return 'Unknown'

            avg_time = float(np.mean(times))

            if avg_time <= 24.0:
                return 'Front-runner'
            if avg_time <= 25.2:
                return 'Prominent'
            return 'Closer'

        except Exception:
            return 'Unknown'

    def _parse_time_to_seconds(self, raw_value) -> Optional[float]:
        """Parse race time values that may be numeric or MM:SS.xx strings."""
        if pd.isna(raw_value):
            return None

        text = str(raw_value).strip()
        if text in ['', '---']:
            return None

        try:
            if ':' in text:
                parts = text.split(':')
                if len(parts) == 2:
                    minutes = int(parts[0])
                    seconds = float(parts[1])
                    return minutes * 60 + seconds
            return float(text)
        except (ValueError, TypeError):
            return None

    def _get_first_valid_column_value(self, result_row: pd.Series, candidates: List[str]):
        """Return first existing column value from candidates."""
        for col in candidates:
            if col in result_row.index:
                return result_row.get(col)
        return None

    def _extract_head_section_time_from_row(self, result_row: pd.Series) -> Optional[float]:
        """Extract head section cumulative time from one historical result row."""
        try:
            available_sections = []
            for i in range(1, 7):
                val = result_row.get(f'第 {i} 段')
                if self._parse_time_to_seconds(val) is not None:
                    available_sections.append(i)

            if not available_sections:
                return None

            last_section = max(available_sections)
            if last_section >= 6:
                raw = self._get_first_valid_column_value(result_row, ['賽事時間5', '賽事時間 5'])
            elif last_section >= 5:
                raw = self._get_first_valid_column_value(result_row, ['賽事時間4', '賽事時間 4'])
            elif last_section >= 4:
                raw = self._get_first_valid_column_value(result_row, ['賽事時間3', '賽事時間 3'])
            elif last_section >= 3:
                raw = self._get_first_valid_column_value(result_row, ['賽事時間2', '賽事時間 2'])
            else:
                raw = self._get_first_valid_column_value(result_row, ['賽事時間1', '賽事時間 1'])

            return self._parse_time_to_seconds(raw)

        except Exception:
            return None

    def get_recent_head_section_times(
        self,
        cloth_no: str,
        venue: Optional[str] = None,
        surface: Optional[str] = None,
        distance: Optional[int] = None,
        limit: int = 3
    ) -> List[float]:
        """Get recent head section times for pace mapping."""
        try:
            if self.historical_df.empty or not cloth_no or '布號' not in self.historical_df.columns:
                return []

            df = self.historical_df.copy()
            df = df[df['布號'].astype(str) == str(cloth_no)]

            if venue and '馬場' in df.columns:
                df = df[df['馬場'] == venue]
            if surface and '泥草' in df.columns:
                df = df[df['泥草'] == surface]
            if distance and '路程' in df.columns:
                df = df[df['路程'] == distance]

            if df.empty:
                return []

            if '日期' in df.columns:
                df = df.sort_values('日期')

            recent = df.tail(max(limit, 1))
            times: List[float] = []
            for _, row in recent.iterrows():
                t = self._extract_head_section_time_from_row(row)
                if t is not None:
                    times.append(t)

            return times[-limit:]

        except Exception as e:
            logger.debug(f"Error getting recent head section times for {cloth_no}: {e}")
            return []
    
    # ==================== 一致性評分 ====================
    
    def calculate_consistency_score(self, form_string: str) -> Tuple[float, Dict]:
        """
        計算馬匹的一致性評分 (基於6次近績)
        
        Args:
            form_string: 近績字符串 (例如: "311242" 表示第1、1、2、4、2名)
            
        Returns:
            (一致性評分 0-100, 詳細統計字典)
        
        Examples:
            >>> analyzer.calculate_consistency_score("311242")
            (78.5, {'連貫': 2, '最差排名': 4, ...})
        """
        try:
            if not form_string or pd.isna(form_string):
                return 0.0, {}
            
            form_str = str(form_string).strip()
            
            # 解析每個位置
            positions = []
            for char in form_str:
                if char.isdigit() and char != '0':
                    try:
                        positions.append(int(char))
                    except ValueError:
                        pass
            
            if not positions:
                return 0.0, {}
            
            # 計算統計數據
            stats = {
                '總場次': len(positions),
                '最佳排名': min(positions),
                '最差排名': max(positions),
                '平均排名': np.mean(positions),
                '入四強': sum(1 for p in positions if p <= 4),
                '連貫得分': 0,  # 連續強表現
                '穩定性': 0,  # 標準差反向 (越小越穩定)
            }
            
            # 連貫得分：連續入四強的場數
            consecutive = 0
            max_consecutive = 0
            for pos in positions:
                if pos <= 4:
                    consecutive += 1
                    max_consecutive = max(max_consecutive, consecutive)
                else:
                    consecutive = 0
            stats['連貫得分'] = max_consecutive
            
            # 穩定性：標準差越小越穩定
            if len(positions) > 1:
                std_dev = np.std(positions)
                stats['穩定性'] = max(0, 100 - std_dev * 10)  # 標準差每增加0.1減1分
            else:
                stats['穩定性'] = 50
            
            # 計算綜合一致性評分 (0-100)
            score = 0
            
            # 最佳排名權重 (20%)
            if stats['最佳排名'] == 1:
                score += 20
            elif stats['最佳排名'] <= 3:
                score += 15
            elif stats['最佳排名'] <= 6:
                score += 10
            
            # 入四強比例權重 (30%)
            entry_rate = stats['入四強'] / stats['總場次']
            score += entry_rate * 30
            
            # 連貫得分權重 (30%)
            consistency_bonus = min(stats['連貫得分'] / 3 * 30, 30)
            score += consistency_bonus
            
            # 穩定性權重 (20%)
            score += stats['穩定性'] * 0.2
            
            consistency_score = max(0, min(100, score))
            
            return consistency_score, stats
            
        except Exception as e:
            logger.warning(f"Error calculating consistency score for {form_string}: {e}")
            return 0.0, {}
    
    # ==================== 綜合分析 ====================
    
    def calculate_composite_score(self, 
                                 horse_rating: int,
                                 jockey_name: str,
                                 form_string: str,
                                 race_class: str,
                                 field_size: int = 10) -> Dict:
        """
        計算馬匹的綜合投注評分
        
        Args:
            horse_rating: 馬匹評分
            jockey_name: 騎師名稱
            form_string: 近績 (如 "311242")
            race_class: 比賽班次
            field_size: 比賽馬匹數
            
        Returns:
            包含各項得分的字典
        
        Examples:
            >>> analyzer.calculate_composite_score(75, "蔡明紹", "311242", "第 1 班", 10)
            {
                '評分得分': 35,
                '騎師得分': 38,
                '一致性得分': 62,
                '綜合得分': 45,
                '評估': 'STRONG'
            }
        """
        try:
            result = {}
            
            # 1. 評分得分 (0-50)
            win_prob = self.calculate_win_probability(horse_rating, race_class, field_size)
            rating_score = min(50, win_prob * 100)
            result['評分得分'] = rating_score
            
            # 2. 騎師得分 (0-30)
            jockey_factor = self.get_jockey_factor(jockey_name)
            jockey_score = min(30, (jockey_factor - 0.5) * 30)  # 0.5-2.0 對應 0-45
            result['騎師得分'] = max(0, jockey_score)
            
            # 3. 一致性得分 (0-20)
            consistency_score, _ = self.calculate_consistency_score(form_string)
            result['一致性得分'] = consistency_score / 5  # 標準化為 0-20
            
            # 綜合評分 (加權)
            composite = (rating_score * 0.50 + 
                        result['騎師得分'] * 0.30 + 
                        result['一致性得分'] * 0.20)
            result['綜合評分'] = composite
            
            # 評估等級
            if composite >= 70:
                result['評估'] = '★★★ STRONG'
            elif composite >= 55:
                result['評估'] = '★★ MODERATE'
            elif composite >= 40:
                result['評估'] = '★ WEAK'
            else:
                result['評估'] = '✗ POOR'
            
            return result
            
        except Exception as e:
            logger.error(f"Error calculating composite score: {e}")
            return {}


# ==================== 導出函式 ====================

def export_jockey_stats_to_csv(historical_df: pd.DataFrame, filename: str = 'jockey_stats.csv') -> None:
    """
    將騎師統計導出為 CSV 檔案
    
    Args:
        historical_df: 歷史比賽資料
        filename: 輸出檔案名稱
    """
    try:
        analyzer = RacingAnalyzer(historical_df)
        if analyzer.jockey_stats is not None and not analyzer.jockey_stats.empty:
            analyzer.jockey_stats.to_csv(filename, index=False, encoding='utf-8-sig')
            logger.info(f"Exported jockey stats to {filename}")
        else:
            logger.warning("No jockey stats to export")
    except Exception as e:
        logger.error(f"Error exporting jockey stats: {e}")


def export_trainer_stats_to_csv(historical_df: pd.DataFrame, filename: str = 'trainer_stats.csv') -> None:
    """
    將練馬師統計導出為 CSV 檔案
    
    Args:
        historical_df: 歷史比賽資料
        filename: 輸出檔案名稱
    """
    try:
        analyzer = RacingAnalyzer(historical_df)
        if analyzer.trainer_stats is not None and not analyzer.trainer_stats.empty:
            analyzer.trainer_stats.to_csv(filename, index=False, encoding='utf-8-sig')
            logger.info(f"Exported trainer stats to {filename}")
        else:
            logger.warning("No trainer stats to export")
    except Exception as e:
        logger.error(f"Error exporting trainer stats: {e}")
