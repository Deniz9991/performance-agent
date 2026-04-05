import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from typing import Dict, List, Tuple, Any
from sklearn.linear_model import LinearRegression
import warnings
warnings.filterwarnings('ignore')

class PerformanceMatrixAgent:
    """
    ИИ-агент для анализа эффективности сотрудников
    Использует линейную регрессию для прогнозирования
    """
    
    def __init__(self):
        self.targets = {
            'dispatcher': {
                'transfer_rate': 70.0,
                'max_failed_calls': 15.0,
                'max_rejected': 10.0,
                'min_leads_per_day': 20,
                'max_duplicate_rate': 5.0,
                'max_unassigned_rate': 2.0,
            }
        }
        
        self.statuses = {
            'new': ['новый лид', 'новый', 'new', 'новое'],
            'transferred': ['передан менеджеру', 'передан', 'передано', 'передана'],
            'failed_call': ['недозвон', 'нет ответа', 'не берут трубку', 'no answer'],
            'recall': ['перезвон', 'перезвонить', 'call back'],
            'test': ['тест', 'test'],
            'duplicate': ['дубль', 'дубликат', 'duplicate', 'повтор', 'повторно'],
            'wrong_number': ['ошибка номера', 'неверный номер', 'wrong number'],
            'refused': ['отказ', 'отказался', 'не заинтересован', 'refused'],
            'in_work': []
        }
        
        self.poor_quality_statuses = ['failed_call', 'duplicate', 'wrong_number', 'refused', 'recall']
    
    def detect_columns(self, df: pd.DataFrame) -> Dict[str, str]:
        """Определяет нужные колонки в DataFrame"""
        columns_map = {
            'маркетолог': None,
            'ответственный': None,
            'этап': None,
            'дата_создания': None,
            'ленд': None,
            'клиент': None,
            'телефон': None,
            'название': None
        }
        
        for col in df.columns:
            col_lower = col.lower()
            if 'маркетолог' in col_lower or 'маркет' in col_lower:
                columns_map['маркетолог'] = col
            elif 'ответственный' in col_lower or 'ответств' in col_lower:
                columns_map['ответственный'] = col
            elif 'этап' in col_lower or 'status' in col_lower:
                columns_map['этап'] = col
            elif 'дата' in col_lower and ('создан' in col_lower or 'create' in col_lower):
                columns_map['дата_создания'] = col
            elif 'ленд' in col_lower or 'land' in col_lower or 'источник' in col_lower:
                columns_map['ленд'] = col
            elif 'клиент' in col_lower or 'client' in col_lower:
                columns_map['клиент'] = col
            elif 'телефон' in col_lower or 'phone' in col_lower:
                columns_map['телефон'] = col
            elif 'название' in col_lower or 'name' in col_lower:
                columns_map['название'] = col
        
        return columns_map
    
    def is_aggregated_data(self, df: pd.DataFrame) -> bool:
        """Определяет, являются ли данные агрегированными"""
        if len(df) < 100:
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            for col in numeric_cols:
                if any(word in col.lower() for word in ['%', 'средн', 'конверси', 'процент']):
                    return True
            name_cols = [col for col in df.columns if any(word in col.lower() for word in ['имя', 'фио', 'сотрудник', 'маркетолог'])]
            if name_cols and len(numeric_cols) >= 2:
                return True
        return False
    
    def analyze_aggregated_table(self, df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, Dict]:
        """Анализирует агрегированную таблицу"""
        name_col = None
        for col in df.columns:
            if any(word in col.lower() for word in ['имя', 'фио', 'сотрудник', 'маркетолог', 'диспетчер']):
                name_col = col
                break
        
        if not name_col:
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), {'error': 'Не найдена колонка с именами'}
        
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        
        leads_col = None
        transferred_col = None
        failed_col = None
        rejected_col = None
        
        for col in numeric_cols:
            col_lower = col.lower()
            if 'лид' in col_lower or 'получен' in col_lower or 'всего' in col_lower:
                leads_col = col
            elif 'передан' in col_lower:
                transferred_col = col
            elif 'недозвон' in col_lower:
                failed_col = col
            elif 'отказ' in col_lower or 'брак' in col_lower:
                rejected_col = col
        
        if not leads_col and len(numeric_cols) >= 1:
            leads_col = numeric_cols[0]
        if not transferred_col and len(numeric_cols) >= 2:
            transferred_col = numeric_cols[1]
        if not failed_col and len(numeric_cols) >= 3:
            failed_col = numeric_cols[2]
        if not rejected_col and len(numeric_cols) >= 4:
            rejected_col = numeric_cols[3]
        
        results = []
        
        for _, row in df.iterrows():
            name = str(row[name_col]).strip()
            if not name or name == 'nan':
                continue
            
            total_leads = float(row[leads_col]) if leads_col and pd.notna(row[leads_col]) else 0
            transferred = float(row[transferred_col]) if transferred_col and pd.notna(row[transferred_col]) else 0
            failed = float(row[failed_col]) if failed_col and pd.notna(row[failed_col]) else 0
            rejected = float(row[rejected_col]) if rejected_col and pd.notna(row[rejected_col]) else 0
            
            transfer_rate = (transferred / total_leads * 100) if total_leads > 0 else 0
            failed_rate = (failed / total_leads * 100) if total_leads > 0 else 0
            rejected_rate = (rejected / total_leads * 100) if total_leads > 0 else 0
            
            if transfer_rate < self.targets['dispatcher']['transfer_rate']:
                predicted_transfer = transfer_rate + (self.targets['dispatcher']['transfer_rate'] - transfer_rate) * 0.3
            else:
                predicted_transfer = transfer_rate * 0.97
            
            is_problem = (
                transfer_rate < self.targets['dispatcher']['transfer_rate'] or
                failed_rate > self.targets['dispatcher']['max_failed_calls'] or
                rejected_rate > self.targets['dispatcher']['max_rejected']
            )
            
            results.append({
                'Сотрудник': name,
                'Получено лидов': int(total_leads),
                'Передано продажам': int(transferred),
                'Недозвоны': int(failed),
                'Отказы': int(rejected),
                'Дубли/Повторы': 0,
                'Ошибка номера': 0,
                '% переданных': round(transfer_rate, 1),
                '% недозвонов': round(failed_rate, 1),
                '% отказов': round(rejected_rate, 1),
                '% дублей/повторов': 0,
                'Лидов в день': 0,
                'Загруженность': 'Нет данных',
                '📈 ПРОГНОЗ: % переданных': round(predicted_transfer, 1),
                '📈 ПРОГНОЗ: передач (шт)': round(total_leads * predicted_transfer / 100, 0),
                'is_problem': is_problem
            })
        
        summary_df = pd.DataFrame(results)
        return summary_df, pd.DataFrame(), pd.DataFrame(), {'type': 'aggregated'}
    
    def get_dispatcher_name(self, row, columns_map):
        if columns_map['маркетолог'] and pd.notna(row[columns_map['маркетолог']]):
            val = str(row[columns_map['маркетолог']]).strip()
            if val and val != 'nan' and val != 'None':
                return val
        return None
    
    def get_responsible_name(self, row, columns_map):
        if columns_map['ответственный'] and pd.notna(row[columns_map['ответственный']]):
            val = str(row[columns_map['ответственный']]).strip()
            if val and val != 'nan' and val != 'None':
                return val
        return None
    
    def get_status_category(self, status_text):
        if pd.isna(status_text):
            return 'unknown'
        status_lower = str(status_text).lower().strip()
        for category, keywords in self.statuses.items():
            for keyword in keywords:
                if keyword in status_lower:
                    return category
        return 'in_work'
    
    def is_lead_transferred(self, row, columns_map):
        marketer = self.get_dispatcher_name(row, columns_map)
        responsible = self.get_responsible_name(row, columns_map)
        
        if not marketer:
            return False, False
        
        status_text = row[columns_map['этап']] if columns_map['этап'] else ''
        status_category = self.get_status_category(status_text)
        
        if status_category == 'transferred':
            if marketer == responsible:
                return True, True
            else:
                return True, False
        
        if responsible and marketer != responsible:
            return True, False
        
        return False, False
    
    def predict_with_regression(self, historical_data):
        if len(historical_data) < 2:
            return historical_data[-1] if historical_data else 50
        
        X = np.array(range(len(historical_data))).reshape(-1, 1)
        y = np.array(historical_data)
        
        model = LinearRegression()
        model.fit(X, y)
        
        next_period = np.array([[len(historical_data)]])
        prediction = model.predict(next_period)[0]
        
        return max(0, min(100, prediction))
    
    def analyze_dispatcher_from_raw(self, df: pd.DataFrame, selected_dispatchers: List[str] = None) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, Dict]:
        columns_map = self.detect_columns(df)
        
        if not columns_map['маркетолог']:
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), {'error': 'Не найдена колонка с маркетологами'}
        
        if not columns_map['этап']:
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), {'error': 'Не найдена колонка со статусами'}
        
        if columns_map['дата_создания']:
            df[columns_map['дата_создания']] = pd.to_datetime(df[columns_map['дата_создания']], errors='coerce')
        
        # Проблемные лиды (не сменили ответственного)
        transferred_with_issue = []
        
        for _, row in df.iterrows():
            is_transferred, has_issue = self.is_lead_transferred(row, columns_map)
            if is_transferred and has_issue:
                client_info = {
                    'Дата': row[columns_map['дата_создания']] if columns_map['дата_создания'] else None,
                    'Клиент': row[columns_map['клиент']] if columns_map['клиент'] else 'Не указан',
                    'Телефон': row[columns_map['телефон']] if columns_map['телефон'] else 'Не указан',
                    'Название': row[columns_map['название']] if columns_map['название'] else 'Не указан',
                    'Лэнд': row[columns_map['ленд']] if columns_map['ленд'] else 'Не указан',
                    'Маркетолог': self.get_dispatcher_name(row, columns_map),
                    'Ответственный': self.get_responsible_name(row, columns_map)
                }
                transferred_with_issue.append(client_info)
        
        issues_df = pd.DataFrame(transferred_with_issue)
        
        # Нераспределенные лиды (пустое поле Маркетолог)
        unassigned_df = df[df[columns_map['маркетолог']].isna() | (df[columns_map['маркетолог']].astype(str).str.strip() == '')]
        unassigned_count = len(unassigned_df)
        total_leads_all = len(df)
        unassigned_rate = (unassigned_count / total_leads_all * 100) if total_leads_all > 0 else 0
        
        # Детали незакрепленных лидов
        unassigned_leads = []
        for _, row in unassigned_df.iterrows():
            lead_info = {
                'Дата': row[columns_map['дата_создания']] if columns_map['дата_создания'] else None,
                'Клиент': row[columns_map['клиент']] if columns_map['клиент'] else 'Не указан',
                'Телефон': row[columns_map['телефон']] if columns_map['телефон'] else 'Не указан',
                'Название': row[columns_map['название']] if columns_map['название'] else 'Не указан',
                'Лэнд': row[columns_map['ленд']] if columns_map['ленд'] else 'Не указан',
                'Статус': row[columns_map['этап']] if columns_map['этап'] else 'Не указан'
            }
            unassigned_leads.append(lead_info)
        
        unassigned_df_detailed = pd.DataFrame(unassigned_leads)
        
        # Фильтрация по выбранным диспетчерам
        all_dispatchers = df[columns_map['маркетолог']].dropna().unique()
        all_dispatchers = [str(d).strip() for d in all_dispatchers if pd.notna(d) and str(d).strip()]
        
        if selected_dispatchers:
            df = df[df[columns_map['маркетолог']].isin(selected_dispatchers)]
        
        results = []
        daily_details = []
        
        for dispatcher in df[columns_map['маркетолог']].unique():
            if pd.isna(dispatcher) or str(dispatcher).strip() == '':
                continue
            
            dispatcher_df = df[df[columns_map['маркетолог']] == dispatcher]
            
            total_leads = len(dispatcher_df)
            transferred_leads = 0
            failed_calls = 0
            rejected = 0
            duplicates = 0
            recalls = 0
            wrong_numbers = 0
            
            historical_transfer_rates = []
            
            if columns_map['дата_создания']:
                daily_groups = dispatcher_df.groupby(dispatcher_df[columns_map['дата_создания']].dt.date)
                
                for date, day_df in daily_groups:
                    total = len(day_df)
                    transferred = 0
                    failed = 0
                    rejected_day = 0
                    
                    for _, row in day_df.iterrows():
                        if self.is_lead_transferred(row, columns_map)[0]:
                            transferred += 1
                        
                        status = self.get_status_category(row[columns_map['этап']] if columns_map['этап'] else '')
                        if status == 'failed_call':
                            failed += 1
                        elif status == 'refused':
                            rejected_day += 1
                    
                    rate = (transferred / total * 100) if total > 0 else 0
                    historical_transfer_rates.append(rate)
                    
                    daily_details.append({
                        'Сотрудник': dispatcher,
                        'Дата': date,
                        'Получено лидов': total,
                        'Передано': transferred,
                        'Недозвоны': failed,
                        'Отказы': rejected_day,
                        '% переданных': round(rate, 1)
                    })
            
            for _, row in dispatcher_df.iterrows():
                is_transferred, _ = self.is_lead_transferred(row, columns_map)
                if is_transferred:
                    transferred_leads += 1
                
                status_text = row[columns_map['этап']] if columns_map['этап'] else ''
                status_category = self.get_status_category(status_text)
                
                if status_category == 'failed_call':
                    failed_calls += 1
                elif status_category == 'refused':
                    rejected += 1
                elif status_category == 'duplicate':
                    duplicates += 1
                elif status_category == 'recall':
                    recalls += 1
                elif status_category == 'wrong_number':
                    wrong_numbers += 1
            
            transfer_rate = (transferred_leads / total_leads * 100) if total_leads > 0 else 0
            failed_rate = (failed_calls / total_leads * 100) if total_leads > 0 else 0
            rejected_rate = (rejected / total_leads * 100) if total_leads > 0 else 0
            duplicate_rate = ((duplicates + recalls) / total_leads * 100) if total_leads > 0 else 0
            
            if len(historical_transfer_rates) >= 2:
                predicted_transfer = self.predict_with_regression(historical_transfer_rates)
            else:
                if transfer_rate < self.targets['dispatcher']['transfer_rate']:
                    predicted_transfer = transfer_rate + (self.targets['dispatcher']['transfer_rate'] - transfer_rate) * 0.3
                else:
                    predicted_transfer = transfer_rate * 0.97
            
            # Загруженность
            if columns_map['дата_создания']:
                try:
                    dates = pd.to_datetime(dispatcher_df[columns_map['дата_создания']])
                    if len(dates) > 1:
                        days_range = (dates.max() - dates.min()).days
                        if days_range > 0:
                            leads_per_day = total_leads / days_range
                        else:
                            leads_per_day = total_leads
                    else:
                        leads_per_day = total_leads
                except:
                    leads_per_day = total_leads
            else:
                leads_per_day = total_leads
            
            if leads_per_day >= 20:
                load_status = 'Норма ✅'
            elif leads_per_day >= 10:
                load_status = 'Недогруз ⚠️'
            else:
                load_status = 'Критическая недогрузка 🔴'
            
            is_problem = (
                transfer_rate < self.targets['dispatcher']['transfer_rate'] or
                failed_rate > self.targets['dispatcher']['max_failed_calls'] or
                rejected_rate > self.targets['dispatcher']['max_rejected'] or
                duplicate_rate > self.targets['dispatcher']['max_duplicate_rate']
            )
            
            results.append({
                'Сотрудник': dispatcher,
                'Получено лидов': total_leads,
                'Передано продажам': transferred_leads,
                'Недозвоны': failed_calls,
                'Отказы': rejected,
                'Дубли/Повторы': duplicates + recalls,
                'Ошибка номера': wrong_numbers,
                '% переданных': round(transfer_rate, 1),
                '% недозвонов': round(failed_rate, 1),
                '% отказов': round(rejected_rate, 1),
                '% дублей/повторов': round(duplicate_rate, 1),
                'Лидов в день': round(leads_per_day, 1),
                'Загруженность': load_status,
                '📈 ПРОГНОЗ: % переданных': round(predicted_transfer, 1),
                '📈 ПРОГНОЗ: передач (шт)': round(total_leads * predicted_transfer / 100, 0),
                'is_problem': is_problem
            })
        
        summary_df = pd.DataFrame(results)
        daily_df = pd.DataFrame(daily_details)
        
        # Анализ лэндов
        landings_stats = []
        
        if columns_map['ленд']:
            for landing in df[columns_map['ленд']].dropna().unique():
                landing_df = df[df[columns_map['ленд']] == landing]
                total = len(landing_df)
                transferred = 0
                poor = 0
                
                for _, row in landing_df.iterrows():
                    if self.is_lead_transferred(row, columns_map)[0]:
                        transferred += 1
                    
                    status = self.get_status_category(row[columns_map['этап']] if columns_map['этап'] else '')
                    if status in self.poor_quality_statuses:
                        poor += 1
                
                conv = (transferred / total * 100) if total > 0 else 0
                poor_rate = (poor / total * 100) if total > 0 else 0
                
                if conv >= 70 and poor_rate < 15:
                    quality = 'Отлично'
                elif conv >= 50 and poor_rate < 30:
                    quality = 'Средне'
                else:
                    quality = 'Плохо'
                
                landings_stats.append({
                    'Лэнд': landing,
                    'Всего лидов': total,
                    'Передано': transferred,
                    'Конверсия %': round(conv, 1),
                    '% проблемных': round(poor_rate, 1),
                    'Оценка': quality
                })
        
        landings_df = pd.DataFrame(landings_stats)
        
        unassigned_info = {
            'count': unassigned_count,
            'rate': round(unassigned_rate, 1),
            'is_problem': unassigned_rate > self.targets['dispatcher']['max_unassigned_rate'],
        }
        
        return summary_df, daily_df, landings_df, {
            'unassigned': unassigned_info,
            'unassigned_df': unassigned_df_detailed,
            'issues': issues_df,
            'all_dispatchers': all_dispatchers,
            'type': 'raw'
        }
    
    def analyze_data(self, df: pd.DataFrame, selected_dispatchers: List[str] = None) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, Dict]:
        if self.is_aggregated_data(df):
            return self.analyze_aggregated_table(df)
        else:
            return self.analyze_dispatcher_from_raw(df, selected_dispatchers)
    
    def generate_email_report(self, summary_df: pd.DataFrame, landings_df: pd.DataFrame, extra_info: Dict = None) -> Tuple[str, bool]:
        has_problems = False
        email_lines = []
        email_lines.append("📊 Отчет AI-агента по эффективности сотрудников")
        email_lines.append(f"📅 {datetime.now().strftime('%d.%m.%Y %H:%M')}")
        email_lines.append("=" * 50)
        
        if extra_info and 'unassigned' in extra_info:
            unassigned = extra_info['unassigned']
            if unassigned.get('is_problem', False):
                has_problems = True
                email_lines.append(f"\n⚠️ {unassigned['count']} лидов ({unassigned['rate']}%) не закреплены за диспетчерами!")
        
        if not summary_df.empty:
            problem_emps = summary_df[summary_df['is_problem'] == True]
            if len(problem_emps) > 0:
                has_problems = True
                email_lines.append("\n⚠️ ПРОБЛЕМНЫЕ СОТРУДНИКИ:")
                for _, emp in problem_emps.iterrows():
                    email_lines.append(f"  • {emp['Сотрудник']}:")
                    if emp['% переданных'] < 70:
                        email_lines.append(f"    - Низкий % переданных: {emp['% переданных']}%")
                    if emp['% недозвонов'] > 15:
                        email_lines.append(f"    - Много недозвонов: {emp['% недозвонов']}%")
        
        return "\n".join(email_lines), has_problems