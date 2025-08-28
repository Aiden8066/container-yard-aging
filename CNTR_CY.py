import os
import re
import sys
import numpy as np
import pandas as pd
import sqlite3
import win32com.client
import qdarkstyle
import matplotlib
matplotlib.use('Qt5Agg')
import matplotlib.pyplot as plt
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout, QTabWidget,
    QWidget, QFileDialog, QMessageBox, QDialog, QLineEdit, QComboBox,
    QTableView, QHBoxLayout, QTableWidget, QTableWidgetItem, QGridLayout, QLabel, QInputDialog, QDateEdit,
    QDialogButtonBox, QTextEdit, QAction, QScrollArea, QMenu, QListWidgetItem, QAbstractItemView, QListWidget,
    QGroupBox, QHeaderView, QCheckBox, QProgressBar, QColorDialog)
from PyQt5.QtCore import Qt, QSortFilterProxyModel, QDate, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QStandardItemModel, QStandardItem, QCursor, QColor
from datetime import datetime, timedelta
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure 
from matplotlib.ticker import FuncFormatter
import warnings
import time
import io
import csv
from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import QShortcut, QMenu
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')
matplotlib.rc('font', family='Malgun Gothic')
matplotlib.rcParams['axes.unicode_minus'] = False  # 마이너스 부호 깨짐 방지


# 데이터베이스 파일 경로 설정
#db_file = r"C:\Users\SyncthingServiceAcct\test\aiden\master_database.db"
db_file = r"\\10.193.232.18\Java\우현 테스트\master_database.db"
#contact_db_path = r"C:\Users\SyncthingServiceAcct\test\aiden\Vessel contact.db"
contact_db_path = r"\\10.193.232.18\Java\우현 테스트\Vessel contact.db"




class LoadingWorker(QThread):
    progress = pyqtSignal(int)
    
    def run(self):
        try:
            # 전체 작업량 계산
            total_steps = 100
            current_step = 0
            
            # 1. 데이터베이스 연결 (30%)
            conn = connect_db()
            for i in range(30):
                self.progress.emit(current_step)
                current_step += 1
                time.sleep(0.02)
            
            # 2. 테이블 정보 가져오기 (30%)
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = cursor.fetchall()
            for i in range(30):
                self.progress.emit(current_step)
                current_step += 1
                time.sleep(0.02)
            
            # 3. 초기 데이터 로드 (40%)
            initial_data = {}
            steps_per_table = 40 // (len(tables) or 1)
            for table in tables:
                table_name = table[0]
                cursor.execute(f"SELECT * FROM {table_name} LIMIT 1")
                initial_data[table_name] = cursor.description
                self.progress.emit(current_step)
                current_step += steps_per_table
                time.sleep(0.02)
            
            # 남은 진행률을 100%까지 채움
            while current_step <= 100:
                self.progress.emit(current_step)
                current_step += 1
                time.sleep(0.02)
            
        except Exception as e:
            print(f"Error in worker thread: {str(e)}")

class LoadingScreen(QWidget):
    def __init__(self):
        super().__init__()
        self.setFixedSize(200, 100)
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.CustomizeWindowHint)
        
        # UI 설정
        layout = QVBoxLayout()
        
        self.progress = QProgressBar()
        self.progress.setStyleSheet("""
            QProgressBar {
                border: 2px solid grey;
                border-radius: 5px;
                text-align: center;
                background-color: #19232D;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                width: 10px;
            }
        """)
        layout.addWidget(self.progress)
        
        self.label = QLabel("Loading...")
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("""
            QLabel {
                color: #333;
                font-size: 16px;
                font-weight: bold;
            }
        """)
        layout.addWidget(self.label)
        
        self.setLayout(layout)
        
        # 작업자 스레드 설정
        self.worker = LoadingWorker()
        self.worker.progress.connect(self.on_progress)
        self.worker.start()

    def on_progress(self, value):
        self.progress.setValue(value)
        if value >= 100:
            self.start_main_program()

    def start_main_program(self):
        try:
            self.main_window = MainWindow()
            self.main_window.show()
            self.close()
        except Exception as e:
            QMessageBox.critical(None, "Error", f"Failed to start main program: {str(e)}")
            self.close()



class KPICalculator:
    def __init__(self, tab_info=None):  # tab_info 파라미터 추가
        self.storage_utils = StorageUtils()
        self.tab_info = tab_info  # tab_info 저장
        
    def calculate_container_bonus(self, container_count):
        """컨테이너 수에 따른 가산점 계산"""
        if container_count >= 300:
            return 15
        elif container_count >= 200:
            return 10
        elif container_count >= 150:
            return 8
        elif container_count >= 100:
            return 5
        elif container_count >= 50:
            return 2
        return 0

    def calculate_division_kpi(self, table_name, target_month):
        try:
            # container_count 데이터 먼저 가져오기
            container_data = StorageUtils.get_kpi_container_count(table_name)
            if container_data.empty:
                return None
                
            # storage cost 데이터 가져오기
            storage_data = self.storage_utils.get_monthly_storage_data(table_name)
            if storage_data.empty:
                return None

            # 데이터 병합
            df = pd.merge(storage_data, container_data[['month', 'container_count']], 
                         on='month', how='left')

            if df.empty:
                return None
                
            current_data = df[df['month'] == target_month]
            if current_data.empty:
                return None
                
            current_cost = current_data.iloc[0]['total_storage_cost']
            
            # MoM 변화율 계산
            prev_month_data = df[df['month'] == (target_month - pd.DateOffset(months=1))]
            if prev_month_data.empty or prev_month_data.iloc[0]['total_storage_cost'] == 0:
                mom_change = 0 if current_cost == 0 else 100
            else:
                mom_change = ((current_cost - prev_month_data.iloc[0]['total_storage_cost']) / 
                             prev_month_data.iloc[0]['total_storage_cost'] * 100)
            
            # YoY 변화율 계산
            prev_year_data = df[df['month'] == (target_month - pd.DateOffset(years=1))]
            if prev_year_data.empty or prev_year_data.iloc[0]['total_storage_cost'] == 0:
                yoy_change = 0 if current_cost == 0 else 100
            else:
                prev_year_cost = prev_year_data.iloc[0]['total_storage_cost']
                yoy_change = ((current_cost - prev_year_cost) / prev_year_cost * 100)
            
            # 모든 테이블의 MoM, YoY 변화율 수집
            tables = [f'table{i}' for i in range(1, 12)]
            all_mom_changes = []
            all_yoy_changes = []
            
            for table in tables:
                df = self.storage_utils.get_monthly_storage_data(table)
                if not df.empty:
                    current_data = df[df['month'] == target_month]
                    if not current_data.empty:
                        current_cost = current_data.iloc[0]['total_storage_cost']
                        
                        # MoM 변화율 계산
                        prev_month_data = df[df['month'] == (target_month - pd.DateOffset(months=1))]
                        if prev_month_data.empty or prev_month_data.iloc[0]['total_storage_cost'] == 0:
                            mom_change_temp = 0 if current_cost == 0 else 100
                        else:
                            mom_change_temp = ((current_cost - prev_month_data.iloc[0]['total_storage_cost']) / 
                                        prev_month_data.iloc[0]['total_storage_cost'] * 100)
                        all_mom_changes.append(mom_change_temp)
                        
                        # YoY 변화율 계산
                        prev_year_data = df[df['month'] == (target_month - pd.DateOffset(years=1))]
                        if prev_year_data.empty or prev_year_data.iloc[0]['total_storage_cost'] == 0:
                            yoy_change_temp = 0 if current_cost == 0 else 100
                        else:
                            yoy_change_temp = ((current_cost - prev_year_data.iloc[0]['total_storage_cost']) / 
                                        prev_year_data.iloc[0]['total_storage_cost'] * 100)
                        all_yoy_changes.append(yoy_change_temp)
            
            # 점수 계산
            mom_score = self._calculate_mom_score(mom_change, all_mom_changes)
            yoy_score = self._calculate_yoy_score(yoy_change, all_yoy_changes)
            cost_per_container_data = self.calculate_cost_per_container_rank(target_month)
            trend_score = self.calculate_trend_score(table_name, target_month)
            
            # 컨테이너 수 데이터 가져오기
            container_df = self.storage_utils.get_kpi_container_count(table_name)
            current_month_data = container_df[container_df['month'] == target_month]
            
            if not current_month_data.empty:
                container_count = current_month_data.iloc[0]['container_count']
            else:
                container_count = 0
                
            
            
            # 컨테이너 수에 따른 가산점 계산
            container_bonus = self.calculate_container_bonus(container_count)
            
            # 총점에 가산점 추가
            total_score = (mom_score + yoy_score + 
                          cost_per_container_data.get(table_name, {}).get('score', 0) + 
                          trend_score + container_bonus)
            
            return {
                'total_score': total_score,
                'grade': self._calculate_grade(total_score),
                'mom_change': float(mom_change),
                'mom_score': mom_score,
                'yoy_change': float(yoy_change),
                'yoy_score': yoy_score,
                'cost_per_container_score': cost_per_container_data.get(table_name, {}).get('score', 0),
                'trend_score': trend_score,
                'current_cost': float(current_cost),
                'cost_per_container_rank': cost_per_container_data.get(table_name, {}).get('rank', 0),
                'cost_per_container': float(cost_per_container_data.get(table_name, {}).get('cost_per_container', 0)),
                'container_count': int(container_count),
                'container_bonus': container_bonus
            }
                
        except Exception as e:
            
            return None

    def calculate_cost_per_container_rank(self, target_month):
        try:
            tables = [f'table{i}' for i in range(1, 12)]
            division_costs = []

            # 데이터 수집
            for table in tables:
                df = self.storage_utils.get_monthly_storage_data(table)
                if not df.empty:
                    current_data = df[df['month'] == target_month]
                    if not current_data.empty:
                        cost = current_data.iloc[0]['total_storage_cost']
                        count = current_data.iloc[0]['container_count']
                        if count > 0:
                            cost_per_container = cost / count
                            division_costs.append((table, cost_per_container))

            # 비용 기준으로 정렬
            division_costs.sort(key=lambda x: x[1])
            
            # 동점자 처리를 위한 순위 및 점수 계산
            results = {}
            current_rank = 1
            prev_cost = None
            rank_scores = {1: 25, 2: 22.5, 3: 20, 4: 17.5, 5: 15, 6: 12.5, 7: 10, 8: 7.5, 9: 5, 10: 2.5}
            
            for i, (division, cost) in enumerate(division_costs):
                # 이전 비용과 같으면 같은 순위 부여
                if prev_cost is not None and cost != prev_cost:
                    current_rank = i + 1
                
                score = rank_scores.get(current_rank, 0)
                results[division] = {
                    'rank': current_rank,
                    'score': score,
                    'cost_per_container': cost
                }
                prev_cost = cost
                
            return results
                
        except Exception as e:
            print(f"Error calculating cost per container rank: {str(e)}")
            return {}

    def calculate_trend_score(self, table_name, target_month):
        """추세 분석 점수 계산"""
        try:
            # 12개월 데이터 가져오기 시도
            end_date = target_month
            start_date = end_date - pd.DateOffset(months=11)
            
            df = self.storage_utils.get_monthly_storage_data(table_name)
            df = df[(df['month'] >= start_date) & (df['month'] <= end_date)]
            
            if len(df) < 2:  # 최소 2개월 이상의 데이터가 필요
                return 5     # 비교 불가능한 경우 0점
            
            # 가용 데이터 내에서 증가 횟수 계산
            increases = 0
            for i in range(1, len(df)):
                if df.iloc[i]['total_storage_cost'] > df.iloc[i-1]['total_storage_cost']:
                    increases += 1
            
            # 동일한 점수 매핑 로직 적용
            if increases <= 3:     # 0~3회 증가
                return 15
            elif increases >= 11:  # 11회 이상 증가
                return 7
            else:
                # 4~10회 증가에 대한 점수 매핑
                score_mapping = {
                    4: 14,   
                    5: 13,   
                    6: 12,   
                    7: 11,   
                    8: 10,   
                    9: 9,   
                    10: 8,  
                }
                return score_mapping.get(increases, 0)
                
        except Exception as e:
            print(f"Error calculating trend score: {str(e)}")
            return 0
            
    def _calculate_mom_score(self, change_percent, all_changes):
        """MoM 변화율에 따른 점수 계산 (35점 만점)"""
        if not all_changes:
            return 17.5  # 비교 데이터가 없는 경우 중간 점수
        
        # 현재 change_percent를 all_changes에 추가
        if change_percent not in all_changes:
            all_changes.append(change_percent)
        
        # 변화율 오름차순 정렬 (-100%가 가장 앞에 오도록)
        sorted_changes = sorted(all_changes)  # reverse=True 제거
        total_divisions = len(sorted_changes)
        
        if total_divisions == 1:
            return 35  # 단일 부서인 경우 만점
        
        # 동점자 처리를 위한 순위 및 점수 계산
        rank_scores = {}  # 각 변화율에 대한 점수를 저장할 딕셔너리
        current_rank = 1
        prev_change = None
        
        for i, change in enumerate(sorted_changes):
            # 이전 값과 다른 경우에만 순위 증가
            if prev_change is not None and change != prev_change:
                current_rank = i + 1
                
            # 점수 계산 (1등: 35점, 꼴등: 17.5점)
            score = 35 - ((current_rank - 1) * (17.5 / (total_divisions - 1)))
            rank_scores[change] = score
            prev_change = change
        
        return rank_scores[change_percent]

    def _calculate_yoy_score(self, change_percent, all_changes):
        """YoY 변화율에 따른 점수 계산 (25점 만점)"""
        if not all_changes:
            return 12.5  # 비교 데이터가 없는 경우 중간 점수
        
        # 현재 change_percent를 all_changes에 추가
        if change_percent not in all_changes:
            all_changes.append(change_percent)
        
        # 변화율 오름차순 정렬 (-100%가 가장 앞에 오도록)
        sorted_changes = sorted(all_changes)  # reverse=True 제거
        total_divisions = len(sorted_changes)
        
        if total_divisions == 1:
            return 25  # 단일 부서인 경우 만점
        
        # 동점자 처리를 위한 순위 및 점수 계산
        rank_scores = {}  # 각 변화율에 대한 점수를 저장할 딕셔너리
        current_rank = 1
        prev_change = None
        
        for i, change in enumerate(sorted_changes):
            # 이전 값과 다른 경우에만 순위 증가
            if prev_change is not None and change != prev_change:
                current_rank = i + 1
                
            # 점수 계산 (1등: 25점, 꼴등: 12.5점)
            score = 25 - ((current_rank - 1) * (12.5 / (total_divisions - 1)))
            rank_scores[change] = score
            prev_change = change
        
        return rank_scores[change_percent]

    def _calculate_grade(self, total_score):
        """점수에 따른 등급 계산"""
        if total_score >= 90:
            return 'A'
        elif total_score >= 80:
            return 'B'
        elif total_score >= 70:
            return 'C'
        elif total_score >= 60:
            return 'D'
        else:
            return 'F'

class CustomTableWidgetItem(QTableWidgetItem):
    def __init__(self, value):
        super().__init__(str(value))
        self.value = value

    def __lt__(self, other):
        if isinstance(self.value, (int, float)) and isinstance(other.value, (int, float)):
            return self.value < other.value
        return super().__lt__(other)

class KPIWindow(QDialog):
    def __init__(self, parent=None, tab_info=None):
        super().__init__(parent)
        self.setWindowTitle("Division KPI Dashboard - 2024 Ver.")
        self.setGeometry(100, 100, 1400, 800)
        self.setWindowFlags(Qt.Window | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)

        # tab_info 초기화
        self.tab_info = tab_info
        self.table_to_tab = {}
        if tab_info:
            self.table_to_tab = {v: k for k, v in tab_info.items()}
        
        # KPI 계산기 초기화 시 tab_info 전달
        self.kpi_calculator = KPICalculator(tab_info=self.tab_info)
        
        # 날짜 설정 초기화
        self._init_date_selectors()
        
        # UI 구성
        self._init_ui()

    def _init_date_selectors(self):
        current_date = QDate.currentDate()
        last_month = current_date.addMonths(-1)
        
        # Division KPI용 날짜 선택기
        self.month_selector = QDateEdit()
        self.month_selector.setDisplayFormat("yyyy-MM")
        self.month_selector.setDate(last_month)
        self.month_selector.setMaximumDate(last_month)
        self.month_selector.setKeyboardTracking(False)
        
        # Monthly Average용 날짜 선택기
        self.start_month_selector = QDateEdit()
        self.start_month_selector.setDisplayFormat("yyyy-MM")
        self.start_month_selector.setDate(last_month.addMonths(-11))
        self.start_month_selector.setMaximumDate(last_month)
        self.start_month_selector.setKeyboardTracking(False)
        
        self.end_month_selector = QDateEdit()
        self.end_month_selector.setDisplayFormat("yyyy-MM")
        self.end_month_selector.setDate(last_month)
        self.end_month_selector.setMaximumDate(last_month)
        self.end_month_selector.setKeyboardTracking(False)

    def _init_ui(self):
        # 메인 레이아웃
        main_layout = QVBoxLayout()
        
        # === 상단 KPI 섹션 ===
        top_group = QGroupBox("Division KPI")
        top_layout = QVBoxLayout()
        
        # KPI 테이블 헤더 영역 (월 선택기와 Export 버튼)
        kpi_header_layout = QHBoxLayout()
        
        # 월 선택 컨트롤
        month_control = QHBoxLayout()
        month_control.addWidget(QLabel("Select Month:"))
        month_control.addWidget(self.month_selector)
        
        # KPI Run 버튼
        kpi_run_btn = QPushButton("Run")
        kpi_run_btn.clicked.connect(self.update_kpi_table_only)
        month_control.addWidget(kpi_run_btn)
        
        # 월 선택 컨트롤을 헤더 레이아웃에 추가
        kpi_header_layout.addLayout(month_control)
        
        # Export 버튼 (오른쪽 정렬)
        export_kpi_btn = QPushButton("Export to Excel")
        export_kpi_btn.clicked.connect(lambda: self.export_to_excel(self.kpi_table, "KPI"))
        export_kpi_btn.setStyleSheet("""
            QPushButton {
                background-color: #2D3F50;
                color: white;
                padding: 5px 10px;
                border: none;
                border-radius: 3px;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #34495E;
            }
        """)
        kpi_header_layout.addStretch()  # 중간 공간을 채움
        kpi_header_layout.addWidget(export_kpi_btn)
        
        # KPI 테이블 설정
        self.kpi_table = QTableWidget()
        self.kpi_table.setColumnCount(13)  # 컬럼 수 13개로 증가
        self.kpi_table.setHorizontalHeaderLabels([
            "Division", "Total Score", "Grade", "MoM Change(%)",
            "MoM Score(35)", "YoY Change(%)", "YoY Score(25)",
            "Cost/Cont", "Rank", "Score(25)", "Trend Score(15)",
            "Container Count", "Bonus Score(15)"  # 새로운 컬럼 추가
        ])
        
        # 레이아웃에 추가
        top_layout.addLayout(kpi_header_layout)
        top_layout.addWidget(self.kpi_table)
        top_group.setLayout(top_layout)
        
        # === 하단 Monthly Average 섹션 ===
        bottom_group = QGroupBox("Monthly Average Scores")
        bottom_layout = QVBoxLayout()
        
        # Monthly Average 헤더 영역 (기간 선택기와 Export 버튼)
        avg_header_layout = QHBoxLayout()
        
        # 기간 선택 컨트롤
        range_control = QHBoxLayout()
        range_control.addWidget(QLabel("Start Month:"))
        range_control.addWidget(self.start_month_selector)
        range_control.addWidget(QLabel("End Month:"))
        range_control.addWidget(self.end_month_selector)
        
        # Calculate Average 버튼
        calculate_avg_btn = QPushButton("Calculate Average")
        calculate_avg_btn.clicked.connect(self.update_monthly_averages)
        range_control.addWidget(calculate_avg_btn)
        
        # 기간 선택 컨트롤을 헤더 레이아웃에 추가
        avg_header_layout.addLayout(range_control)
        
        # Export 버튼 (오른쪽 정렬)
        export_avg_btn = QPushButton("Export to Excel")
        export_avg_btn.clicked.connect(lambda: self.export_to_excel(self.monthly_avg_table, "Monthly_Average"))
        export_avg_btn.setStyleSheet("""
            QPushButton {
                background-color: #2D3F50;
                color: white;
                padding: 5px 10px;
                border: none;
                border-radius: 3px;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #34495E;
            }
        """)
        avg_header_layout.addStretch()  # 중간 공간을 채움
        avg_header_layout.addWidget(export_avg_btn)
        
        # Monthly Average 테이블 설정
        self.monthly_avg_table = QTableWidget()
        self.monthly_avg_table.setColumnCount(6)
        self.monthly_avg_table.setHorizontalHeaderLabels([
            "Division", 
            "Avg Total Score",
            "Avg MoM Score(35)",
            "Avg YoY Score(25)",
            "Avg Cost/Cont Score(25)",
            "Avg Trend Score(15)"
        ])
        
        # 레이아웃에 추가
        bottom_layout.addLayout(avg_header_layout)
        bottom_layout.addWidget(self.monthly_avg_table)
        bottom_group.setLayout(bottom_layout)
        
        # 테이블 정렬 기능 활성화
        for table in [self.kpi_table, self.monthly_avg_table]:
            table.setSortingEnabled(True) 
            table.horizontalHeader().setSectionsClickable(True)
            table.setStyleSheet("""
                QHeaderView::section {
                    background-color: #2D3F50;
                    color: white;
                    padding: 5px;
                    border: 1px solid #1D2B3A;
                }
                QHeaderView::section:checked {
                    background-color: #34495E;
                }
            """)
        
        # 전체 레이아웃 구성
        main_layout.addWidget(top_group)
        main_layout.addWidget(bottom_group)
        self.setLayout(main_layout)

    def update_kpi_table_only(self):
        """Division KPI 테이블만 업데이트"""
        self.update_kpi_table()

    def update_monthly_averages(self):
        try:
            start_date = pd.to_datetime(f"{self.start_month_selector.date().year()}-{self.start_month_selector.date().month():02d}-01")
            end_date = pd.to_datetime(f"{self.end_month_selector.date().year()}-{self.end_month_selector.date().month():02d}-01")
            
            months = pd.date_range(start_date, end_date, freq='MS')
            
            # 디비전 목록 가져오기
            if self.tab_info:
                divisions = list(self.tab_info.values())
            else:
                conn = connect_db()
                cursor = conn.cursor()
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
                divisions = [table[0] for table in cursor.fetchall()]
                conn.close()
            
            # 각 디비전별 평균 계산
            division_averages = []
            for division in divisions:
                scores = {
                    'total_score': [],
                    'mom_score': [],
                    'yoy_score': [],
                    'cost_per_container_score': [],
                    'trend_score': []
                }
                
                for month in months:
                    kpi_data = self.kpi_calculator.calculate_division_kpi(division, month)
                    if kpi_data:
                        scores['total_score'].append(kpi_data['total_score'])
                        scores['mom_score'].append(kpi_data['mom_score'])
                        scores['yoy_score'].append(kpi_data['yoy_score'])
                        scores['cost_per_container_score'].append(kpi_data['cost_per_container_score'])
                        scores['trend_score'].append(kpi_data['trend_score'])
                
                if scores['total_score']:  # 점수가 있는 경우에만 평균 계산
                    avg_scores = {
                        'division': self.table_to_tab.get(division, division),
                        'total_score': sum(scores['total_score']) / len(scores['total_score']),
                        'mom_score': sum(scores['mom_score']) / len(scores['mom_score']),
                        'yoy_score': sum(scores['yoy_score']) / len(scores['yoy_score']),
                        'cost_score': sum(scores['cost_per_container_score']) / len(scores['cost_per_container_score']),
                        'trend_score': sum(scores['trend_score']) / len(scores['trend_score'])
                    }
                    division_averages.append(avg_scores)
            
            # 테이블 업데이트
            self.monthly_avg_table.setSortingEnabled(False)  # 정렬 임시 비활성화
            self.monthly_avg_table.setRowCount(len(division_averages))
            
            for row, avg_scores in enumerate(division_averages):
                # Division 열
                division_item = CustomTableWidgetItem(avg_scores['division'])
                division_item.setTextAlignment(Qt.AlignCenter)
                self.monthly_avg_table.setItem(row, 0, division_item)
                
                # 점수 열들
                scores = [
                    avg_scores['total_score'],
                    avg_scores['mom_score'],
                    avg_scores['yoy_score'],
                    avg_scores['cost_score'],
                    avg_scores['trend_score']
                ]
                
                for col, score in enumerate(scores, 1):
                    item = CustomTableWidgetItem(score)
                    item.setData(Qt.DisplayRole, f"{score:.1f}")
                    item.setTextAlignment(Qt.AlignCenter)
                    self.monthly_avg_table.setItem(row, col, item)
            
            self.monthly_avg_table.setSortingEnabled(True)  # 정렬 다시 활성화
            self.monthly_avg_table.resizeColumnsToContents()
            
        except Exception as e:
            print(f"Debug - Error in update_monthly_averages: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to update monthly averages: {str(e)}")

    def update_kpi_table(self):
        try:
            selected_date = self.month_selector.date()
            selected_month = pd.to_datetime(f"{selected_date.year()}-{selected_date.month():02d}-01")
        
            if self.tab_info:
                divisions = list(self.tab_info.values())
            else:
                conn = connect_db()
                cursor = conn.cursor()
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
                divisions = [table[0] for table in cursor.fetchall()]
                conn.close()
        
            self.kpi_table.setRowCount(len(divisions))
        
            # 정렬 기능 임시 비활성화 (데이터 입력 중 자동 정렬 방지)
            self.kpi_table.setSortingEnabled(False)
            
            for row, division in enumerate(divisions):
                kpi_data = self.kpi_calculator.calculate_division_kpi(division, selected_month)
                
                if kpi_data:

                    
                    # Division 열
                    display_name = self.table_to_tab.get(division, division)
                    self.kpi_table.setItem(row, 0, CustomTableWidgetItem(display_name))
                    
                    # 숫자 데이터 처리
                    numeric_items = [
                        (1, kpi_data['total_score']),
                        (2, kpi_data['grade']),
                        (3, kpi_data['mom_change']),
                        (4, kpi_data['mom_score']),
                        (5, kpi_data['yoy_change']),
                        (6, kpi_data['yoy_score']),
                        (7, kpi_data['cost_per_container']),
                        (8, kpi_data['cost_per_container_rank']),
                        (9, kpi_data['cost_per_container_score']),
                        (10, kpi_data['trend_score']),
                        (11, kpi_data['container_count']),  # 추가
                        (12, kpi_data['container_bonus'])   # 추가
                    ]
                    
                    for col, value in numeric_items:
                        if isinstance(value, (int, float)):
                            formatted_value = f"{value:,.2f}" if col not in [2, 8] else str(value)
                            item = CustomTableWidgetItem(value)
                            item.setData(Qt.DisplayRole, formatted_value)
                        else:
                            item = CustomTableWidgetItem(value)
                        item.setTextAlignment(Qt.AlignCenter)
                        self.kpi_table.setItem(row, col, item)
            
            # 정렬 기능 다시 활성화
            self.kpi_table.setSortingEnabled(True)
            self.kpi_table.resizeColumnsToContents()
            
        except Exception as e:
            print(f"Debug - Error in update_kpi_table: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to update KPI table: {str(e)}")

    def _calculate_change_score(self, change_percent):
        if change_percent <= 0:  # 감소
            return 50
        elif change_percent <= 5:  # 5% 이내 증가
            return 35
        elif change_percent <= 10:  # 5~10% 증가
            return 20
        else:  # 10% 이상 증가
            return 0
            
    def _calculate_grade(self, total_score):
        if total_score >= 90:
            return 'A'
        elif total_score >= 80:
            return 'B'
        elif total_score >= 70:
            return 'C'
        elif total_score >= 60:
            return 'D'
        else:
            return 'F'

    def export_to_excel(self, table, table_type):
        try:
            # 현재 날짜 가져오기
            current_date = datetime.now().strftime("%Y%m%d")
            
            # 테이블 타입에 따라 파일명 다르게 생성
            if table_type == "Monthly_Average":
                # 시작월과 마감월 가져오기 
                start_month = self.start_month_selector.date().toString("yyyyMM")
                end_month = self.end_month_selector.date().toString("yyyyMM")
                default_filename = f"Storage Cost_{table_type}_{start_month}_to_{end_month}.xlsx"
            else:
                selected_month = self.month_selector.date().toString("yyyyMM")
                default_filename = f"Storage Cost_{table_type}_{selected_month}.xlsx"
            
            # 파일 저장 다이얼로그
            file_name, _ = QFileDialog.getSaveFileName(
                self,
                f"Export {table_type} Table",
                default_filename,
                "Excel Files (*.xlsx)"
            )
            
            if file_name:
                # 테이블 데이터를 DataFrame으로 변환
                df = self._table_to_dataframe(table)
                
                # DataFrame을 엑셀로 저장
                with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name=table_type, index=False)
                    
                    # 워크시트 가져오기
                    worksheet = writer.sheets[table_type]
                    
                    # 열 너비 자동 조정
                    for idx, col in enumerate(df.columns):
                        max_length = max(
                            df[col].astype(str).apply(len).max(),
                            len(str(col))
                        )
                        worksheet.column_dimensions[chr(65 + idx)].width = max_length + 2
                
                QMessageBox.information(self, "Success", f"{table_type} table has been exported successfully!")
                
        except Exception as e:
            print(f"Debug - Error in export_to_excel: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to export table: {str(e)}")
    
    def _table_to_dataframe(self, table):
        # 테이블 데이터를 DataFrame으로 변환
        data = []
        headers = []
        
        # 헤더 가져오기
        for col in range(table.columnCount()):
            headers.append(table.horizontalHeaderItem(col).text())
        
        # 데이터 가져오기
        for row in range(table.rowCount()):
            row_data = []
            for col in range(table.columnCount()):
                item = table.item(row, col)
                if item is not None:
                    # CustomTableWidgetItem의 실제 값 사용
                    if hasattr(item, 'value'):
                        row_data.append(item.value)
                    else:
                        row_data.append(item.text())
                else:
                    row_data.append("")
            data.append(row_data)
        
        # DataFrame 생성
        return pd.DataFrame(data, columns=headers)

class KPIPasswordDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("KPI Dashboard Authentication")
        self.setFixedSize(300, 150)
        self.setWindowFlags(Qt.Window | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)

        
        layout = QVBoxLayout()
        
        # 비밀번호 입력 필드
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setPlaceholderText("Enter Password")
        
        # 레이아웃에 위젯 추가
        layout.addWidget(QLabel("Please enter the password:"))
        layout.addWidget(self.password_input)
        
        # 버튼
        buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel,
            Qt.Horizontal, self)
        buttons.accepted.connect(self.check_password)
        buttons.rejected.connect(self.reject)
        
        layout.addWidget(buttons)
        self.setLayout(layout)
        
        # 비밀번호 설정 
        self.correct_password = "0103"  
        
    def check_password(self):
        if self.password_input.text() == self.correct_password:
            self.accept()
        else:
            QMessageBox.warning(self, "Error", "Wrong password")
            self.password_input.clear()




# 월별 데이터 표시용 창 클래스
class MonthlyDataWindow(QDialog):

    def __init__(self, monthly_data, parent=None):
        super().__init__(parent)
        self.setWindowTitle("월별 예상 스토리지 비용")
        self.setGeometry(100, 100, 1200, 800)

        # 탭 위젯 생성
        self.tabs = QTabWidget(self)

        # 각 월별 데이터 추가
        for month, df in monthly_data.items():
            tab = QWidget()
            layout = QVBoxLayout()
            table_widget = QTableWidget()
            
            # 테이블 설정 추가
            table_widget.setSelectionMode(QAbstractItemView.ContiguousSelection)  # 연속 선택 가능
            table_widget.setContextMenuPolicy(Qt.ActionsContextMenu)  # 컨텍스트 메뉴 활성화
            
            # 복사 액션 추가
            copy_action = QAction("Copy", table_widget)
            copy_action.setShortcut(QKeySequence.Copy)
            copy_action.triggered.connect(lambda: self.copy_selection(table_widget))
            table_widget.addAction(copy_action)
            
            layout.addWidget(table_widget)
            tab.setLayout(layout)

            # 테이블에 월별 데이터 표시
            self.display_data(df, table_widget)
            
            self.tabs.addTab(tab, month)

        # 레이아웃 설정
        layout = QVBoxLayout(self)
        layout.addWidget(self.tabs)
        self.setLayout(layout)

    def display_data(self, df, table_widget):
        table_widget.setColumnCount(len(df.columns))
        table_widget.setRowCount(len(df.index))
        table_widget.setHorizontalHeaderLabels(df.columns)

        for i in range(len(df.index)):
            for j in range(len(df.columns)):
                item = QTableWidgetItem(str(df.iloc[i, j]))
                item.setFont(QFont("Arial", 10))
                table_widget.setItem(i, j, item)

        # 테이블 크기 조정
        table_widget.resizeColumnsToContents()
        table_widget.resizeRowsToContents()

    def copy_selection(self, table_widget):
        selected = table_widget.selectedRanges()
        if not selected:
            return
        
        text = []
        for r in selected:
            for row in range(r.topRow(), r.bottomRow() + 1):
                row_data = []
                for col in range(r.leftColumn(), r.rightColumn() + 1):
                    item = table_widget.item(row, col)
                    if item is not None:
                        row_data.append(item.text())
                    else:
                        row_data.append('')
                text.append('\t'.join(row_data))
        
        QApplication.clipboard().setText('\n'.join(text))

    def keyPressEvent(self, event):
        if event.matches(QKeySequence.Copy):
            current_tab = self.tabs.currentWidget()
            if current_tab:
                table_widget = current_tab.findChild(QTableWidget)
                if table_widget:
                    self.copy_selection(table_widget)
        else:
            super().keyPressEvent(event)

class StorageUtils:
    @staticmethod
    def get_monthly_storage_data(table_name):
        """
        특정 테이블의 월별 스토리지 비용과 컨테이너 수를 계산하는 유틸리티 메서드
        """
        if table_name == "all_tables":
            # 모든 테이블의 데이터를 합치는 로직
            combined_df = pd.DataFrame()
            conn = connect_db()
            
            # 모든 테이블에서 데이터 가져오기
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = cursor.fetchall()
            
            for table in tables:
                table_name = table[0]
                query = f"""
                WITH calculated_days AS (
                    SELECT * ,
                        date(
                            CASE
                                WHEN [unloadingterminal] LIKE '%/%'
                                THEN substr([unloadingterminal], 7, 4) || '-' ||
                                     substr([unloadingterminal], 4, 2) || '-' ||
                                     substr([unloadingterminal], 1, 2)
                                ELSE [unloadingterminal]
                            END
                        ) AS unloading_terminal_date,
                        date([terminalappointment]) AS terminal_appointment_date
                    FROM {table_name}
                    WHERE container IS NOT NULL
                ),
                days_over_seven AS (
                    SELECT *,
                        julianday(terminal_appointment_date) - julianday(unloading_terminal_date) + 1 AS total_stay_days
                    FROM calculated_days
                ),
                storage_costs AS (
                    SELECT * ,
                        CASE
                            WHEN total_stay_days > 7 THEN 
                                CASE 
                                    WHEN destinationport IN ('LZO', 'LAZARO', 'Lazaro Cardenas', 'LAZARO CARDENAS') THEN
                                        CASE
                                            WHEN total_stay_days > 10 THEN (total_stay_days - 10) * 2027
                                            ELSE 0
                                        END
                                    ELSE 4356.42 + (total_stay_days - 7) * 2194.43
                                END
                            ELSE 0
                        END AS storage_cost,
                        strftime('%Y-%m', date([terminalappointment])) AS month
                    FROM days_over_seven
                )
                SELECT 
                    month, 
                    SUM(storage_cost) as total_storage_cost,
                    COUNT(*) as container_count
                FROM storage_costs
                GROUP BY month
                """
                try:
                    df = pd.read_sql(query, conn)
                    combined_df = pd.concat([combined_df, df], ignore_index=True)
                except:
                    continue
            
            conn.close()
            
            # 월별로 데이터 합치기
            if not combined_df.empty:
                combined_df = combined_df.groupby('month').agg({
                    'total_storage_cost': 'sum',
                    'container_count': 'sum'
                }).reset_index()
                
                # 날짜 형식 변환
                combined_df['month'] = pd.to_datetime(combined_df['month'], format='%Y-%m')
                
            return combined_df
            
        else:
            # 기존의 단일 테이블 처리 코드
            conn = connect_db()
            query = f"""
            WITH calculated_days AS (
                SELECT * ,
                    date(
                        CASE
                            WHEN [unloadingterminal] LIKE '%/%'
                            THEN substr([unloadingterminal], 7, 4) || '-' ||
                                 substr([unloadingterminal], 4, 2) || '-' ||
                                 substr([unloadingterminal], 1, 2)
                            ELSE [unloadingterminal]
                        END
                    ) AS unloading_terminal_date,
                    date([terminalappointment]) AS terminal_appointment_date
                FROM {table_name}
                WHERE container IS NOT NULL
            ),
            days_over_seven AS (
                SELECT *,
                    julianday(terminal_appointment_date) - julianday(unloading_terminal_date) + 1 AS total_stay_days
                FROM calculated_days
            ),
            storage_costs AS (
                SELECT * ,
                    CASE
                        WHEN destinationport IN ('LZO', 'LAZARO', 'Lazaro Cardenas', 'LAZARO CARDENAS') THEN
                            CASE
                                WHEN total_stay_days > 10 THEN total_stay_days - 10
                                ELSE 0
                            END
                        ELSE
                            CASE
                                WHEN total_stay_days > 7 THEN total_stay_days - 7
                                ELSE 0
                            END
                    END AS days_over,
                    CASE
                        WHEN total_stay_days > 7 THEN 
                            CASE 
                                WHEN destinationport IN ('LZO', 'LAZARO', 'Lazaro Cardenas', 'LAZARO CARDENAS') THEN
                                    CASE
                                        WHEN total_stay_days > 10 THEN (total_stay_days - 10) * 2027
                                        ELSE 0
                                    END
                                ELSE 4356.42 + (total_stay_days - 7) * 2194.43
                            END
                        ELSE 0
                    END AS storage_cost,
                    strftime('%Y-%m', date([terminalappointment])) AS month
                FROM days_over_seven
            )
            SELECT 
                month, 
                SUM(storage_cost) as total_storage_cost,
                COUNT(*) as container_count
            FROM storage_costs
            GROUP BY month
            ORDER BY month
            """
            df = pd.read_sql(query, conn)
            conn.close()

            # 날짜 표식 처리 - 수정된 부분
            df['month'] = pd.to_datetime(df['month'], errors='coerce', format='%Y-%m')
            # NaT 값을 가진 행 제거
            df = df.dropna(subset=['month'])

            return df

    @staticmethod
    def get_individual_storage_data(table_name):
        """
        개별 컨테이너의 스토리지 비용을 계산하는 메서드
        """
        conn = connect_db()
        query = f"""
        WITH calculated_days AS (
            SELECT * ,
                date(
                    CASE
                        WHEN [unloadingterminal] LIKE '%/%'
                        THEN substr([unloadingterminal], 7, 4) || '-' ||
                             substr([unloadingterminal], 4, 2) || '-' ||
                             substr([unloadingterminal], 1, 2)
                        ELSE [unloadingterminal]
                    END
                ) AS unloading_terminal_date,
                date([terminalappointment]) AS terminal_appointment_date,
                strftime('%Y-%m', date([terminalappointment])) AS month
            FROM {table_name}
            WHERE container IS NOT NULL
        ),
        days_over_seven AS (
            SELECT *,
                julianday(terminal_appointment_date) - julianday(unloading_terminal_date) + 1 AS total_stay_days
            FROM calculated_days
        )
        SELECT *,
            CASE
                WHEN destinationport IN ('LZO', 'LAZARO', 'Lazaro Cardenas', 'LAZARO CARDENAS') THEN
                    CASE
                        WHEN total_stay_days > 10 THEN total_stay_days - 10
                        ELSE 0
                    END
                ELSE
                    CASE
                        WHEN total_stay_days > 7 THEN total_stay_days - 7
                        ELSE 0
                    END
            END AS days_over,
            CASE
                WHEN total_stay_days > 7 THEN 
                    CASE 
                        WHEN destinationport IN ('LZO', 'LAZARO', 'Lazaro Cardenas', 'LAZARO CARDENAS') THEN
                            CASE
                                WHEN total_stay_days > 10 THEN (total_stay_days - 10) * 2027
                                ELSE 0
                            END
                        ELSE 4356.42 + (total_stay_days - 7) * 2194.43
                    END
                ELSE 0
            END AS storage_cost
        FROM days_over_seven
        ORDER BY month
        """
        df = pd.read_sql(query, conn)
        conn.close()
        return df

    @staticmethod
    def get_dd_container_data(table_name):
        """
        DD 고객의 월별 컨테이너 수를 계산하는 메서드
        """
        conn = connect_db()
        query = f"""
        WITH calculated_days AS (
            SELECT 
                strftime('%Y-%m', date([terminalappointment])) AS month,
                [f.dest],
                COUNT(*) as container_count
            FROM {table_name}
            WHERE container IS NOT NULL
            GROUP BY month, [f.dest]
        )
        SELECT 
            month,
            SUM(CASE WHEN [f.dest] = 'DD' THEN container_count ELSE 0 END) as dd_count
        FROM calculated_days
        GROUP BY month
        ORDER BY month
        """
        df = pd.read_sql(query, conn)
        conn.close()
        
        # month 컬럼을 datetime으로 변환
        df['month'] = pd.to_datetime(df['month'], format='%Y-%m')
        return df

    # StorageUtils 클래스에 새로운 메서드 추가
    @staticmethod
    def get_kpi_container_count(table_name):
        """KPI용 컨테이너 카운트 계산"""
        conn = connect_db()
        query = f"""
        SELECT 
            strftime('%Y-%m', date([terminalappointment])) AS month,
            COUNT(*) as container_count
        FROM {table_name}
        WHERE container IS NOT NULL
        GROUP BY month
        ORDER BY month
        """
        df = pd.read_sql(query, conn)
        conn.close()
        
        
        
        # 날짜 표식 처리
        df['month'] = pd.to_datetime(df['month'], errors='coerce', format='%Y-%m')
        df = df.dropna(subset=['month'])
        
        return df

class StorageCostAnalysisWindow(QDialog):
    def __init__(self, table_name, parent=None, tab_info=None):
        super().__init__(parent)
        self.table_name = table_name
        self.tab_info = tab_info  # tab_info 저장
        self.setWindowTitle("Storage Cost Analysis")
        self.setGeometry(100, 100, 1200, 800)

        self.setWindowFlags(Qt.Window | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)
        
        
        # 메인 레이아웃
        main_layout = QVBoxLayout()
        self.setLayout(main_layout)

        # 1. 월별 데이터 테이블
        monthly_group = QGroupBox("Monthly Storage Data")
        monthly_group.setFixedSize(850, 300)
        monthly_layout = QVBoxLayout()
        self.monthly_table = QTableWidget()
        self.monthly_table.setColumnCount(6)
        self.monthly_table.setHorizontalHeaderLabels([
            "Month", "Storage Cost (MXN)", "Container Count", "Cost per Container", "D.Delivery","DD Ratio (%)"
        ])
        
        # 테이블 크기 조정 설정
        self.monthly_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.monthly_table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.monthly_table.setMaximumHeight(200)  # 최대 높이 제한
        
        monthly_layout.addWidget(self.monthly_table)
        monthly_group.setLayout(monthly_layout)
        main_layout.addWidget(monthly_group)

        # 2. Storage Cost per Container Analysis 그룹박스 (오른쪽)
        analysis_group = QGroupBox("Storage Cost per Container Analysis")
        analysis_group.setFixedSize(600, 300)
        analysis_layout = QVBoxLayout()

         # 날짜 필터 추가
        date_filter_layout = QHBoxLayout()
        date_filter_layout.addWidget(QLabel("Start Date:"))
        self.start_date = QDateEdit(calendarPopup=True)
        self.start_date.setDate(QDate.currentDate().addMonths(-6))
        date_filter_layout.addWidget(self.start_date)

        date_filter_layout.addWidget(QLabel("End Date:"))
        self.end_date = QDateEdit(calendarPopup=True)
        self.end_date.setDate(QDate.currentDate())
        date_filter_layout.addWidget(self.end_date)

        # Apply 버튼 추가
        apply_button = QPushButton("Apply Filter")
        apply_button.clicked.connect(self.update_analysis)
        date_filter_layout.addWidget(apply_button)

        analysis_layout.addLayout(date_filter_layout)

        # 분석 결과를 표시할 테이블 위젯
        self.analysis_table = QTableWidget()
        self.analysis_table.setColumnCount(2)
        self.analysis_table.setRowCount(3)
        self.analysis_table.setHorizontalHeaderLabels(["Metric", "Value"])
        self.analysis_table.setVerticalHeaderLabels(["Average", "Standard Deviation", "Coefficient of Variation"])

        # 데이터 계산
        storage_data = StorageUtils.get_monthly_storage_data(self.table_name)
        if not storage_data.empty:
            # 컨테이너당 비용 계산
            storage_data['cost_per_container'] = storage_data['total_storage_cost'] / storage_data['container_count']
            
            # 통계 계산
            mean_cost = storage_data['cost_per_container'].mean()
            std_cost = storage_data['cost_per_container'].std()
            cv_cost = (std_cost / mean_cost) * 100 if mean_cost != 0 else 0

            # 결과를 테이블에 표시
            metrics = [
                (mean_cost, "MXN {:,.2f}"),
                (std_cost, "MXN {:,.2f}"),
                (cv_cost, "{:.2f}%")
            ]

            for row, (value, format_str) in enumerate(metrics):
                # Metric 열
                metric_item = QTableWidgetItem()
                metric_item.setFont(QFont("Arial", 10))
                self.analysis_table.setItem(row, 0, metric_item)
                
                # Value 열
                value_item = QTableWidgetItem(format_str.format(value))
                value_item.setFont(QFont("Arial", 10))
                self.analysis_table.setItem(row, 1, value_item)

        # 테이블 크기 조정
        self.analysis_table.resizeColumnsToContents()
        self.analysis_table.resizeRowsToContents()

        analysis_layout.addWidget(self.analysis_table)
        analysis_group.setLayout(analysis_layout)

        # 수평 레이아웃에 두 그룹박스 추가
        h_layout = QHBoxLayout()
        h_layout.addWidget(monthly_group)
        h_layout.addWidget(analysis_group)

        # 메인 레이아웃에 수평 레이아웃 추가
        main_layout.addLayout(h_layout)

        # 2. 월별 비교 분석
        comparison_group = QGroupBox("Month-to-Month Comparison")
        comparison_layout = QHBoxLayout()
        
        # 월 선택 콤보박스
        self.month1_combo = QComboBox()
        self.month2_combo = QComboBox()
        comparison_layout.addWidget(QLabel("Compare:"))
        comparison_layout.addWidget(self.month1_combo)
        comparison_layout.addWidget(QLabel("with:"))
        comparison_layout.addWidget(self.month2_combo)
        
        # 비교 버튼
        compare_btn = QPushButton("Run")
        compare_btn.clicked.connect(self.compare_months)
        comparison_layout.addWidget(compare_btn)
        
        comparison_group.setLayout(comparison_layout)
        main_layout.addWidget(comparison_group)

        # 3. 그래프를 표시할 영역
        self.figure = Figure(figsize=(10, 6))
        self.canvas = FigureCanvas(self.figure)
        main_layout.addWidget(self.canvas)

        # 데이터 로드 및 초기 표시
        self.load_data()
        
    def load_data(self):
        try:
            if self.table_name == "all_tables":
                # 모든 테이블의 데이터 합치기
                storage_data = pd.DataFrame()
                dd_data = pd.DataFrame()
                
                for table_name in self.tab_info.values():
                    # 각 테이블의 데이터 가져오기
                    table_storage = StorageUtils.get_monthly_storage_data(table_name)
                    table_dd = StorageUtils.get_dd_container_data(table_name)
                    
                    storage_data = pd.concat([storage_data, table_storage], ignore_index=True)
                    dd_data = pd.concat([dd_data, table_dd], ignore_index=True)
                
                # 월별로 데이터 집계
                storage_data = storage_data.groupby('month').agg({
                    'total_storage_cost': 'sum',
                    'container_count': 'sum'
                }).reset_index()
                
                dd_data = dd_data.groupby('month').agg({
                    'dd_count': 'sum'
                }).reset_index()
                
                # DD 비율 계산
                merged_data = pd.merge(storage_data, dd_data, on='month', how='left')
                merged_data['dd_ratio'] = (merged_data['dd_count'] / merged_data['container_count'] * 100).fillna(0)
                
                # 테이블에 데이터 표시
                self.monthly_table.setRowCount(len(merged_data))
                for i, row in merged_data.iterrows():
                    self.monthly_table.setItem(i, 0, QTableWidgetItem(row['month'].strftime('%Y-%m')))
                    self.monthly_table.setItem(i, 1, QTableWidgetItem(f"{row['total_storage_cost']:,.2f}"))
                    self.monthly_table.setItem(i, 2, QTableWidgetItem(f"{row['container_count']:,}"))
                    cost_per_container = row['total_storage_cost'] / row['container_count'] if row['container_count'] > 0 else 0
                    self.monthly_table.setItem(i, 3, QTableWidgetItem(f"{cost_per_container:,.2f}"))
                    self.monthly_table.setItem(i, 4, QTableWidgetItem(f"{row['dd_count']:,}"))
                    self.monthly_table.setItem(i, 5, QTableWidgetItem(f"{row['dd_ratio']:.1f}"))
                    
                # 콤보박스 업데이트
                months = merged_data['month'].dt.strftime('%Y-%m').tolist()
                self.month1_combo.addItems(months)
                self.month2_combo.addItems(months)
                
                # 그래프 업데이트
                self.plot_monthly_trends(merged_data)
                
            else:
                # 스토리지 데이터와 DD 데이터 가져오기
                storage_data = StorageUtils.get_monthly_storage_data(self.table_name)
                dd_data = StorageUtils.get_dd_container_data(self.table_name)
                
                if storage_data.empty:
                    QMessageBox.warning(self, "Warning", "스토리지 데이터가 없습니다.")
                    return

                # DD 데이터를 storage_data와 병합
                storage_data = pd.merge(storage_data, dd_data, on='month', how='left')
                storage_data['dd_count'] = storage_data['dd_count'].fillna(0)
                storage_data['dd_ratio'] = (storage_data['dd_count'] / storage_data['container_count'] * 100)

                # 테이블에 데이터 표시
                self.monthly_table.setRowCount(len(storage_data))
                for i, row in storage_data.iterrows():
                    cost_per_container = row['total_storage_cost'] / row['container_count'] if row['container_count'] > 0 else 0
                    
                    # 테이블에 데이터 입력
                    self.monthly_table.setItem(i, 0, QTableWidgetItem(row['month'].strftime('%Y-%m')))
                    self.monthly_table.setItem(i, 1, QTableWidgetItem(f"{row['total_storage_cost']:,.2f}"))
                    self.monthly_table.setItem(i, 2, QTableWidgetItem(str(int(row['container_count']))))
                    self.monthly_table.setItem(i, 3, QTableWidgetItem(f"{cost_per_container:,.2f}"))
                    self.monthly_table.setItem(i, 4, QTableWidgetItem(str(int(row['dd_count']))))
                    self.monthly_table.setItem(i, 5, QTableWidgetItem(f"{row['dd_ratio']:.1f}"))

                # 콤보박스에 월 목록 추가
                months = [d.strftime('%Y-%m') for d in storage_data['month']]
                self.month1_combo.clear()
                self.month2_combo.clear()
                self.month1_combo.addItems(months)
                self.month2_combo.addItems(months)

                # 초기 그래프 표시
                self.plot_monthly_trends(storage_data)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"데이터 로드 중 오류 발생: {str(e)}")

    def plot_monthly_trends(self, df):
        """월별 추세 그래프 표시"""
        self.figure.clear()
        
        # 2024년 이후 데이터만 필터링 (명시적 복사 생성)
        df = df[df['month'].dt.year >= 2024].copy()
        
        ax1 = self.figure.add_subplot(111)
        ax2 = ax1.twinx()
        ax3 = ax1.twinx()  # cost per container를 위한 새로운 y축
        
        # 스토리지 비용 선 그래프 (cyan)
        line1 = ax1.plot(df['month'], df['total_storage_cost'], 'c-o', label='Storage Cost')
        ax1.set_xlabel('Month', color='white')
        ax1.set_ylabel('Storage Cost (MXN)', color='cyan')
        
        # 컨테이너 수 막대 그래프 (yellow)
        bars = ax2.bar(df['month'], df['container_count'], width=8, alpha=0.2, color='yellow', label='Container Count')
        ax2.set_ylabel('Container Count', color='yellow')
        
        # cost per container 선 그래프 (magenta)
        df.loc[:, 'cost_per_container'] = df['total_storage_cost'] / df['container_count']
        line3 = ax3.plot(df['month'], df['cost_per_container'], 'm-o', label='Cost per Container')
        ax3.set_ylabel('Cost per Container (MXN)', color='magenta')
        
        # ax3를 오른쪽으로 조금 이동 (그래프가 겹치지 않도록)
        ax3.spines['right'].set_position(('outward', 60))
        
        # 데이터 레이블 추가
        for x, y in zip(df['month'], df['total_storage_cost']):
            ax1.annotate(f'{y:,.0f}', 
                        (x, y), 
                        textcoords="offset points", 
                        xytext=(0,10), 
                        ha='center',
                        color='white')

        # 스타일링
        self.figure.patch.set_facecolor('#19232D')
        ax1.set_facecolor('#19232D')
        ax1.tick_params(colors='white')
        ax2.tick_params(colors='white')
        ax3.tick_params(colors='white')
        
        # x축 레이블 회전
        plt.xticks(rotation=45)
        
        # 범례 추가
        lines1 = line1
        lines2 = [bars]
        lines3 = line3
        labels1 = ['Storage Cost']
        labels2 = ['Container Count']
        labels3 = ['Cost per Container']
        ax1.legend(lines1 + lines2 + lines3, labels1 + labels2 + labels3, loc='upper left')

        self.canvas.draw()

    def compare_months(self):
        """선택한 두 월의 데이터 비교"""
        month1 = self.month1_combo.currentText()
        month2 = self.month2_combo.currentText()

        try:
            # StorageUtils를 사용하여 데이터 가져오기
            storage_data = StorageUtils.get_monthly_storage_data(self.table_name)
            
            # 선택한 월에 해당하는 데이터 필터링
            df = storage_data[storage_data['month'].dt.strftime('%Y-%m').isin([month1, month2])]

            if len(df) != 2:
                QMessageBox.warning(self, "Warning", "두 개의 월을 선택해주세요.")
                return

            # 비교 결과를 새 창으로 표시
            comparison_dialog = QDialog(self)
            comparison_dialog.setWindowTitle(f"Comparison: {month1} vs {month2}")
            comparison_dialog.setGeometry(200, 200, 800, 400)

            layout = QVBoxLayout()
            
            # 비교 테이블 생성
            table = QTableWidget()
            table.setColumnCount(4)  # 지표, 월1, 월2, 변화율
            table.setRowCount(3)
            table.setHorizontalHeaderLabels(["", month1, month2, "Change %"])
            table.setVerticalHeaderLabels(["Storage Cost", "Container Count", "Cost per Container"])

            # 데이터 정렬
            df = df.sort_values('month')

            # 데이터 입력 및 변화율 계산
            for i, (metric, row1, row2) in enumerate([
                ('Storage Cost', df.iloc[0]['total_storage_cost'], df.iloc[1]['total_storage_cost']),
                ('Container Count', df.iloc[0]['container_count'], df.iloc[1]['container_count']),
                ('Cost per Container', 
                 df.iloc[0]['total_storage_cost']/df.iloc[0]['container_count'],
                 df.iloc[1]['total_storage_cost']/df.iloc[1]['container_count'])
            ]):
                # 각 월의 데이터
                table.setItem(i, 1, QTableWidgetItem(f"{row1:,.2f}"))
                table.setItem(i, 2, QTableWidgetItem(f"{row2:,.2f}"))
                
                # 변화율 계산
                if row1 != 0:
                    change_pct = ((row2 - row1) / row1) * 100
                    table.setItem(i, 3, QTableWidgetItem(f"{change_pct:+.2f}%"))
                else:
                    table.setItem(i, 3, QTableWidgetItem("N/A"))

            layout.addWidget(table)
            comparison_dialog.setLayout(layout)
            comparison_dialog.show()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"비교 분석 중 오류 발생: {str(e)}")

    def update_analysis(self):
        try:
            start_date = pd.to_datetime(self.start_date.date().toPyDate())
            end_date = pd.to_datetime(self.end_date.date().toPyDate())

            if self.table_name == "all_tables":
                # 모든 테이블의 데이터 합치기
                storage_data = pd.DataFrame()
                dd_data = pd.DataFrame()
                
                for table_name in self.tab_info.values():
                    table_storage = StorageUtils.get_monthly_storage_data(table_name)
                    table_dd = StorageUtils.get_dd_container_data(table_name)
                    
                    storage_data = pd.concat([storage_data, table_storage], ignore_index=True)
                    dd_data = pd.concat([dd_data, table_dd], ignore_index=True)
                
                # 월별로 데이터 집계
                storage_data = storage_data.groupby('month').agg({
                    'total_storage_cost': 'sum',
                    'container_count': 'sum'
                }).reset_index()
                
                dd_data = dd_data.groupby('month').agg({
                    'dd_count': 'sum'
                }).reset_index()
                
                # DD 비율 계산
                merged_data = pd.merge(storage_data, dd_data, on='month', how='left')
                merged_data['dd_ratio'] = (merged_data['dd_count'] / merged_data['container_count'] * 100).fillna(0)
                
                filtered_data = merged_data
            else:
                # 단일 테이블 데이터 가져오기
                storage_data = StorageUtils.get_monthly_storage_data(self.table_name)
                dd_data = StorageUtils.get_dd_container_data(self.table_name)
                
                # DD 데이터를 storage_data와 병합
                filtered_data = pd.merge(storage_data, dd_data, on='month', how='left')
                filtered_data['dd_count'] = filtered_data['dd_count'].fillna(0)
                filtered_data['dd_ratio'] = (filtered_data['dd_count'] / filtered_data['container_count'] * 100)

            # 날짜 필터링
            filtered_data = filtered_data[
                (filtered_data['month'] >= start_date) & 
                (filtered_data['month'] <= end_date)
            ].copy()

            if not filtered_data.empty:
                # 컨테이너당 비용 계산
                filtered_data.loc[:, 'cost_per_container'] = (
                    filtered_data['total_storage_cost'] / filtered_data['container_count']
                )
                
                # 통계 계산 및 테이블 업데이트
                mean_cost = filtered_data['cost_per_container'].mean()
                std_cost = filtered_data['cost_per_container'].std()
                cv_cost = (std_cost / mean_cost) * 100 if mean_cost != 0 else 0

                metrics = [
                    (mean_cost, "MXN {:.2f}"),
                    (std_cost, "MXN {:.2f}"),
                    (cv_cost, "{:.2f}%")
                ]

                for row, (value, format_str) in enumerate(metrics):
                    value_item = QTableWidgetItem(format_str.format(value))
                    value_item.setFont(QFont("Arial", 10))
                    self.analysis_table.setItem(row, 1, value_item)

                # 그래프 업데이트
                self.plot_monthly_trends(filtered_data)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"분석 업데이트 중 오류 발생: {str(e)}")



# SQLite 데이터베이스에 연결하는 함수
def connect_db():
    try:
        # 최대 3번 재시도
        max_retries = 3
        retry_count = 0
        
        while retry_count < max_retries:
            try:
                conn = sqlite3.connect(db_file, timeout=20)  # timeout 값 증가
                return conn
            except sqlite3.OperationalError as e:
                retry_count += 1
                if retry_count == max_retries:
                    error_msg = f"""
                    Database connection failure ({retry_count}/{max_retries})
                    
                    Possible workarounds:
                    1. Compruebe su conexión VPN o de red corporativa
                    2. Check shared folder access permissions
                    3. Contact your network administrator
                    
                    Error messages: {str(e)}
                    """
                    QMessageBox.critical(None, "Database connection errors", error_msg)
                    raise
                print(f"Connection retrying... ({retry_count}/{max_retries})")
                time.sleep(2)  # 재시도 전 2초 대기
                
    except Exception as e:
        error_msg = f"""
        An unexpected error occurred during the database connection.
        
        Error messages: {str(e)}
        """
        QMessageBox.critical(None, "Error", error_msg)
        raise

class CustomFilterProxyModel(QSortFilterProxyModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.column_filters = {}

    def set_column_filter(self, column, pattern):
        if pattern:
            self.column_filters[column] = re.compile(pattern, re.IGNORECASE)
        else:
            self.column_filters.pop(column, None)
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row, source_parent):
        # 컬럼 필터 검사
        for col, pattern in self.column_filters.items():
            idx = self.sourceModel().index(source_row, col, source_parent)
            text = self.sourceModel().data(idx, Qt.DisplayRole)
            if not pattern.search(str(text)):
                return False
        # 기본 필터 검사 (검색창 필터)
        return super().filterAcceptsRow(source_row, source_parent)

class DataDisplayWindow(QDialog):
    def __init__(self, df, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Data View")
        self.setGeometry(100, 100, 1200, 800)
        
        # 윈도우 플래그 설정 - 기본 윈도우 컨트롤 활성화
        self.setWindowFlags(Qt.Window | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)
        self.df = df  
        
        # 메인 레이아웃
        main_layout = QVBoxLayout()
        
        # 상단 컨트롤 레이아웃
        control_layout = QHBoxLayout()
        
        # 검색 입력 필드
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search")
        self.search_input.textChanged.connect(self.apply_filter)
        control_layout.addWidget(self.search_input)
        
        # 필터 콤보박스
        self.filter_combo = QComboBox()
        self.filter_combo.addItems(["ALL"] + df.columns.tolist())
        self.filter_combo.currentIndexChanged.connect(self.apply_filter)
        control_layout.addWidget(self.filter_combo)
        
        # 칼럼 선택 버튼
        select_columns_btn = QPushButton("Select Columns")
        select_columns_btn.clicked.connect(self.show_column_selector)
        control_layout.addWidget(select_columns_btn)
        
        # 엑셀 내보내기 버튼
        export_excel_btn = QPushButton("Export to Excel")
        export_excel_btn.clicked.connect(self.export_to_excel)
        control_layout.addWidget(export_excel_btn)
        
        main_layout.addLayout(control_layout)
        
        # 테이블 설정
        self.table_view = QTableView()
        self.update_table_model(df)
        
        # 다중 선택 활성화
        self.table_view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        
        # 복사 단축키 설정
        copy_shortcut = QShortcut(QKeySequence.Copy, self.table_view)
        copy_shortcut.activated.connect(self.copy_selection)
        
        # 컨텍스트 메뉴 활성화
        self.table_view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table_view.customContextMenuRequested.connect(self.show_context_menu)
        
        main_layout.addWidget(self.table_view)
        
        self.setLayout(main_layout)
        
        # 헤더 컨텍스트 메뉴 연결
        header = self.table_view.horizontalHeader()
        header.setContextMenuPolicy(Qt.CustomContextMenu)
        header.customContextMenuRequested.connect(self.show_header_context_menu)

    def show_column_selector(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Select Columns")
        dialog.setGeometry(100, 100, 250, 800)  
        layout = QVBoxLayout()
        
        # 체크박스 리스트 위젯
        list_widget = QListWidget()
        list_widget.setSelectionMode(QAbstractItemView.MultiSelection)
        
        # 현재 표시된 칼럼
        current_columns = [self.model.headerData(i, Qt.Horizontal) for i in range(self.model.columnCount())]
        
        # 모든 가능한 칼럼들에 대해 체크박스 생성
        for column in self.df.columns:
            item = QListWidgetItem(column)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            # 현재 표시된 칼럼이면 Checked로, 아니면 Unchecked로 설정
            if column in current_columns:
                item.setCheckState(Qt.Checked)
            else:
                item.setCheckState(Qt.Unchecked)
            list_widget.addItem(item)
        
        # 전체 선택/해제 버튼 추가
        button_layout = QHBoxLayout()
        select_all_btn = QPushButton("Select All")
        deselect_all_btn = QPushButton("Deselect All")
        
        def select_all():
            for i in range(list_widget.count()):
                list_widget.item(i).setCheckState(Qt.Checked)
                
        def deselect_all():
            for i in range(list_widget.count()):
                list_widget.item(i).setCheckState(Qt.Unchecked)
                
        select_all_btn.clicked.connect(select_all)
        deselect_all_btn.clicked.connect(deselect_all)
        
        button_layout.addWidget(select_all_btn)
        button_layout.addWidget(deselect_all_btn)
        layout.addLayout(button_layout)
        
        layout.addWidget(list_widget)
        
        # 버튼 박스
        button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        )
        button_box.accepted.connect(lambda: self.apply_column_selection(list_widget, dialog))
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)
        
        dialog.setLayout(layout)
        dialog.exec_()
    
    def apply_column_selection(self, list_widget, dialog):
        # 선택된 칼럼들 가져오기
        selected_columns = []
        for i in range(list_widget.count()):
            item = list_widget.item(i)
            if item.checkState() == Qt.Checked:
                selected_columns.append(item.text())
        
        if selected_columns:  # 최소한 하나의 칼럼이 선택되었는지 확인
            # 선택된 칼럼들로 새로운 데이터프레임 생성
            filtered_df = self.df[selected_columns]
            self.update_table_model(filtered_df)
            dialog.accept()
        else:
            QMessageBox.warning(self, "Warning", "You must select at least one column.")
    
    def update_table_model(self, df):
        """테이블 모델 업데이트"""
        self.model = QStandardItemModel()
        self.model.setHorizontalHeaderLabels(df.columns)
        
        for row in df.itertuples(index=False):
            items = [QStandardItem(str(cell)) for cell in row]
            # 셀 선택 및 복사 가능하도록 설정
            for item in items:
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.model.appendRow(items)
        
        self.proxy_model = CustomFilterProxyModel()  # 커스텀 프록시 모델 사용
        self.proxy_model.setSourceModel(self.model)
        self.proxy_model.setFilterCaseSensitivity(Qt.CaseInsensitive)
        self.table_view.setModel(self.proxy_model)
        self.table_view.resizeColumnsToContents()
    
    def apply_filter(self):
        filter_text = self.search_input.text()
        selected_column = self.filter_combo.currentText()
        
        if selected_column == "All":
            self.proxy_model.setFilterKeyColumn(-1)
        else:
            column_index = next(
                (idx for idx in range(self.model.columnCount())
                 if self.model.headerData(idx, Qt.Horizontal) == selected_column),
                -1
            )
            self.proxy_model.setFilterKeyColumn(column_index)
        
        self.proxy_model.setFilterFixedString(filter_text)
        
    def export_to_excel(self):
        """
        현재 테이블에 표시된 데이터를 엑셀 파일로 내보내는 메서드
        """
        try:
            # 현재 필터링된 데이터 가져오기
            filtered_data = []
            for row in range(self.proxy_model.rowCount()):
                row_data = []
                for col in range(self.proxy_model.columnCount()):
                    index = self.proxy_model.index(row, col)
                    row_data.append(self.proxy_model.data(index))
                filtered_data.append(row_data)
            
            # 데이터프레임 생성
            columns = [self.model.headerData(i, Qt.Horizontal) for i in range(self.model.columnCount())]
            filtered_df = pd.DataFrame(filtered_data, columns=columns)
            
            # 파일 저장 다이얼로그
            file_path, _ = QFileDialog.getSaveFileName(
                self, 
                "Save as Excel File", 
                "", 
                "Excel Files (*.xlsx)"
            )
            
            if file_path:
                if not file_path.endswith('.xlsx'):
                    file_path += '.xlsx'
                
                # 엑셀 파일로 저장
                filtered_df.to_excel(file_path, index=False)
                QMessageBox.information(self, "Success", f"File saved successfully:\n{file_path}")
        
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while saving the Excel file:\n{str(e)}")

    def copy_selection(self):
        """선택된 셀들을 클립보드로 복사"""
        selection = self.table_view.selectedIndexes()
        if not selection:
            return
            
        # 선택된 행과 열의 범위 찾기
        rows = sorted(index.row() for index in selection)
        columns = sorted(index.column() for index in selection)
        rowcount = rows[-1] - rows[0] + 1
        colcount = columns[-1] - columns[0] + 1
        
        # 복사할 데이터를 담을 표
        table = [[''] * colcount for _ in range(rowcount)]
        
        # 선택된 셀들의 데이터를 표에 채우기
        for index in selection:
            row = index.row() - rows[0]
            column = index.column() - columns[0]
            table[row][column] = index.data()
        
        # 표를 탭으로 구분된 문자열로 변환
        stream = io.StringIO()
        csv.writer(stream, delimiter='\t').writerows(table)
        QApplication.clipboard().setText(stream.getvalue())

    def show_context_menu(self, position):
        """우클릭 메뉴 표시"""
        menu = QMenu()
        copy_action = menu.addAction("Copy")
        copy_action.triggered.connect(self.copy_selection)
        
        # 전체 선택 액션 추가
        select_all_action = menu.addAction("Select All")
        select_all_action.triggered.connect(self.table_view.selectAll)
        
        menu.exec_(self.table_view.viewport().mapToGlobal(position))

    def show_header_context_menu(self, pos):
        header = self.table_view.horizontalHeader()
        col = header.logicalIndexAt(pos)
        
        menu = QMenu(self)
        filter_action = menu.addAction("Filter...")
        clear_filter_action = menu.addAction("Clear Filter")
        
        action = menu.exec_(self.table_view.mapToGlobal(pos))
        
        if action == filter_action:
            self.show_filter_dialog(col)
        elif action == clear_filter_action:
            self.proxy_model.set_column_filter(col, None)

    def show_filter_dialog(self, column):
        # 현재 보이는 컬럼 이름 조회
        col_name = self.model.headerData(column, Qt.Horizontal)
        
        # 현재 표시된 데이터프레임에서 컬럼 인덱스 찾기 (수정된 부분)
        filtered_df = self.df[[self.model.headerData(i, Qt.Horizontal) for i in range(self.model.columnCount())]]
        current_col_idx = filtered_df.columns.get_loc(col_name)
        
        # 고유값 추출 대상 변경 (원본 -> 필터링된 DF)
        unique_values = filtered_df.iloc[:, current_col_idx].astype(str).unique().tolist()
        
        # 기존 필터 패턴 가져오기 (현재 컬럼 인덱스 기준)
        current_filter = self.proxy_model.column_filters.get(current_col_idx, None)
        selected_values = set()
        if current_filter:
            pattern = current_filter.pattern
            if pattern.startswith('^(') and pattern.endswith(')$'):
                values = pattern[2:-2].split('|')
                selected_values = {v.replace('\\', '') for v in values}

        dialog = QDialog(self)
        dialog.setWindowTitle(f"Filter: {col_name}")
        layout = QVBoxLayout()
        
        list_widget = QListWidget()
        for value in sorted(unique_values, key=lambda x: x.lower()):
            item = QListWidgetItem(value)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if value in selected_values else Qt.Unchecked)
            list_widget.addItem(item)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        
        layout.addWidget(list_widget)
        layout.addWidget(button_box)
        dialog.setLayout(layout)
        
        if dialog.exec_() == QDialog.Accepted:
            selected = []
            for i in range(list_widget.count()):
                item = list_widget.item(i)
                if item.checkState() == Qt.Checked:
                    selected.append(re.escape(item.text()))
            
            if selected:
                pattern = "^(" + "|".join(selected) + ")$"
                # 현재 컬럼 인덱스로 필터 적용
                self.proxy_model.set_column_filter(current_col_idx, pattern)
            else:
                self.proxy_model.set_column_filter(current_col_idx, None)

class DateRangeDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Date Range")
        self.setWindowFlags(Qt.Window | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)

        # 레이아웃 설정
        layout = QVBoxLayout()

        # 시작 날짜 선택
        start_date_label = QLabel("Start Date:")
        self.start_date_edit = QDateEdit(calendarPopup=True)
        self.start_date_edit.setDate(QDate.currentDate())
        self.start_date_edit.setCalendarPopup(True)
        layout.addWidget(start_date_label)
        layout.addWidget(self.start_date_edit)

        # 종료 날짜 선택
        end_date_label = QLabel("End Date:")
        self.end_date_edit = QDateEdit(calendarPopup=True)
        self.end_date_edit.setDate(QDate.currentDate())
        self.end_date_edit.setCalendarPopup(True)
        layout.addWidget(end_date_label)
        layout.addWidget(self.end_date_edit)

        # 확인 및 취소 버튼
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

        self.setLayout(layout)



# 이메일 전송 검토용 창 클래스
class EmailReviewDialog(QDialog):
    def __init__(self, recipient, cc, subject, body, attachment, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Email Review")
        self.setGeometry(100, 100, 600, 400)
        self.setWindowFlags(Qt.Window | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)


        # 레이아웃 설정
        layout = QVBoxLayout()

        # 수신자 이메일 (원본 이메일 주소 사용)
        recipient_label = QLabel(f"Recipient: {recipient}")
        layout.addWidget(recipient_label)

        # 참조 이메일을 가로 4개씩 표시하는 위한 그리드 레이아웃 설정
        cc_label = QLabel("CC:")
        layout.addWidget(cc_label)

        cc_grid_layout = QGridLayout()
        cc_emails = cc.split('; ')  # 참조 이메일을 리스트로 변환 (세미콜론 구분)

        # 참조 이메일을 4개씩 가로로 배치
        for i, email in enumerate(cc_emails):
            cc_grid_layout.addWidget(QLabel(email), i // 4, i % 4)

        layout.addLayout(cc_grid_layout)  # 그리드 레이아웃을 메인 레이아웃에 추가

        # 이메일 제목 입력 필드
        subject_label = QLabel("Subject:")
        layout.addWidget(subject_label)

        self.subject_input = QLineEdit()  # 사용자가 제목을 입력할 수 있는 필드
        self.subject_input.setText(subject)  # 기본 제목을 설정
        layout.addWidget(self.subject_input)

        # 이메일 본문 스크롤 영역 설정
        body_label = QLabel("Body:")
        layout.addWidget(body_label)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)

        body_text = QTextEdit()
        body_text.setPlainText(body)  # 이메일 본문을 그대로 표시
        body_text.setReadOnly(False)  
        scroll_area.setWidget(body_text)
        layout.addWidget(scroll_area)  # 스크롤 영역을 레이아웃에 추가

        # 첨부 파일 정보
        attachment_label = QLabel(f"Attachment: {attachment}")
        layout.addWidget(attachment_label)

        # 확인 및 메일 전송 버튼
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)  
        button_box.rejected.connect(self.reject)  
        layout.addWidget(button_box)

        self.setLayout(layout)

        # 창의 크기를 조정하여 스크롤 가능한 내에 맞도록 설정
        self.resize(1500, 900)

    def get_subject(self):
        """
        사용자가 입력한 제목을 반환하는 메서드
        """
        return self.subject_input.text()

class ShippingLineDataDisplayWindow(QDialog):
    def __init__(self, df, shipping_line, start_date=None, end_date=None, parent=None):
        super().__init__(parent)
        # shipping_line은 이미 매핑된 이름으로 전달됨
        self.mapped_shipping_line = shipping_line
        self.setWindowTitle(f"{shipping_line} Data View")
        self.setGeometry(100, 100, 1200, 800)
        self.df = df
        self.start_date = start_date
        self.end_date = end_date
        self.setWindowFlags(Qt.Window | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)

        # 기본 메시지 설정 - 매핑된 이름 사용
        self.default_email_body = f"""
        Dear {self.mapped_shipping_line} Team,

        We hope this message finds you well.

        We are sending you the expected delivery plan for next week. Please note that the detailed schedule may be subject to change.

        If your company has any changes or requests, we would appreciate it if you could contact the relevant person in charge.


        To.  LX PANTOS FF Team,

        We would appreciate it if you could promptly share any schedule changes with the respective selected shipping line.


        Thank you for your cooperation.

        ----------------------------------------------------------------------------------------------------------------------------------------

        Estimado/a {self.mapped_shipping_line},

        Esperamos que este mensaje le encuentre bien.

        Le enviamos el plan de entrega previsto para la próxima semana. Tenga en cuenta que el calendario detallado puede estar sujeto a cambios.

        Si su empresa tiene algún cambio o solicitud, le agradeceríamos que se pusiera en contacto con la persona responsable correspondiente.


        Para el Equipo de LX PANTOS FF:

        Agradeceríamos que compartieran de manera rápida cualquier cambio en el cronograma con la línea naviera seleccionada correspondiente.


        Gracias por su cooperación.



        Best regards,
        LX PANTOS FF TEAM
        LX PANTOS Mexico City Branch
        Nave 10, Km 42.5 Autopista Mex-Qro, Parque Industrial Cedros,
        Tepotzotlan, Estado de Mexico. C.P. 54602, Mexico
        E. PANTOSUS_403660@lxpantos.com
        """

        # 레이아웃 설정
        layout = QVBoxLayout()

        # 테이블 설정 및 데이터 표시
        self.table_view = QTableView()
        self.model = self.create_model(df)  # create_model 메서드 호출
        self.table_view.setModel(self.model)
        layout.addWidget(self.table_view)

        # 버튼 레이아웃
        button_layout = QHBoxLayout()
        export_btn = QPushButton("Export to Excel File")
        export_btn.clicked.connect(self.export_to_excel)
        email_btn = QPushButton("Send by Email")  # 추가
        email_btn.clicked.connect(self.send_email)  # 추가
        button_layout.addStretch()
        button_layout.addWidget(export_btn)
        button_layout.addWidget(email_btn)  # 추가

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def create_model(self, df):
        """
        데이터프레임을 QStandardItemModel로 변환하는 메서드
        """
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(df.columns)
        for row in df.itertuples(index=False):
            items = [QStandardItem(str(cell)) for cell in row]
            model.appendRow(items)
        return model

    def export_to_excel(self):
        """
        데이터를 엑셀 파일로 저장하는 메서드
        """
        try:
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Save as Excel File", "", "Excel Files (*.xlsx *.xls)"
            )
            if file_path:
                self.df.to_excel(file_path, index=False)
                QMessageBox.information(self, "Success", f"File saved successfully: {file_path}")
            else:
                QMessageBox.warning(self, "Cancel", "File saving was canceled.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while saving the Excel file: {str(e)}")

    def is_valid_email(self, email):
       
        pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
        return re.match(pattern, email.strip()) is not None

    def send_email(self):
        # 데이터베이스에서 원본 shipping line 이름들 가져오기
        conn = connect_db()
        cursor = conn.cursor()
        
        original_shipping_lines = []
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'table%'")
        tables = cursor.fetchall()
        
        for (table_name,) in tables:
            cursor.execute(f"SELECT DISTINCT shippingline FROM [{table_name}]")
            shipping_lines = cursor.fetchall()
            original_shipping_lines.extend([sl[0] for sl in shipping_lines if sl[0]])
        
        # 매칭되는 원본 shipping line 이름들 찾기
        matching_shipping_lines = [
            sl for sl in set(original_shipping_lines)
            if Mapping.standardize_shipping_line_name(sl) == self.mapped_shipping_line
        ]
        
        if not matching_shipping_lines:
            matching_shipping_lines = [self.mapped_shipping_line]
        
        # 이메일 주소 가져오기 - 원본 이름들로 검색
        to_emails = []
        cc_emails = []
        for original_sl in matching_shipping_lines:
            sl_to, sl_cc = self.get_shipping_line_emails(original_sl)
            to_emails.extend(sl_to)
            cc_emails.extend(sl_cc)
        
        # 중복 제거
        to_emails = list(set(to_emails))
        cc_emails = list(set(cc_emails))
        
        # 이메일 검토 창 표시
        date_range_str = ""
        if self.start_date and self.end_date:
            date_range_str = f"{self.start_date.strftime('%Y%m%d')}_{self.end_date.strftime('%Y%m%d')}"
        
        email_subject = f"{self.mapped_shipping_line} Weekly Delivery Plan_{date_range_str}"
        
        review_dialog = EmailReviewDialog(
            recipient='; '.join(to_emails),
            cc='; '.join(cc_emails),
            subject=email_subject,
            body=self.default_email_body,
            attachment="(Will be created when sending)"
        )
        
        if review_dialog.exec_() == QDialog.Accepted:
            try:
                # 임시 파일 생성
                temp_file_path = f"{self.mapped_shipping_line}_weekly_plan_{date_range_str}.xlsx"
                self.df.to_excel(temp_file_path, index=False)
                
                # 메일 전송
                self.send_email_via_outlook(
                    '; '.join(to_emails),
                    '; '.join(cc_emails),
                    email_subject,
                    self.default_email_body,
                    temp_file_path
                )
                QMessageBox.information(self, "Success", "Email sent successfully.")
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"An error occurred while sending the email: {str(e)}")
            finally:
                try:
                    os.remove(temp_file_path)
                except Exception as e:
                    print(f"Failed to delete temporary file: {str(e)}")

    def send_email_via_outlook(self, recipient_emails, cc_emails, subject, body, attachment_path):
        
        try:
            outlook = win32com.client.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)  # 0: MailItem

            mail.To = recipient_emails
            if cc_emails.strip():  
                mail.CC = cc_emails
            mail.Subject = subject

            # 이메일 본문을 UTF-8로 인코딩하여 설정
            mail.Body = body.encode('utf-8').decode('utf-8')

            # 파일 첨부
            attachment_full_path = os.path.abspath(attachment_path)
            mail.Attachments.Add(attachment_full_path)

            # 메일 보내기
            mail.Send()

        except Exception as e:
            raise e  # 상위에서 예외를 처리하도록 던짐

    def get_shipping_line_emails(self, shipping_line):
        """
        Get all 'To' and 'CC' email addresses for a given Shipping Line name from the Vessel contact.db file.
        """
        #contact_db_path = r"C:\Users\SyncthingServiceAcct\test\aiden\Vessel contact.db"
        try:
            conn = sqlite3.connect(contact_db_path)
            cursor = conn.cursor()

            # 'VESSEL CONTACT' 테이블에서 모든 행을 조회
            query = "SELECT [HAPAG LLOYD], [HYUNDAI], [MAERSK], [MSC], [ONE], [CC] FROM [VESSEL CONTACT]"
            cursor.execute(query)
            results = cursor.fetchall()
            conn.close()

            to_emails = []
            cc_emails = []

            # 각 행을 순회하며 해당 Shipping Line의 이메일과 CC 이메일을 수집
            for row in results:
                hapag_email, hyundai_email, maersk_email, msc_email, one_email, cc_email = row
                shipping_line_key = shipping_line.upper()
                shipping_line_email = None

                if shipping_line_key == "HAPAG LLOYD" and hapag_email:
                    shipping_line_email = hapag_email
                elif shipping_line_key == "HYUNDAI" and hyundai_email:
                    shipping_line_email = hyundai_email
                elif shipping_line_key == "MAERSK" and maersk_email:
                    shipping_line_email = maersk_email
                elif shipping_line_key == "MSC" and msc_email:
                    shipping_line_email = msc_email
                elif shipping_line_key == "ONE" and one_email:
                    shipping_line_email = one_email

                if shipping_line_email:
                    # 쉼표 또는 세미콜론으로 구분된 이메일 주소 분리
                    emails = re.split(r'[;,]', shipping_line_email)
                    for email in emails:
                        email = email.strip()
                        if email:
                            to_emails.append(email)

                # CC 이메일 수집
                if cc_email:
                    cc_split = re.split(r'[;,]', cc_email)
                    for email in cc_split:
                        email = email.strip()
                        if email:
                            cc_emails.append(email)

            # 중복 제거
            to_emails = list(set(to_emails))
            cc_emails = list(set(cc_emails))

            return to_emails, cc_emails
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while retrieving emails from the contact database: {str(e)}")
            return [], []

class EmailInputDialog(QDialog):
    def __init__(self, parent=None, default_body="", shipping_line=None):
        super().__init__(parent)
        self.setWindowTitle("Email Information Input")
        self.setGeometry(100, 100, 400, 350)  # 창 크기 약간 확장
        self.setWindowFlags(Qt.Window | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)

        # shipping_line이 제공된 경우 매핑된 이름 사용
        if shipping_line:
            self.mapped_shipping_line = Mapping.standardize_shipping_line_name(shipping_line)
        else:
            self.mapped_shipping_line = None

        layout = QVBoxLayout()

        # 수신자 이메일 입력
        recipient_label = QLabel("Recipient Email:")
        self.recipient_email_input = QLineEdit()
        layout.addWidget(recipient_label)
        layout.addWidget(self.recipient_email_input)

        # 참조(CC) 이메일 입력
        cc_label = QLabel("CC Email:")
        self.cc_email_input = QLineEdit()
        self.cc_email_input.setPlaceholderText("Enter the email addresses to be referenced (comma separated)")
        layout.addWidget(cc_label)
        layout.addWidget(self.cc_email_input)

        # 이메일 제목 입력
        subject_label = QLabel("Email Subject:")
        self.subject_input = QLineEdit()
        layout.addWidget(subject_label)
        layout.addWidget(self.subject_input)

        # 이메일 본문 입력
        body_label = QLabel("Email Body:")
        self.body_input = QTextEdit()
        self.body_input.setText(default_body)  # 기본 메시지 설정
        layout.addWidget(body_label)
        layout.addWidget(self.body_input)

        # 확인 및 취소 버튼
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

        self.setLayout(layout)



class Mapping:
    @staticmethod
    def standardize_shipping_line_name(name):
        
        if not name:
            return ''

        # 표준화 매핑 딕셔너리
        mapping = {
            'HAPAG LLOYD': 'HAPAG LLOYD',
            'HAPAG': 'HAPAG LLOYD',
            'ONE': 'ONE',
            'ONE ': 'ONE',
            'CMA ' : 'CMA',
            'EVERGREEN ' : 'EVERGREEN'
            # 필요한 경우 다른 Shipping Line에 대한 매핑 추가
        }
        # 매핑된 이름 반환
        return mapping.get(name, name)
    
    def standardize_origin_name(name):
        if not name:
            return ''

        # Origin 이름 매핑 딕셔너리
        mapping = {
            'COREA': 'KRPUS',
            'COREA ' : 'KRPUS',
            'KRSHG' : 'KRPUS',
            'SHANGHAI' : 'CNSHA',
            'VIETNAM' : 'VNHPH'

            # 필요한 경우 다른 Origin에 대한 매핑 추가
        }

        # 매핑된 이름 반환
        return mapping.get(name, name)



class OriginWindow(QDialog):
    def __init__(self, parent=None, current_table=""):
        super().__init__(parent)
        self.setWindowTitle("Origin Analysis")
        self.setGeometry(200, 200, 800, 600)
        self.parent = parent  # MainWindow 참조
        self.current_table = current_table  # 현재 테이블 이름 저장

        self.setWindowFlags(Qt.Window | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)
        

        # 전체 레이아웃 설정
        main_layout = QHBoxLayout()
        self.setLayout(main_layout)

        # 왼쪽 영역: Origin 리스트
        left_layout = QVBoxLayout()
        left_layout.addWidget(QLabel("Origins"))

        # Origin 리스트 위젯 설정
        self.origin_list_widget = QListWidget()
        self.origin_list_widget.setSelectionMode(QAbstractItemView.SingleSelection)
        self.origin_list_widget.itemSelectionChanged.connect(self.update_shipping_lines)

        # Origin 리스트 채우기 - 매핑 적용 및 빈 값 처리
        origins = self.parent.get_unique_origins()
        standardized_origins = []
        
        for origin in origins:
            if origin and not pd.isna(origin):  # 빈 값과 NA 체크
                standardized_origin = Mapping.standardize_origin_name(origin)
                if standardized_origin:  # 빈 문자열 체크
                    standardized_origins.append(standardized_origin)
        
        # 중복 제거 후 정렬하여 리스트에 추가
        unique_origins = sorted(set(standardized_origins))
        if unique_origins:  # 최종 리스트가 비어있지 않은 경우만 추가
            self.origin_list_widget.addItems(unique_origins)

        left_layout.addWidget(self.origin_list_widget)
        main_layout.addLayout(left_layout)

        # 오른쪽 영역: 날짜 선택, Shipping Line, 분석 결과
        right_layout = QVBoxLayout()

        # 날짜 범위 선택 위젯 추가
        date_layout = QHBoxLayout()
        date_layout.addWidget(QLabel("Start Date (etaport):"))
        self.start_date_edit = QDateEdit(calendarPopup=True)
        self.start_date_edit.setDate(QDate.currentDate().addMonths(-1))  # 기본값: 한 달 전
        date_layout.addWidget(self.start_date_edit)

        date_layout.addWidget(QLabel("End Date (etaport):"))
        self.end_date_edit = QDateEdit(calendarPopup=True)
        self.end_date_edit.setDate(QDate.currentDate())  # 기본값: 오늘
        date_layout.addWidget(self.end_date_edit)

        right_layout.addLayout(date_layout)

        right_layout.addWidget(QLabel("Shipping Lines"))
        # Shipping Line 리스트 위젯 설정
        self.shipping_line_list_widget = QListWidget()
        self.shipping_line_list_widget.setSelectionMode(QAbstractItemView.SingleSelection)
        self.shipping_line_list_widget.itemSelectionChanged.connect(self.perform_analysis)
        right_layout.addWidget(self.shipping_line_list_widget)

        # Destination Port 리스트 추가
        right_layout.addWidget(QLabel("Destination Ports"))
        # Destination Port 리스트 위젯 설정
        self.destination_port_list_widget = QListWidget()
        self.destination_port_list_widget.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.destination_port_list_widget.itemSelectionChanged.connect(self.perform_analysis)
        right_layout.addWidget(self.destination_port_list_widget)
        
        # Destination Ports 목록 채우기
        self.update_destination_ports()

        right_layout.addWidget(QLabel("Analysis Results"))
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        right_layout.addWidget(self.result_text)

        # 날짜 선택 위젯에 이벤트 연결
        self.start_date_edit.dateChanged.connect(self.perform_analysis)
        self.end_date_edit.dateChanged.connect(self.perform_analysis)

        # 버튼 레이아웃
        button_layout = QHBoxLayout()
        
        # Verify Data 버튼 추가
        verify_button = QPushButton("Verify Data")
        verify_button.clicked.connect(self.verify_data)
        button_layout.addWidget(verify_button)
        
        # Close 버튼
        close_button = QPushButton("Close")
        close_button.clicked.connect(self.accept)
        button_layout.addWidget(close_button)
        
        right_layout.addLayout(button_layout)
        main_layout.addLayout(right_layout)

    def update_destination_ports(self):
        """Destination Ports 목록 업데이트"""
        try:
            conn = connect_db()
            cursor = conn.cursor()
            
            # 모든 테이블에서 unique한 destinationport 값들을 가져옴
            all_ports = set()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'table%'")
            tables = cursor.fetchall()
            
            for (table_name,) in tables:
                cursor.execute(f"SELECT DISTINCT destinationport FROM [{table_name}]")
                ports = cursor.fetchall()
                all_ports.update([port[0] for port in ports if port[0]])
            
            conn.close()
            
            # 정렬하여 리스트에 추가
            sorted_ports = sorted(all_ports)
            self.destination_port_list_widget.clear()
            self.destination_port_list_widget.addItems(sorted_ports)
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load destination ports: {str(e)}")

    def verify_data(self):
        try:
            # 선택된 Origin 확인
            selected_origin_items = self.origin_list_widget.selectedItems()
            if not selected_origin_items:
                QMessageBox.warning(self, "Warning", "Please select an Origin.")
                return
            mapped_origin = selected_origin_items[0].text()
            
            # 선택된 Shipping Line 확인
            selected_shipping_items = self.shipping_line_list_widget.selectedItems()
            if not selected_shipping_items:
                QMessageBox.warning(self, "Warning", "Please select a Shipping Line.")
                return
            mapped_shipping_line = selected_shipping_items[0].text()

            # 선택된 Destination Port 확인 - 다중 선택 지원
            selected_port_items = self.destination_port_list_widget.selectedItems()
            if not selected_port_items:
                QMessageBox.warning(self, "Warning", "Please select at least one Destination Port.")
                return
            selected_ports = [item.text() for item in selected_port_items]
            
            # 날짜 범위 가져오기
            start_date = self.start_date_edit.date().toString("yyyy-MM-dd")
            end_date = self.end_date_edit.date().toString("yyyy-MM-dd")
            
            # 데이터베이스에서 모든 원본 값들을 가져오기
            conn = connect_db()
            cursor = conn.cursor()
            
            # 모든 테이블에서 unique한 origin과 shipping line 값들을 가져옴
            original_origins = []
            original_shipping_lines = []
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'table%'")
            tables = cursor.fetchall()
            
            for (table_name,) in tables:
                # Origin 값들 가져오기
                cursor.execute(f"SELECT DISTINCT origin FROM [{table_name}]")
                origins = cursor.fetchall()
                original_origins.extend([origin[0] for origin in origins if origin[0]])
                
                # Shipping Line 값들 가져오기
                cursor.execute(f"SELECT DISTINCT shippingline FROM [{table_name}]")
                shipping_lines = cursor.fetchall()
                original_shipping_lines.extend([sl[0] for sl in shipping_lines if sl[0]])
            
            # 중복 제거
            original_origins = list(set(original_origins))
            original_shipping_lines = list(set(original_shipping_lines))
            
            # 매칭되는 원본 값들 찾기
            matching_origins = [
                orig for orig in original_origins 
                if Mapping.standardize_origin_name(orig) == mapped_origin
            ]
            
            matching_shipping_lines = [
                sl for sl in original_shipping_lines 
                if Mapping.standardize_shipping_line_name(sl) == mapped_shipping_line
            ]
            
            if not matching_origins:
                matching_origins = [mapped_origin]
            if not matching_shipping_lines:
                matching_shipping_lines = [mapped_shipping_line]
            
            # 데이터 검색
            all_data = []
            for (table_name,) in tables:
                origin_placeholders = ','.join(['?' for _ in matching_origins])
                shipping_placeholders = ','.join(['?' for _ in matching_shipping_lines])
                port_placeholders = ','.join(['?' for _ in selected_ports])
                query = f"""
                SELECT *, '{table_name}' as source_table
                FROM [{table_name}]
                WHERE [origin] IN ({origin_placeholders})
                AND [shippingline] IN ({shipping_placeholders})
                AND [destinationport] IN ({port_placeholders})
                AND date([etaport]) BETWEEN ? AND ?
                """
                try:
                    params = matching_origins + matching_shipping_lines + selected_ports + [start_date, end_date]
                    df = pd.read_sql(query, conn, params=params)
                    # 빈 데이터프레임이 아닐 때만 추가
                    if not df.empty:
                        # NA 값이 있는 컬럼 처리
                        df = df.dropna(how='all', axis=1)  # 모든 값이 NA인 컬럼 제거
                        all_data.append(df)
                except Exception as e:
                    print(f"Error querying {table_name}: {str(e)}")
                    continue

            conn.close()

            if all_data:
                # 빈 데이터프레임이 없는 상태에서만 concat 실행
                if len(all_data) > 0:
                    combined_df = pd.concat(all_data, ignore_index=True)
                    combined_df = combined_df.sort_values('etaport')

                    # 데이터 표시 창 생성
                    window = DataDisplayWindow(combined_df, self)
                    window.setWindowTitle(f"Data Verification - {mapped_origin} / {mapped_shipping_line}")
                    window.setModal(False)
                    window.show()
            else:
                QMessageBox.information(self, "Information", "No data available for the selected criteria.")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while verifying data: {str(e)}")

    def update_shipping_lines(self):
        selected_items = self.origin_list_widget.selectedItems()
        if not selected_items:
            self.shipping_line_list_widget.clear()
            self.result_text.clear()
            return
            
        # 선택된 origin 표준화
        origin = selected_items[0].text()
        standardized_origin = Mapping.standardize_origin_name(origin)
        
        # 표준화된 origin으로 shipping lines 가져오기
        shipping_lines = self.parent.get_shipping_lines_for_origin(standardized_origin)
        
        self.shipping_line_list_widget.clear()
        # shipping lines도 표준화가 필요합니다
        standardized_shipping_lines = [Mapping.standardize_shipping_line_name(sl) for sl in shipping_lines]
        unique_shipping_lines = sorted(set(standardized_shipping_lines))  # 중복 제거 후 정렬
        self.shipping_line_list_widget.addItems(unique_shipping_lines)
        
        # shipping line이 업데이트되면 분석도 다시 수행
        self.perform_analysis()

    def perform_analysis(self):
        try:
            # 모든 필수 선택 항목이 선택되었는지 확인
            if (not self.origin_list_widget.selectedItems() or
                not self.shipping_line_list_widget.selectedItems() or
                not self.destination_port_list_widget.selectedItems()):
                return  # 필수 항목이 선택되지 않았으면 조용히 리턴
            
            # origin 표준화 적용
            origin = self.origin_list_widget.selectedItems()[0].text()
            standardized_origin = Mapping.standardize_origin_name(origin)

            # shipping line 표준화 적용
            shipping_line = self.shipping_line_list_widget.selectedItems()[0].text()
            standardized_shipping_line = Mapping.standardize_shipping_line_name(shipping_line)

            # destination ports 가져오기
            selected_ports = [item.text() for item in self.destination_port_list_widget.selectedItems()]

            # 날짜 범위 가져오기
            start_date = self.start_date_edit.date().toString("yyyy-MM-dd")
            end_date = self.end_date_edit.date().toString("yyyy-MM-dd")

            # 분석 수행
            analysis_results = self.parent.calculate_analyses(
                standardized_origin, 
                [standardized_shipping_line],
                selected_ports,
                start_date, 
                end_date
            )

            # 결과 표시
            if isinstance(analysis_results, dict):
                result_str = f"Analysis Results for:\n"
                result_str += f"Origin: {standardized_origin}\n"
                result_str += f"Shipping Line: {standardized_shipping_line}\n"
                result_str += f"Destination Ports: {', '.join(selected_ports)}\n"
                result_str += f"Date Range: {start_date} ~ {end_date}\n\n"
                
                for key, value in analysis_results.items():
                    if isinstance(value, (int, float)):
                        result_str += f"{key}: {value:.2f}\n"
                    else:
                        result_str += f"{key}: {value}\n"
                self.result_text.setText(result_str)
            elif isinstance(analysis_results, str):
                self.result_text.setText(analysis_results)
            else:
                self.result_text.setText("An unknown error occurred.")
                
        except Exception as e:
            self.result_text.setText(f"An error occurred during analysis: {str(e)}")
            print(f"Error in perform_analysis: {str(e)}")  # 디버깅을 위한 출력


# 메인 윈도우 클래스
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CY Aging Visualization")
        self.setGeometry(100, 100, 1200, 800)

        self.current_data = {}

        self.donut_click_cid = {}

        self.division_colors = {}  # {(year, division): color} 형태로 저장

        # 메뉴바 생성
        self.create_menu_bar()

        # Tab 위젯 생성
        self.tabs = QTabWidget()

        # 탭 이름과 테이블 이름 매핑
        self.tab_info = {
            "WM": "table1",
            "AC": "table2",
            "AV_AO": "table3",
            "REF": "table4",
            "MWO": "table5",
            "DW": "table6",
            "MN": "table7",
            "TV": "table8",
            "RB": "table9",
            "JEM": "table10",
            "FCL": "table11"
        }

        # 각각의 Tab 생성
        self.tab_widgets = {}
        for tab_name, table_name in self.tab_info.items():
            tab = self.create_tab(tab_name, table_name)
            self.tabs.addTab(tab["widget"], tab_name)
            self.tab_widgets[table_name] = tab

        # 빈 탭 생성 및 그래프를 전체 화면으로 배치
        empty_tab = QWidget()
        empty_layout = QVBoxLayout()  # QGridLayout 대신 QVBoxLayout 사용

        # 그래프 생성 및 전체 화면으로 배치
        self.combined_storage_figure = Figure()
        self.combined_storage_canvas = FigureCanvas(self.combined_storage_figure)
        empty_layout.addWidget(self.combined_storage_canvas)  # 그래프를 레이아웃에 추가

        # 레이아웃 여백 제거
        empty_layout.setContentsMargins(0, 0, 0, 0)
        empty_layout.setSpacing(0)

        empty_tab.setLayout(empty_layout)
        self.tabs.addTab(empty_tab, "Total Storage Cost")

        # 메인 레이아웃에 Tab 위젯 설정
        central_widget = QWidget()
        central_layout = QVBoxLayout()
        central_widget.setLayout(central_layout)

        # 탭 위젯 추가
        central_layout.addWidget(self.tabs)

        self.setCentralWidget(central_widget)

        # 새로운 Billing Storage 탭 생성
        self.create_billing_storage_tab()

        # Origin 메뉴를 메뉴바에 추가
        self.add_origin_menu()

        # 메인 레이아웃에 Tab 위젯 설정
        central_widget = QWidget()
        central_layout = QVBoxLayout()
        central_widget.setLayout(central_layout)

        # 탭 위젯 추가
        central_layout.addWidget(self.tabs)

        self.setCentralWidget(central_widget)

        # 프로그램 시작 시 각 탭에 차트 표시
        for table_name, tab in self.tab_widgets.items():
            self.show_storage_cost_chart(table_name, tab["storage_figure"], tab["storage_canvas"])
            self.show_chart(table_name, tab["analysis_figure"], tab["analysis_canvas"])

            # 도넛 차트 초기화 시, update_dual_donut_chart 사용
            if tab["combo_box"] and tab["combo_box"].count() > 0:
                self.update_dual_donut_chart(table_name, tab["combo_box"], tab["truck_figure"], tab["truck_canvas"],
                                             tab["rail_figure"], tab["rail_canvas"])

        # 전체 스토리지 비용 차트 표시
        self.show_combined_storage_cost_chart()

        # MainWindow 클래스의 __init__ 메서드에 추가
        kpi_menu = self.menuBar().addMenu('KPI')
        show_kpi_action = QAction('Show KPI Dashboard', self)
        show_kpi_action.triggered.connect(self.show_kpi_dashboard)
        kpi_menu.addAction(show_kpi_action)

    def show_kpi_dashboard(self):
        """KPI 대시보드 표시"""
        password_dialog = KPIPasswordDialog(self)
        if password_dialog.exec_() == QDialog.Accepted:
            # KPI 창의 크기를 더 크게 설정하여 모든 정보를 표시
            kpi_window = KPIWindow(self, self.tab_info)
            kpi_window.setGeometry(100, 100, 1600, 900)  # 창 크기 증가
            
            # 창 제목에 평가 기준 버전 표시
            kpi_window.setWindowTitle("Division KPI Dashboard - 2024 Ver.")
            
            # 모달리스로 설정하여 메인 창과 동시에 조작 가능
            kpi_window.setWindowModality(Qt.NonModal)
            
            # 항상 최상위에 표시
            kpi_window.setWindowFlags(
                kpi_window.windowFlags() | 
                Qt.WindowStaysOnTopHint
            )
            
            kpi_window.show()

    def add_origin_menu(self):
        origin_action = QAction("Origin", self)
        origin_action.triggered.connect(self.open_origin_window)
        self.menuBar().addAction(origin_action)

    def open_origin_window(self):
        try:
            # 현재 선택된 테이블 이름을 가져옵다.
            current_table = self.get_current_table_name()
            
            self.origin_window = OriginWindow(self, current_table)
            self.origin_window.show()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while opening the Origin window: {str(e)}")

    def get_current_table_name(self):
        # 현재 선택된 탭의 테이블 이름을 반환하는 메서드
        current_index = self.tabs.currentIndex()
        tab_name = self.tabs.tabText(current_index)
        return self.tab_info.get(tab_name, "")

    def create_billing_storage_tab(self):
        # 탭 위젯 생성
        billing_tab = QWidget()
        layout = QHBoxLayout()

        # 왼쪽에 디비전 리스트 뷰 생성 (크기 축소)
        list_container = QWidget()
        list_layout = QVBoxLayout()
        list_container.setLayout(list_layout)
        list_container.setFixedWidth(120)  # 리스트 뷰 너비 축소
        
        self.division_list_view = QListWidget()
        self.division_list_view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.division_list_view.setMaximumWidth(120)  # 리스트 최대 너비 제한
        
        # Billing_storage 테이블에서 디비전 목록 가져오기
        conn = sqlite3.connect(db_file)
        query = "SELECT DISTINCT [divlgems] FROM Billing_storage"
        df_divisions = pd.read_sql(query, conn)
        conn.close()

        divisions = df_divisions['divlgems'].dropna().tolist()
        divisions.sort()

        # 'Total' 항목을 리스트의 첫 번째에 추가
        self.division_list_view.addItem('total')
        self.division_list_view.addItems(divisions)

        # 리스트 뷰에서 드래그 기능 활성화
        self.division_list_view.setDragEnabled(True)
        
        list_layout.addWidget(self.division_list_view)
        layout.addWidget(list_container)

        # 그래프를 표시할 위젯 생성 (크기 확장)
        self.billing_figure = Figure(figsize=(12, 6))  # 그래프 크기 증가
        self.billing_canvas = FigureCanvas(self.billing_figure)
        
        # 그래프 위젯 드롭 대상으로 설정
        self.billing_canvas.setAcceptDrops(True)
        self.billing_canvas.dragEnterEvent = self.canvas_drag_enter_event
        self.billing_canvas.dropEvent = self.canvas_drop_event
        
        # 그래프에 더 많은 공간 할당
        layout.addWidget(self.billing_canvas, stretch=4)

        billing_tab.setLayout(layout)
        self.tabs.addTab(billing_tab, "Billing Storage")

        # 처음에 월별 총 청구 스토리지 비용을 표시
        self.show_billing_storage_chart()

    def show_billing_storage_chart(self, divisions=None):
        try:
            # divisions가 None이거나 비어있는 경우 안내 메시지 표시
            if not divisions:
                self.billing_figure.clear()
                ax = self.billing_figure.add_subplot(111)
                
                # 배경색 설정
                self.billing_figure.patch.set_facecolor('#19232D')
                ax.set_facecolor('#19232D')
                
                # 축 숨기기
                ax.set_xticks([])
                ax.set_yticks([])
                
                # 안내 메시지
                instruction_text = """
                How to use Billing Storage Chart:
                
                1. Select desired Division from the left list
                2. Drag and drop selected items to the graph area
                3. Multiple Divisions can be selected simultaneously
                4. Adjust data period using year checkboxes
                5. Right-click to show/hide graphs
                
                To begin, select and drag a Division from the left.

                -------------------------------------------------------------------------

                Cómo usar la Gráfica de Almacenamiento de Facturación:
                
                1. Seleccione la División deseada de la lista izquierda
                2. Arrastre y suelte los elementos seleccionados en el área de la gráfica
                3. Se pueden seleccionar múltiples Divisiones simultáneamente
                4. Ajuste el período de datos usando las casillas de años
                5. Haga clic derecho para mostrar/ocultar gráficas
                
                Para comenzar, seleccione y arrastre una División desde la izquierda.


                """
                
                # 텍스트 추가
                ax.text(0.45, 0.5, instruction_text,
                       ha='center', va='center',
                       color='white',
                       fontsize=12,
                       transform=ax.transAxes
                       )
                
                self.billing_canvas.draw()
                return
                
            conn = sqlite3.connect(db_file)
            df = pd.read_sql("SELECT * FROM Billing_storage", conn)
            conn.close()

            if df.empty:
                QMessageBox.warning(self, "Warning", "No data available")
                return

            # 기존 드래그앤드롭 관련 설정 유지
            self.billing_canvas.setAcceptDrops(True)
            self.billing_canvas.dragEnterEvent = self.canvas_drag_enter_event
            self.billing_canvas.dropEvent = self.canvas_drop_event

            if not df.empty:
                # podeta를 datetime으로 변환
                df['podeta'] = pd.to_datetime(df['podeta'])
                df['year'] = df['podeta'].dt.year
                df['month'] = df['podeta'].dt.month  # 직접 월 추출
                
                # 연도 체크박스 위젯 생성 (없는 경우에만)
                if not hasattr(self, 'year_checkboxes'):
                    checkbox_widget = QWidget()
                    checkbox_layout = QHBoxLayout()
                    checkbox_layout.setAlignment(Qt.AlignLeft)
                    
                    self.year_checkboxes = {}
                    years = sorted(df['year'].unique())
                    
                    for year in years:
                        checkbox = QCheckBox(str(year))
                        checkbox.setChecked(True)  # 기본적으로 모든 연도 선택
                        checkbox.stateChanged.connect(
                            lambda: self.show_billing_storage_chart(divisions)
                        )
                        self.year_checkboxes[year] = checkbox
                        checkbox_layout.addWidget(checkbox)
                    
                    checkbox_widget.setLayout(checkbox_layout)
                    # 체크박스 위젯을 billing_tab의 레이아웃 상단에 추가
                    self.tabs.widget(self.tabs.count() - 1).layout().insertWidget(0, checkbox_widget)

                # 선택된 연도에 따라 데이터 필터링
                selected_years = [year for year, checkbox in self.year_checkboxes.items() 
                                if checkbox.isChecked()]
                df = df[df['year'].isin(selected_years)]

                # 날짜 형식 변환
                df['podeta'] = pd.to_datetime(df['podeta'], errors='coerce')

                # 유효한 podeta 값만 사용
                df = df.dropna(subset=['podeta'])

                # 'Total' 컬럼에서 불필요한 문자 제거 및 숫자 형식으로 변환
                df['total'] = df['total'].replace({',': '', r'\$': '', ' ': ''}, regex=True)
                df['total'] = pd.to_numeric(df['total'], errors='coerce')
                df = df.dropna(subset=['total'])

                # 월로 그룹화
                df['month'] = df['podeta'].dt.month

                self.billing_figure.clear()
                ax1 = self.billing_figure.add_subplot(111)
                ax2 = ax1.twinx()

                # x축 설정 - 월만 표시
                months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                         'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
                ax1.set_xticks(range(12))
                ax1.set_xticklabels(months)

                # 그래프의 가시성 상태를 저장하는 딕셔너리 초기화
                if not hasattr(self, 'billing_visibility'):
                    self.billing_visibility = {}

                self.billing_lines = []
                self.billing_bars = []

                # 여기서부터 수정 시작
                # 기존 색상 설정 위젯 제거
                tab_layout = self.tabs.widget(self.tabs.count() - 1).layout()
                for i in range(tab_layout.count()):
                    widget = tab_layout.itemAt(i).widget()
                    if isinstance(widget, QWidget) and widget.findChild(QGroupBox):
                        widget.setParent(None)
                        break

                    
                # 색상 설정 버튼을 포함할 컨테이너 위젯
                color_settings_widget = QWidget()
                color_layout = QVBoxLayout()
                
                # 연도별 색상 설정 그룹
                for year in selected_years:
                    year_group = QGroupBox(f"Colors for {year}")
                    year_layout = QVBoxLayout()
                    
                    for division in divisions:
                        color_key = (year, division)
                        if color_key not in self.division_colors:
                            # 초기 색상 랜덤 설정
                            self.division_colors[color_key] = QColor(*np.random.randint(0, 255, 3))
                        
                        # 색상 선택 버튼 생성
                        color_button = QPushButton(f"{division}")
                        color_button.setStyleSheet(
                            f"background-color: {self.division_colors[color_key].name()};"
                            f"color: {'black' if self.division_colors[color_key].lightness() > 128 else 'white'};"
                        )
                        
                        # 색상 선택 다이얼로그 연결
                        def create_color_picker(y, d):
                            def pick_color():
                                color = QColorDialog.getColor(
                                    initial=self.division_colors[(y, d)],
                                    parent=self,
                                    title=f"Select color for {d} ({y})"
                                )
                                if color.isValid():
                                    self.division_colors[(y, d)] = color
                                    sender = self.sender()
                                    sender.setStyleSheet(
                                        f"background-color: {color.name()};"
                                        f"color: {'black' if color.lightness() > 128 else 'white'};"
                                    )
                                    self.show_billing_storage_chart(divisions)
                            return pick_color
                        
                        color_button.clicked.connect(create_color_picker(year, division))
                        year_layout.addWidget(color_button)
                    
                    year_group.setLayout(year_layout)
                    color_layout.addWidget(year_group)
                
                color_settings_widget.setLayout(color_layout)
                
                # 색상 설정 위젯을 billing_tab의 왼쪽에 추가
                tab_layout = self.tabs.widget(self.tabs.count() - 1).layout()
                if tab_layout.indexOf(color_settings_widget) == -1:
                    tab_layout.insertWidget(0, color_settings_widget)

                if divisions:
                    bar_width = 0.3
                    month_spacing = 1
                    
                    for year_idx, year in enumerate(selected_years):
                        year_offset = (year - min(selected_years)) * 0.3
                        bottom_values = np.zeros(12)
                        
                        for division in divisions:
                            if division == 'total':
                                df_filtered = df[df['year'] == year]
                            else:
                                df_filtered = df[(df['divlgems'] == division) & (df['year'] == year)]
                            
                            if not df_filtered.empty:
                                df_grouped = df_filtered.groupby('month').agg({
                                    'total': 'sum',
                                    'cntrno.': 'count'
                                }).reset_index()
                                
                                monthly_volumes = np.zeros(12)
                                monthly_costs = np.zeros(12)
                                month_indices = df_grouped['month'] - 1 
                                monthly_volumes[month_indices] = df_grouped['cntrno.']
                                monthly_costs[month_indices] = df_grouped['total']
                                
                                # 사용자가 선택한 색상 사용
                                color = self.division_colors[(year, division)]
                                color_str = color.name()
                                
                                # 막대 그래프
                                bars = ax2.bar(
                                    np.arange(12) + year_offset,
                                    monthly_volumes,
                                    bar_width,
                                    bottom=bottom_values,
                                    label=f'{division} {year}',
                                    color=color_str,
                                    alpha=0.3,
                                    zorder=1
                                )
                                
                                bottom_values += monthly_volumes
                                
                                self.billing_bars.append({
                                    'bars': bars, 
                                    'label': f'{division} {year} Volume', 
                                    'year': year
                                })
                                self.billing_visibility[f'{division} {year} Volume'] = True
                                
                                # 선 그래프
                                line, = ax1.plot(
                                    np.arange(12),
                                    monthly_costs,
                                    marker='o',
                                    label=f'{division} Cost {year}',
                                    color=color_str,
                                    linewidth=2,
                                    markersize=6,
                                    zorder=5
                                )
                                
                                self.billing_lines.append({
                                    'line': line, 
                                    'label': f'{division} Cost {year}', 
                                    'year': year
                                })
                                self.billing_visibility[f'{division} Cost {year}'] = True

                    # x축 범위 및 레이블 설정
                    total_years = len(selected_years)
                    total_width = (total_years - 1) * 0.1 + 1 # 변경된 간격 값 반영
                    ax1.set_xlim(-0.5, (11 * month_spacing) + total_width)
                    
                    # x축 레이블을 월별로만 표시
                    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                             'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
                    
                    # x축 눈금 위치 설정 (년도 구분 없이 1월부터 12월까지)
                    ax1.set_xticks(np.arange(0, 12 * month_spacing, month_spacing)) # 월별 간격을 1로 설정
                    ax1.set_xticklabels(months, rotation=45, ha='center')

                # 그래프 시각적 설정
                ax1.set_xlabel('Month', color='white', fontsize=12)
                ax1.set_ylabel('Total Amount (Million MXN)', color='white', fontsize=12)  # 단위 표시 추가
                ax1.set_title('Monthly Billing Storage Cost and Volume (Drag and Drop)', color='white', fontsize=14)

                def millions_formatter(x, pos):
                    return f'{x/1000000:.1f}'

                ax1.yaxis.set_major_formatter(FuncFormatter(millions_formatter))

                ax1.tick_params(axis='x', colors='white', rotation=45)
                ax1.tick_params(axis='y', colors='white')
                ax2.set_ylabel('Volume (Count of cntrno.)', color='white', fontsize=12)
                ax2.tick_params(axis='y', colors='white')

                # 배경색 설정
                self.billing_figure.patch.set_facecolor('#19232D')
                ax1.set_facecolor('#19232D')
                ax2.set_facecolor('#19232D')

                # 틱 라벨 색상 설정
                for label in ax1.get_xticklabels():
                    label.set_color('white')
                for label in ax1.get_yticklabels():
                    label.set_color('white')
                for label in ax2.get_yticklabels():
                    label.set_color('white')

                
                
                # 마우스 오른쪽 클릭 이벤트 핸들러 정의
                def on_canvas_click(event):
                    if event.button == 3:  # 오른쪽 클릭
                        menu = QMenu(self)
                        visibility = self.billing_visibility

                        # 스토리지 비용 그래프 메뉴
                        for item in self.billing_lines:
                            label = item['label']
                            action = QAction(label, menu)
                            action.setCheckable(True)
                            action.setChecked(visibility.get(label, True))
                            
                            def create_toggle_function(item, label):
                                def toggle(checked):
                                    visibility[label] = checked
                                    item['line'].set_visible(checked)
                                    self.billing_canvas.draw()
                                    menu.close()  # 메뉴 닫기 추가
                                return toggle

                            action.triggered.connect(create_toggle_function(item, label))
                            menu.addAction(action)

                        # 컨테이너 수 그래프 메뉴
                        for item in self.billing_bars:
                            label = item['label']
                            action = QAction(label, menu)
                            action.setCheckable(True)
                            action.setChecked(visibility.get(label, True))
                            
                            def create_toggle_function(item, label):
                                def toggle(checked):
                                    visibility[label] = checked
                                    for bar in item['bars']:
                                        bar.set_visible(checked)
                                    self.billing_canvas.draw()
                                    menu.close()  # 메뉴 닫기 추가
                                return toggle

                            action.triggered.connect(create_toggle_function(item, label))
                            menu.addAction(action)

                        # 메뉴 표시
                        menu.popup(QCursor.pos())  # exec_ 대신 popup 사용

                # 캔버스에 이벤트 핸들러 연결
                self.billing_canvas.mpl_connect('button_press_event', on_canvas_click)

                # 기존의 마우스 오른쪽 클릭 이벤트 핸들러 유지
                self.billing_canvas.mpl_connect('button_press_event', on_canvas_click)

                self.billing_canvas.draw()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while creating the billing storage chart: {str(e)}")

    def canvas_drag_enter_event(self, event):
        if event.mimeData().hasFormat('application/x-qabstractitemmodeldatalist'):
            event.accept()
        else:
            event.ignore()

    def canvas_drop_event(self, event):
        selected_items = self.division_list_view.selectedItems()
        divisions = [item.text() for item in selected_items]
        self.show_billing_storage_chart(divisions)

    def create_menu_bar(self):
        """
        상단 메뉴바를 생성하고 메뉴 항목들 추가합니다.
        """
        menu_bar = self.menuBar()

        # 'File' 메뉴 생성
        file_menu = menu_bar.addMenu("File")

        # '마스터 엑셀 추출' 액션 생성
        export_master_action = QAction("Master file", self)
        export_master_action.setShortcut("Ctrl+M")  # 단축키 설정 (선택 사항)
        export_master_action.triggered.connect(self.export_master_excel)

        # '엑셀 파일 업로드' 액션 생성
        upload_action = QAction("Excel file upload", self)
        upload_action.setShortcut("Ctrl+U")  # 단축키 설정 (선택 사항)
        upload_action.triggered.connect(self.upload_excel_file)

        # '선사 공유 파일' 액션 생성
        shipping_line_report_action = QAction("Weekly plan for Vessel", self)
        shipping_line_report_action.setShortcut("Ctrl+S")  # 단축키 설정 (선택 사항)
        shipping_line_report_action.triggered.connect(self.export_shipping_line_report)

        # **'데이터 갱신' 액션 생성**
        reload_action = QAction("Reload", self)
        reload_action.setShortcut("Ctrl+R")  # 단축키 설정 (선택 사항)
        reload_action.triggered.connect(self.reload_data)

        # 'File' 메뉴에 액션 추가
        file_menu.addAction(upload_action)
        file_menu.addAction(shipping_line_report_action)
        file_menu.addAction(export_master_action)
        file_menu.addAction(reload_action)  # 추가된 부분

    def export_master_excel(self):
        """
        데이터베이스의 table1부터 table11까지의 데이터를 하나의 마스터 엑셀 파일로 추출하는 기능.
        """
        try:
            conn = connect_db()
            table_names = list(self.tab_info.values())  # ['table1', 'table2', ..., 'table11']
            df_list = []
            columns_list = []

            # table1의 컬럼 순서를 가져오기 위해 먼저 table1을 처리
            table1 = table_names[0]  # assuming 'table1' is the first in the list
            query_table1 = f"SELECT * FROM {table1}"
            df_table1 = pd.read_sql(query_table1, conn)
            df_list.append(df_table1)
            columns_list.append(set(df_table1.columns.tolist()))

            # 나머지 테이블 처리
            for table in table_names[1:]:
                query = f"SELECT * FROM {table}"
                df = pd.read_sql(query, conn)
                df_list.append(df)
                columns_list.append(set(df.columns.tolist()))

            # table1의 컬럼 순서 유지
            table1_columns_order = df_table1.columns.tolist()

            # 모든 테이블의 공통 컬럼 찾기 (모든 테이블에 공통적으로 존재하는 컬럼)
            common_columns = set.intersection(*columns_list) if columns_list else set()

            if not common_columns:
                QMessageBox.warning(self, "Warning", "No common columns found.")
                conn.close()
                return

            # 공통 컬럼을 table1의 순서대로 정리
            ordered_common_columns = [col for col in table1_columns_order if col in common_columns]

            # 모든 테이블의 컬럼을 공통 컬럼 순서에 따라 재정렬하고 고유 컬럼은 뒤에 배치
            for i, df in enumerate(df_list):
                current_columns = df.columns.tolist()
                # 공통 컬럼만 먼저 선택
                ordered_common = [col for col in ordered_common_columns if col in current_columns]
                # 나머지 컬럼
                remaining = [col for col in current_columns if col not in ordered_common_columns]
                # 새로운 컬럼 ��서
                new_order = ordered_common + remaining
                # 컬럼 ��정렬 및 모든 값이 NA인 컬럼 제거
                df_ordered = df[new_order].dropna(axis=1, how='all')  # 모든 값이 NA인 컬럼 제거
                # 재정렬된 데이터프���임을 리스트에 다시 저장
                df_list[i] = df_ordered

            # 마스터 데이터프레임 생성
            master_df = pd.DataFrame()

            for df in df_list:
                master_df = pd.concat([master_df, df], ignore_index=True, sort=False)

            # 공통 컬럼을 table1의 순서대로 배치하고, 고유 컬럼은 �� 뒤에 배치
            # table1_columns_order에서 중복 제거 및 현재 master_df에 존재하는지 확인
            table1_columns_order = [col for col in table1_columns_order if col in master_df.columns]
            # 고유 컬럼 찾기
            unique_columns = [col for col in master_df.columns if col not in table1_columns_order]
            # 최종 컬럼 순서
            final_column_order = table1_columns_order + unique_columns
            # 컬럼 순서 재정렬
            master_df = master_df[final_column_order]

            conn.close()

            if master_df.empty:
                QMessageBox.information(self, "Information", "No data to include in the master Excel.")
                return

            # 저장할 엑셀 파일 경로 선택
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Save as Master Excel File", "", "Excel Files (*.xlsx *.xls)"
            )

            if file_path:
                master_df.to_excel(file_path, index=False)
                QMessageBox.information(self, "Success", f"Master Excel file saved successfully: {file_path}")
            else:
                QMessageBox.warning(self, "Cancel", "Master Excel file saving was canceled.")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while creating the Master Excel file: {str(e)}")

    def standardize_shipping_line_name(self, name):
        name = name.strip().upper()
        if 'HAPAG' in name:
            return 'HAPAG LLOYD'
        elif 'CMA' in name:
            return 'CMA CGM'
        # 추가적인 표준화 규칙을 여기에 포함
        return name

    def export_shipping_line_report(self):
        try:
            conn = connect_db()
            shipping_lines_set = set()
            shipping_line_mapping = {}
            for table_name in self.tab_info.values():
                query = f"SELECT DISTINCT [shippingline] FROM {table_name}"
                df = pd.read_sql(query, conn)
                for original_name in df['shippingline'].dropna().tolist():
                    standardized_name = self.standardize_shipping_line_name(original_name)
                    shipping_lines_set.add(standardized_name)
                    shipping_line_mapping[standardized_name] = shipping_line_mapping.get(standardized_name, []) + [
                        original_name]
            conn.close()

            if not shipping_lines_set:
                QMessageBox.information(self, "Information", "No shipping line data in the database.")
                return

            # shippingline 선택 다이얼로그 표시
            shipping_line, ok = QInputDialog.getItem(
                self,
                "Shipping Line Selection",
                "Select a Shipping Line:",
                sorted(shipping_lines_set),
                0,
                False
            )
            if ok and shipping_line:
                # 선택한 SHIPPING LINE에 대한 데이터 추출 및 엑셀 파일 저장
                self.export_data_for_shipping_line(shipping_line, shipping_line_mapping[shipping_line])
            else:
                QMessageBox.information(self, "Cancel", "Shipping line report file creation was canceled.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while creating the shipping line report file: {str(e)}")

    def export_data_for_shipping_line(self, shipping_line, original_names):
        try:
            conn = connect_db()
            selected_columns = [
                'division',
                'portofloading',
                'destinationport',
                'terminal',
                'vessel',
                'mbl',
                'container',
                'modality',
                'urgentcargo',
                'etaport',
                'unloadingterminal',
                'terminalappointment',
                'f.dest',
                'eta',
                'shippingline'  # shippingline 컬럼 포함
            ]
            df_list = []
            cursor = conn.cursor()

            for table_name in self.tab_info.values():
                # 테이블의 컬럼 정보 가��오기
                cursor.execute(f"PRAGMA table_info({table_name})")
                columns_info = cursor.fetchall()
                table_columns = [info[1] for info in columns_info]

                if 'shippingline' not in table_columns:
                    # shippingline 컬럼이 없는 테이블은 제외
                    continue

                # shippingline 컬럼이 있는 경우: 선사에 해당하는 ���이터 필터링
                placeholders = ','.join(['?'] * len(original_names))
                query = f"""
                    SELECT * FROM {table_name}
                    WHERE UPPER(TRIM([shippingline])) IN ({placeholders})
                """
                cleaned_original_names = [name.strip().upper() for name in original_names]
                df = pd.read_sql(query, conn, params=cleaned_original_names)

                if not df.empty:
                    # 필요한 컬럼만 선택하고, 없는 컬럼은 NaN으로 채움
                    for col in selected_columns:
                        if col not in df.columns:
                            df[col] = np.nan  # 또는 공란으로 채우려면 ''
                    df = df[selected_columns]
                    df_list.append(df)

            conn.close()

            if not df_list:
                QMessageBox.information(self, "Information", f"No data for shipping line '{shipping_line}'.")
                return

            # 날짜 범위 선택 다이얼로그 표시
            date_dialog = DateRangeDialog(self)
            if date_dialog.exec_() == QDialog.Accepted:
                start_date = date_dialog.start_date_edit.date().toPyDate()
                end_date = date_dialog.end_date_edit.date().toPyDate()

                # 시작 날짜가 종료 날짜보다 이후인 경우 검증
                if start_date > end_date:
                    QMessageBox.warning(self, "Error", "Start date must be earlier than end date.")
                    return
            else:
                QMessageBox.information(self, "Cancel", "File creation was canceled.")
                return

            # 합쳐진 데이터프레임 생성
            result_df = pd.concat(df_list, ignore_index=True)

            # eta 컬럼을 datetime으로 변환
            result_df['eta'] = pd.to_datetime(result_df['eta'], errors='coerce').dt.date

            # 날짜 범위 필터링 (ETA가 존재하는 경우에만 필터링)
            result_df = result_df[(result_df['eta'] >= start_date) & (result_df['eta'] <= end_date)]

            if result_df.empty:
                QMessageBox.information(self, "Information", f"No data for shipping line '{shipping_line}' within the specified date range.")
                return

            # 데이터를 보여주는 창을 띄움
            self.shipping_line_data_window = ShippingLineDataDisplayWindow(result_df, shipping_line, start_date, end_date, self)
            self.shipping_line_data_window.setModal(False)
            self.shipping_line_data_window.show()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while creating the shipping line report file: {str(e)}")

    def get_last_update_time(self, table_name):
        try:
            conn = connect_db()
            cursor = conn.cursor()

            # Log 테이블에서 해당 테이블의 마지막 업데이트 시간 가져오기
            cursor.execute("SELECT {} FROM Log WHERE field1 = 'updated_at'".format(table_name))
            result = cursor.fetchone()

            conn.close()

            # 결과가 없으면 '없음'을 반환
            if result and result[0]:
                return result[0]
            else:
                return "None"
        except Exception as e:
            return f"Error: {str(e)}"

    def create_tab(self, tab_name, table_name):
        # 탭 안에 들어갈 위젯과 레이아웃
        tab_widget = QWidget()
        layout = QVBoxLayout()

        # 우측 상단에 최근 업데이트 시간 표시
        update_label = QLabel(f"Last updated: {self.get_last_update_time(table_name)}")
        update_label.setAlignment(Qt.AlignRight)
        layout.addWidget(update_label)

        # 2x2 그리드 레이아웃 생성
        grid_layout = QGridLayout()

        # 각 섹션이 동일한 공간을 차지하도록 설정
        grid_layout.setColumnStretch(0, 1)
        grid_layout.setColumnStretch(1, 1)
        grid_layout.setRowStretch(0, 1)
        grid_layout.setRowStretch(1, 1)

        # 스토리지 비용 차트 생성 (2사분면)
        storage_figure = Figure()
        storage_canvas = FigureCanvas(storage_figure)
        grid_layout.addWidget(storage_canvas, 0, 0)  # 행 0, 열 0 (좌측 상단)

        # 컨테이너 상태 차트 생성 (1사분면)
        analysis_figure = Figure()
        analysis_canvas = FigureCanvas(analysis_figure)
        grid_layout.addWidget(analysis_canvas, 0, 1)  # 행 0, 열 1 (우측 상단)

        # 콤보박스 및 도넛 차트 생성 (3사분면)
        donut_layout = QVBoxLayout()
        combo_box = QComboBox()

        # 'destinationport' 콤보박스 추가
        conn = connect_db()
        query = f"""
            SELECT DISTINCT [destinationport] FROM {table_name}
            WHERE [destinationport] IS NOT NULL
        """
        df_ports = pd.read_sql(query, conn)
        conn.close()
        ports = sorted(df_ports['destinationport'])
        combo_box.addItems(ports)
        donut_layout.addWidget(combo_box)

        # 두 가지 modality를 위한 도넛 차트 생성
        truck_figure = Figure(figsize=(4, 4))
        truck_canvas = FigureCanvas(truck_figure)
        rail_figure = Figure(figsize=(4, 4))
        rail_canvas = FigureCanvas(rail_figure)
        donut_hbox = QHBoxLayout()
        donut_hbox.addWidget(truck_canvas)
        donut_hbox.addWidget(rail_canvas)
        donut_layout.addLayout(donut_hbox)

        # 콤보박스 변경 시 도넛 차트 업데이트
        combo_box.currentTextChanged.connect(
            lambda: self.update_dual_donut_chart(table_name, combo_box, truck_figure, truck_canvas, rail_figure,
                                                 rail_canvas)
        )

        # 도넛 차트를 3사분면에 배치
        grid_layout.addLayout(donut_layout, 1, 0)

        # **[수 ���분 시작]**

        # 4사분면에 월별 도넛 차트 및 콤보박스 추가
        delay_donut_layout = QVBoxLayout()

        # 월 선택 콤보박스 생성
        conn = connect_db()
        query_months = f"""
                SELECT DISTINCT strftime('%Y-%m', date([etaport])) AS month
                FROM {table_name}
                WHERE [initialeta] IS NOT NULL AND [etaport] IS NOT NULL AND date([etaport]) != date([initialeta])
                ORDER BY month
            """
        df_months = pd.read_sql(query_months, conn)
        conn.close()
        months = df_months['month'].dropna().tolist()

        month_combo_box = QComboBox()
        month_combo_box.addItems(months)

        
        delay_donut_layout.addWidget(month_combo_box)

        # 도넛 차트 생성
        delay_donut_figure = Figure(figsize=(4, 4))
        delay_donut_canvas = FigureCanvas(delay_donut_figure)
        delay_donut_layout.addWidget(delay_donut_canvas)

        # 월 변경 시 도넛 차트 업데이트
        month_combo_box.currentTextChanged.connect(
            lambda: self.show_vessel_delay_donut_chart(table_name, month_combo_box.currentText(), delay_donut_figure,
                                                       delay_donut_canvas)
        )

        # 초기 도넛 차트 표시
        if months:
            self.show_vessel_delay_donut_chart(table_name, months[0], delay_donut_figure, delay_donut_canvas)

        # 4사분면에 도넛 차트 배치
        grid_layout.addLayout(delay_donut_layout, 1, 1)

        # **[수정 부분 끝]**

        # 차트 그리드 레이아웃을 메인 레이아웃에 추가
        layout.addLayout(grid_layout)

        # 버튼 생성 및 기능 연결 (수정된 부분)
        master_file_btn = QPushButton("MASTER")
        master_file_btn.clicked.connect(lambda: self.open_master_file_window(table_name))

        storage_cost_btn = QPushButton("STORAGE COST")
        storage_cost_btn.clicked.connect(lambda: self.show_estimated_storage_cost(table_name))

        # 그리드 레이아웃 생성 및 버튼 배치 (1x2로 수정)
        button_layout = QGridLayout()
        button_layout.addWidget(master_file_btn, 0, 0)
        button_layout.addWidget(storage_cost_btn, 0, 1)

        # 버튼 레이아웃을 메인 레이아웃에 추가
        layout.addLayout(button_layout)

        # 최종 레이아웃 설정
        tab_widget.setLayout(layout)

        # 각 탭의 위젯과 FigureCanvas를 딕셔너리로 반환
        return {
            "widget": tab_widget,
            "update_label": update_label,  
            "storage_figure": storage_figure,
            "storage_canvas": storage_canvas,
            "analysis_figure": analysis_figure,
            "analysis_canvas": analysis_canvas,
            "combo_box": combo_box,
            "truck_figure": truck_figure,
            "truck_canvas": truck_canvas,
            "rail_figure": rail_figure,
            "rail_canvas": rail_canvas,
            "month_combo_box": month_combo_box,  
            "delay_donut_figure": delay_donut_figure,
            "delay_donut_canvas": delay_donut_canvas,

        }

    def show_vessel_delay_donut_chart(self, table_name, selected_month, figure, canvas):
        try:
            conn = connect_db()
            # 선택된 월에 대한 vesseldelayreason별 개수 조회
            query = f"""
            SELECT COALESCE(NULLIF(TRIM([vesseldelayreason]), ''), 'Others') as vesseldelayreason, COUNT(*) as count
            FROM {table_name}
            WHERE 
                [initialeta] IS NOT NULL
                AND [etaport] IS NOT NULL
                AND date([etaport]) > date([initialeta])
                AND strftime('%Y-%m', date([etaport])) = ?
            GROUP BY vesseldelayreason
            """
            df = pd.read_sql(query, conn, params=(selected_month,))
            conn.close()

            if not df.empty:
                reasons = df['vesseldelayreason'].fillna('Others').tolist()
                counts = df['count'].tolist()
                total_count = sum(counts)  # 총 개수 계산

                # 도넛 차트 생성
                figure.clear()
                ax = figure.add_subplot(111)

                # 색상 설정 (원하는 경우 수정 가능)
                colors = matplotlib.cm.Set3(np.linspace(0, 1, len(reasons)))

                # 도넛 차트 그리기
                wedges, texts = ax.pie(
                    counts, labels=reasons, startangle=90, colors=colors, wedgeprops=dict(width=0.3)
                )

                # 각 웨지에 reason 속성 추가
                for wedge, reason in zip(wedges, reasons):
                    wedge.reason = reason

                # 제목 설정
                ax.set_title(f"VESSEL DELAY", color='white')

                # 배경색 설정
                figure.patch.set_facecolor('#19232D')
                ax.set_facecolor('#19232D')

                # 텍스트 색상 설정
                for text in texts:
                    text.set_color('white')

                # 도넛 차트 중심에 총 개수 표시
                ax.text(0, 0, f"{total_count} EA", ha='center', va='center', fontsize=14, color='white')

                # 차트 그리기
                canvas.draw()

                # 인터랙티브 툴팁 및 강조 효과 추가
                annot = ax.annotate(
                    "", xy=(0, 0), xytext=(10, 10), textcoords="offset points",
                    bbox=dict(boxstyle="round", fc="w"),
                    arrowprops=dict(arrowstyle="->"))
                annot.set_visible(False)

                # 원래의 웨지 속성을 저장하기 위한 딕셔너리
                original_wedge_props = {}

                def update_annot(wedge):
                    angle = (wedge.theta2 - wedge.theta1) / 2. + wedge.theta1
                    x = np.cos(np.deg2rad(angle)) * 0.6  # 강조를 위해 위치 조정
                    y = np.sin(np.deg2rad(angle)) * 0.6
                    annot.xy = (x, y)
                    reason = getattr(wedge, 'reason', 'Others')
                    idx = reasons.index(reason)
                    text = f"{reason}: {counts[idx]}건"
                    annot.set_text(text)
                    annot.get_bbox_patch().set_facecolor('#ffffff')
                    annot.get_bbox_patch().set_alpha(0.9)

                def hover(event):
                    vis = annot.get_visible()
                    if event.inaxes == ax:
                        is_found = False
                        for wedge in wedges:
                            if wedge.contains_point((event.x, event.y)):
                                update_annot(wedge)
                                annot.set_visible(True)

                                # 웨지 강조 효과 적용
                                if wedge not in original_wedge_props:
                                    # 원래의 속성 저장
                                    original_wedge_props[wedge] = {
                                        'facecolor': wedge.get_facecolor(),
                                        'linewidth': wedge.get_linewidth(),
                                        'edgecolor': wedge.get_edgecolor(),
                                        'alpha': wedge.get_alpha(),
                                        'zorder': wedge.get_zorder()
                                    }
                                    # 강조를 위해 웨지의 속성 변경
                                    wedge.set_edgecolor('yellow')
                                    wedge.set_linewidth(2)
                                    wedge.set_alpha(0.8)
                                    wedge.set_zorder(10)  # 웨지를 위로 올려 강조

                                is_found = True
                            else:
                                # 강조 효과를 적용한 웨지를 원래 상태로 복구
                                if wedge in original_wedge_props:
                                    wedge.set_facecolor(original_wedge_props[wedge]['facecolor'])
                                    wedge.set_linewidth(original_wedge_props[wedge]['linewidth'])
                                    wedge.set_edgecolor(original_wedge_props[wedge]['edgecolor'])
                                    wedge.set_alpha(original_wedge_props[wedge]['alpha'])
                                    wedge.set_zorder(original_wedge_props[wedge]['zorder'])
                                    del original_wedge_props[wedge]

                        if not is_found and vis:
                            annot.set_visible(False)
                            canvas.draw_idle()
                        canvas.draw_idle()
                    else:
                        if vis:
                            annot.set_visible(False)
                            canvas.draw_idle()

                # 이벤트 핸들러 연결 전에 기존 연결 해제
                if table_name in self.donut_click_cid:
                    canvas.mpl_disconnect(self.donut_click_cid[table_name])

                # 이벤트 핸들러 연결 및 ID 저장
                self.donut_click_cid[table_name] = canvas.mpl_connect(
                    "button_press_event",
                    lambda event: self.on_donut_click(event, selected_month, table_name)
                )

                # 마우스 이벤트 연결
                canvas.mpl_connect("motion_notify_event", hover)

            else:
                # 선택된 월에 대한 데이터가 없는 경우 처리
                figure.clear()
                ax = figure.add_subplot(111)
                ax.text(
                    0.5, 0.5, 'No data', horizontalalignment='center', verticalalignment='center',
                    transform=ax.transAxes, color='white'
                )
                ax.set_title(f"{selected_month} vesseldelayreasonS", color='white')
                ax.axis('off')
                figure.patch.set_facecolor('#19232D')
                ax.set_facecolor('#19232D')
                canvas.draw()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while creating the donut chart: {str(e)}")

    def on_donut_click(self, event, selected_month, table_name):
        if event.dblclick and event.inaxes is not None:
            wedges = event.inaxes.patches
            for wedge in wedges:
                if wedge.contains_point((event.x, event.y)):
                    self.open_delay_detail_window(selected_month, table_name)
                    break

    def open_delay_detail_window(self, selected_month, table_name):
        try:
            self.current_table_name = table_name  # 현재 테이블 이름 저장
            conn = connect_db()

            # 쿼리 수정 - vesseldelayreason이 NULL이거나 빈 문자열인 경우도 포함
            query = f"""
            SELECT 
                COALESCE(NULLIF(TRIM([vesseldelayreason]), ''), 'Others') as vesseldelayreason,
                [initialeta], 
                [etaport],
                julianday([etaport]) - julianday([initialeta]) AS delay_days,
                strftime('%Y-%m', date([etaport])) AS month
            FROM {table_name}
            WHERE [initialeta] IS NOT NULL
            AND [etaport] IS NOT NULL
            AND date([etaport]) > date([initialeta])
            AND strftime('%Y-%m', date([etaport])) = ?
            """
            df = pd.read_sql(query, conn, params=(selected_month,))
            conn.close()

            if not df.empty:
                # 딜레이 일자 범주화
                # 유효한 지연일만 필터링 (음수와 NULL 제외)
                df = df[
                    (df['delay_days'].notna()) &  # NULL 제외
                    (df['delay_days'] > 0)        # 음수와 0 제외
                ]
                conditions = [
                    (df['delay_days'] <= 3) & (df['delay_days'] >= 1),
                    (df['delay_days'] >= 4) & (df['delay_days'] <= 7),
                    (df['delay_days'] > 7)
                ]
                categories = ['3 days or less', '4 to 7 days', '7 days or more']
                df['delay_category'] = np.select(conditions, categories, default='Others')

                # 피벗 테이블 생성 (총계 포함)
                pivot_df = df.pivot_table(
                    index='delay_category',
                    columns='vesseldelayreason',
                    values='delay_days',
                    aggfunc='count',
                    fill_value=0
                    
                )
                # 행 방향 합계 추가 (각 delay_category별 총합)
                pivot_df['Total'] = pivot_df.sum(axis=1)

                # 열 방향 합계 추가 (각 vesseldelayreason별 총합)
                pivot_df.loc['Total'] = pivot_df.sum()

                # 딜레이 일자 범주의 순서 정렬
                category_order = categories + ['Total']
                pivot_df = pivot_df.reindex(category_order).fillna(0)

                # 결과를 새로운 창에 표시
                self.show_delay_category_window(selected_month, pivot_df)
            else:
                QMessageBox.information(self, "Information", f"No data for '{selected_month}'.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while loading data: {str(e)}")

    def show_delay_category_window(self, selected_month, pivot_df):
        # 새로운 QDialog 생성
        self.delay_dialog = QDialog(self)
        dialog = self.delay_dialog
        dialog.setWindowTitle(f"'{selected_month}' Delay by Period")
        
        # 창 크기를 더 크게 설정
        dialog.setGeometry(100, 100, 1000, 400)  # 너비를 1000으로 증가

        # 레이아웃 설정
        layout = QVBoxLayout()

        # QTableWidget 생성
        table_widget = QTableWidget()
        num_rows, num_cols = pivot_df.shape
        table_widget.setRowCount(num_rows)
        table_widget.setColumnCount(num_cols)

        # 열 헤더 설정
        table_widget.setHorizontalHeaderLabels(pivot_df.columns.tolist())
        # 행 헤더 설정
        table_widget.setVerticalHeaderLabels(pivot_df.index.tolist())

        # 데이터 채우기
        for i, delay_category in enumerate(pivot_df.index):
            for j, reason in enumerate(pivot_df.columns):
                value = pivot_df.loc[delay_category, reason]
                item = QTableWidgetItem(str(int(value)))
                item.setTextAlignment(Qt.AlignCenter)  # 텍스트 중앙 정렬
                table_widget.setItem(i, j, item)

        # 더블 클릭 이벤트 핸들러 정의
        def cell_double_clicked(row, column):
            delay_category = pivot_df.index[row]
            reason = pivot_df.columns[column]
            count = pivot_df.iloc[row, column]

            # 합계 행이나 합계 열은 제외
            if delay_category == 'Total' or reason == 'Total':
                return

            # 상세 데이터 조회 및 표시
            self.show_detail_data(selected_month, delay_category, reason)

        # 이벤트 핸들러 연결
        table_widget.cellDoubleClicked.connect(cell_double_clicked)

        # 테이블 설정
        table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)  # 컬럼 너비 자동 조정
        table_widget.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)    # 행 높이 자동 조정
        table_widget.horizontalHeader().setStretchLastSection(True)  # 마지막 컬럼을 남은 공간에 맞게 늘림
        

        # 레이아웃에 위젯 추가
        layout.addWidget(table_widget)
        dialog.setLayout(layout)

        # 창 표시 (비모달로)
        dialog.show()

    def show_detail_data(self, selected_month, delay_category, reason):
        try:
            conn = connect_db()
            # 딜레이 일자 범주에 따른 조건 설정
            if delay_category == '3 days or less':
                delay_condition = "(julianday([etaport]) - julianday([initialeta])) <= 3 AND (julianday([etaport]) - julianday([initialeta])) >= 1"
            elif delay_category == '4 to 7 days':
                delay_condition = "(julianday([etaport]) - julianday([initialeta])) >= 4 AND (julianday([etaport]) - julianday([initialeta])) <= 7"
            elif delay_category == '7 days or more':
                delay_condition = "(julianday([etaport]) - julianday([initialeta])) > 7"
            else:
                # Unknown 카테고리도 포함
                delay_condition = "1=1"  # 모든 조건을 허용

            # reason이 'Unknown'인 경우와 아닌 경우에 대한 쿼리 조건 분리
            if reason == 'Others':
                reason_condition = "([vesseldelayreason] IS NULL OR TRIM([vesseldelayreason]) = '')"
            else:
                reason_condition = "[vesseldelayreason] = ?"

            # 데이터베이스 쿼리
            query = f"""
            SELECT *
            FROM {self.current_table_name}
            WHERE {reason_condition}
            AND [initialeta] IS NOT NULL
            AND [etaport] IS NOT NULL
            AND date([etaport]) != date([initialeta])
            AND strftime('%Y-%m', date([etaport])) = ?
            AND {delay_condition}
            """

            # reason이 'Unknown'인 경우와 아닌 경우에 따라 파라미터 설정
            if reason == 'Others':
                params = (selected_month,)
            else:
                params = (reason, selected_month)

            df = pd.read_sql(query, conn, params=params)
            conn.close()

            if not df.empty:
                self.data_window = DataDisplayWindow(df, self)
                self.data_window.setModal(False)
                self.data_window.show()
            else:
                QMessageBox.information(self, "Information", "No data for the selected condition.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while loading data: {str(e)}")

    def update_vessel_delay_report(self, table_name):
        try:
            conn = connect_db()

            # DELAY 계산을 위한 쿼리
            query = f"""
            SELECT [initialeta], [etaport],
                   julianday([etaport]) - julianday([initialeta]) AS delay_days
            FROM {table_name}
            WHERE [initialeta] IS NOT NULL
            AND [etaport] IS NOT NULL
            AND date([etaport]) != date([initialeta])  -- DELAY 조건: 두 날짜가 같지 않을 때
            """

            df = pd.read_sql(query, conn)
            conn.close()

            # DELAY 개수 (initialeta의 개수) 및 평균 DELAY 계산
            vessel_delay_count = df['initialeta'].count()  # DELAY로 간  [initialeta]의 개수
            if vessel_delay_count > 0:
                average_delay_days = df['delay_days'].mean()
            else:
                average_delay_days = 0

            # 4사분면 레이블 업데이트
            self.vessel_delay_label.setText(f"VESSEL DELAY: {vessel_delay_count}")
            self.average_delay_label.setText(f"Average DELAY days: {average_delay_days:.2f} days")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while updating the vessel delay report: {str(e)}")

    def refresh_charts(self, table_name):
        if table_name in self.tab_widgets:
            tab = self.tab_widgets[table_name]
            self.show_storage_cost_chart(table_name, tab["storage_figure"], tab["storage_canvas"])
            self.show_chart(table_name, tab["analysis_figure"], tab["analysis_canvas"])
            if tab["combo_box"]:
                self.update_dual_donut_chart(table_name, tab["combo_box"], tab["truck_figure"], tab["truck_canvas"],
                                             tab["rail_figure"], tab["rail_canvas"])

            # 도넛 차트 갱신 추가
            if "month_combo_box" in tab and "delay_donut_figure" in tab and "delay_donut_canvas" in tab:
                selected_month = tab["month_combo_box"].currentText()
                self.show_vessel_delay_donut_chart(table_name, selected_month, tab["delay_donut_figure"],
                                                   tab["delay_donut_canvas"])

            # 업데이트 시간 라벨 갱신
            tab["update_label"].setText(f"Last updated: {self.get_last_update_time(table_name)}")

    def show_modality_donut_charts_for_tab(self, tab, table_name):
        # 첫 번째 destinationport에 대한 도넛 차트 표시
        if "port1" in tab and tab["port1"]:
            port1 = tab["port1"]
            if tab["donut_figure_truck1"] and tab["donut_canvas_truck1"]:
                self.show_modality_donut_chart_by_port(
                    table_name,
                    "TRUCK",
                    port1,
                    tab["donut_figure_truck1"],
                    tab["donut_canvas_truck1"]
                )
            if tab["donut_figure_rail1"] and tab["donut_canvas_rail1"]:
                self.show_modality_donut_chart_by_port(
                    table_name,
                    "RAIL",
                    port1,
                    tab["donut_figure_rail1"],
                    tab["donut_canvas_rail1"]
                )

        # 두 번째 destinationport에 대한 도넛 차트 표시
        if "port2" in tab and tab["port2"]:
            port2 = tab["port2"]
            if tab["donut_figure_truck2"] and tab["donut_canvas_truck2"]:
                self.show_modality_donut_chart_by_port(
                    table_name,
                    "TRUCK",
                    port2,
                    tab["donut_figure_truck2"],
                    tab["donut_canvas_truck2"]
                )
            if tab["donut_figure_rail2"] and tab["donut_canvas_rail2"]:
                self.show_modality_donut_chart_by_port(
                    table_name,
                    "RAIL",
                    port2,
                    tab["donut_figure_rail2"],
                    tab["donut_canvas_rail2"]
                )

    def open_master_file_window(self, table_name):
        """
        MASTER 버튼 클릭 시 테이블의 전체 데이터를 보여주는 창을 엽니다.
        """
        try:
            conn = connect_db()
            query = f"SELECT * FROM {table_name}"
            df = pd.read_sql(query, conn)
            conn.close()

            if not df.empty:
                # 기존 창이 있다면 닫기
                if hasattr(self, 'data_window') and self.data_window is not None:
                    self.data_window.close()
                
                # 새 창 생성
                self.data_window = DataDisplayWindow(df, self)
                self.data_window.setModal(False)
                self.data_window.show()
            else:
                QMessageBox.information(self, "Information", f"{table_name} has no data.")
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while loading data: {str(e)}")

    def show_estimated_storage_cost(self, table_name):
        """
        스토리지 비용을 계산하고 표시하는 메서드
        """
        try:
            # StorageUtils를 사용하여 개별 스토리지 비용 데이터 가져오기
            storage_df = StorageUtils.get_individual_storage_data(table_name)
            
            if not storage_df.empty:
                # 2024년 이후 데이터만 필터링
                storage_df = storage_df[pd.to_datetime(storage_df['terminalappointment']).dt.year >= 2024]
                
                # storage_cost가 0보다 큰 데이터만 필터링
                storage_df = storage_df[storage_df['storage_cost'] > 0]
                
                if storage_df.empty:
                    QMessageBox.information(self, "Information", "No storage cost data after 2024.")
                    return

                # 필요한 칼럼만 선택
                selected_columns = [
                    'division',
                    'remark',
                    'delays/fee',
                    'destinationport',
                    'shippingline',
                    'terminal',
                    'origin',
                    'container',
                    'shippingdate',
                    'initialeta',
                    'etaport',
                    'vesseldelayreason',
                    'unloadingterminal',
                    'terminalappointment',
                    'eta',
                    'month',
                    'total_stay_days',
                    'days_over',
                    'storage_cost'
                ]
                storage_df = storage_df[selected_columns]

                # 월별 데이터 분리
                monthly_data = {
                    str(month): group for month, group in storage_df.groupby('month')
                }
                
                # 새로운 창에 데이터 표시 (윈도우 플래그 추가)
                self.data_window = MonthlyDataWindow(monthly_data, self)
                self.data_window.setWindowFlags(
                    Qt.Window |  # 기본 윈도우
                    Qt.WindowMinMaxButtonsHint |  # 최소화/최대화 버튼
                    Qt.WindowCloseButtonHint |  # 닫기 버튼
                    Qt.WindowSystemMenuHint  # 시스템 메뉴 (복사/붙여넣기 포함)
                )
                self.data_window.setModal(False)
                # 클립보드 작업 활성화
                self.data_window.setWindowFlag(Qt.WindowContextHelpButtonHint, False)
                # 테이블 위젯에서 복사/붙여넣기 활성화
                if hasattr(self.data_window, 'table_widget'):
                    self.data_window.table_widget.setContextMenuPolicy(Qt.ActionsContextMenu)
                    copy_action = QAction("Copy", self.data_window.table_widget)
                    copy_action.setShortcut(QKeySequence.Copy)
                    copy_action.triggered.connect(self.data_window.copy_selection)
                    self.data_window.table_widget.addAction(copy_action)
                
                self.data_window.show()
            else:
                QMessageBox.information(self, "Information", "No containers with storage costs.")
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while loading data: {str(e)}")

    def upload_excel_file(self):
        try:
            file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
            if file_path:
                xls = pd.ExcelFile(file_path)
                sheet_names = xls.sheet_names
                sheet_to_table = {tab_name: table_name for tab_name, table_name in self.tab_info.items()}
                sheet_to_table['Billing_storage'] = 'Billing_storage'

                conn = connect_db()
                cursor = conn.cursor()

                # Log 테이블 존재 여부 확인 및 생성
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS Log (
                        field1 TEXT,
                        table1 TEXT,
                        table2 TEXT,
                        table3 TEXT,
                        table4 TEXT,
                        table5 TEXT,
                        table6 TEXT,
                        table7 TEXT,
                        table8 TEXT,
                        table9 TEXT,
                        table10 TEXT
                    )
                """)

                # Log 테이블에 'updated_at' 행이 있는지 확인
                cursor.execute("SELECT COUNT(*) FROM Log WHERE field1 = 'updated_at'")
                count = cursor.fetchone()[0]
                if count == 0:
                    # 'updated_at' 이 없으면 추가
                    cursor.execute("""
                        INSERT INTO Log (field1) 
                        VALUES ('updated_at')
                    """)
                    conn.commit()

                for sheet_name in sheet_names:
                    if sheet_name in sheet_to_table:
                        table_name = sheet_to_table[sheet_name]

                        # 1. 기존 테이블의 첫 번째 행과 컬럼 정보 가져오기
                        cursor.execute(f"PRAGMA table_info({table_name})")
                        columns_info = cursor.fetchall()
                        db_columns = [column[1] for column in columns_info]

                        cursor.execute(f"SELECT * FROM {table_name} LIMIT 1")
                        first_row = cursor.fetchone()

                        # 컬럼 수 확인 로직 추가
                        if len(db_columns) != len(columns_info):
                            QMessageBox.warning(self, "Warning", 
                                f"Column count mismatch in sheet '{sheet_name}'. \n"
                                f"Expected {len(db_columns)} columns, but got {len(columns_info)} columns.")
                            continue

                        # 컬럼 매핑 전에 유효성 검사
                        try:
                            df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
                            df.columns = db_columns[:len(df.columns)]
                            
                            # destinationport 컬럼이 있는 경우 공백 제거
                            if 'destinationport' in df.columns:
                                df = df.copy()
                                df['destinationport'] = df['destinationport'].astype(str).str.strip()
                                
                        except Exception as e:
                            QMessageBox.warning(self, "Warning", 
                                f"Failed to map columns in sheet '{sheet_name}': {str(e)}")
                            continue

                        if first_row:
                            # Fixed='F'가 아닌 데이터만 삭제
                            cursor.execute(f"DELETE FROM {table_name} WHERE Fixed IS NULL OR Fixed != 'F'")
                            
                            # 새 데이터 삽입
                            df.to_sql(table_name, conn, if_exists='append', index=False)
                        else:
                            # 테이블이 비어있는 경우
                            # 기존 첫 번째 행 데이터를 가져와서 DataFrame 생성
                            first_row_df = pd.DataFrame([first_row], columns=db_columns)
                            
                            # 첫 번째 행과 새 데이터 결합
                            df = pd.concat([first_row_df, df], ignore_index=True)
                            
                            # 전체 데이터 삽입
                            df.to_sql(table_name, conn, if_exists='replace', index=False)

                        # 현재 시간 업데이트
                        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        cursor.execute(f"UPDATE Log SET {table_name} = ? WHERE field1 = 'updated_at'", (current_time,))
                        conn.commit()

                        # 업데이트 라벨 갱신
                        if table_name in self.tab_widgets:
                            tab = self.tab_widgets[table_name]
                            tab["update_label"].setText(f"Last updated: {current_time}")
                    else:
                        QMessageBox.warning(self, "Warning", f"No tab for sheet '{sheet_name}'.")

                conn.close()
                QMessageBox.information(self, "Success", "Excel file uploaded successfully.")
                self.reload_data()
            else:
                QMessageBox.warning(self, "Cancel", "Excel file upload canceled.")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while uploading the Excel file: {str(e)}")

    def update_billing_storage_table(self, df):
        """
        Billing_storage 테이블을 업데이트하는 함수
        """
        try:
            conn = sqlite3.connect(db_file)

            # 기존의 Billing_storage 테이블이 있으면 삭제하고 새로 만듭니다.
            cursor = conn.cursor()
            cursor.execute("DROP TABLE IF EXISTS Billing_storage")

            # DataFrame을 데이터베이스에 저장합니다.
            df.to_sql('Billing_storage', conn, index=False)

            conn.close()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while updating the Billing_storage table: {str(e)}")

    def standardize_column_names(self, df):
        # 칼럼 이름에서 공백 제거, 소문자로 변환, 특수문자 제거
        df.columns = df.columns.str.strip().str.replace(' ', '').str.replace(r'[^a-zA-Z0-9]', '').str.lower()
        return df

    def show_chart(self, table_name, figure, canvas):
        # 컨테이너 수를 표시하는 차트로 업데이트
        self.show_container_counts_chart(table_name, figure, canvas)
    
        # StorageUtils 클래스를 사용하여 데이터 가져오기
        df = StorageUtils.get_monthly_storage_data(table_name)
    
        return df

    def show_storage_cost_chart(self, table_name, figure, canvas):
        df = StorageUtils.get_monthly_storage_data(table_name)
        if not df.empty:
            # 'month'를 datetime으로 변환하고 연도와 월 추출
            df['month'] = pd.to_datetime(df['month'])
            df['year'] = df['month'].dt.year
            df['month_only'] = df['month'].dt.month

            # 2024년 이후의 데이터만 사용 (2025년 이후 포함)
            df = df[df['year'] >= 2024]

            if df.empty:
                # 데이터가 없는 경우 처리
                figure.clear()
                ax = figure.add_subplot(111)
                ax.text(0.5, 0.5, 'No data after 2024.', horizontalalignment='center', verticalalignment='center',
                        transform=ax.transAxes, color='white')
                ax.axis('off')
                figure.patch.set_facecolor('#19232D')
                ax.set_facecolor('#19232D')
                canvas.draw()
                return

            # 연도 리스트 가져오기 (2024년 이후 모든 연도)
            years = df['year'].unique()
            years.sort()

            # 모든 연도에 대해 1월부터 12월까지의 월을 생성
            full_months = pd.DataFrame()
            for year in years:  # 2024년부터 마지막 연도까지 모두 포함
                months = pd.DataFrame({
                    'month': pd.date_range(start=f'{year}-01-01', end=f'{year}-12-31', freq='MS')
                })
                months['year'] = year
                months['month_only'] = months['month'].dt.month
                full_months = pd.concat([full_months, months], ignore_index=True)

            # 원래 데이터프레임을 모든 달을 포함하는 데이터프레임과 병합
            df = pd.merge(full_months, df, on=['year', 'month', 'month_only'], how='left')
            df['total_storage_cost'] = df['total_storage_cost'].fillna(0)
            df['container_count'] = df['container_count'].fillna(0)

            # 현재 월의 비용 계산 (모든 연도 포함)
            current_month = datetime.now().month
            current_year = datetime.now().year
            current_month_cost = df[
                (df['month_only'] == current_month) & 
                (df['year'] == current_year)
            ]['total_storage_cost'].sum()

            # 차트 초기화
            figure.clear()
            ax1 = figure.add_subplot(111)

            # 색상 리스트 (필요에 따라 색상 추가)
            color_storage_list = ['cyan', 'yellow', 'magenta', 'green']
            color_container_list = ['gray', 'orange', 'purple', 'blue']

            # 라인과 바 객체를 저장할 리스트 초기화
            self.tab_widgets[table_name]['storage_lines'] = []
            self.tab_widgets[table_name]['storage_bars'] = []

            # 그래프의 가시성 상태를 저장하는 딕셔리 초기화
            if 'storage_visibility' not in self.tab_widgets[table_name]:
                self.tab_widgets[table_name]['storage_visibility'] = {}

            for idx, year in enumerate(years):
                df_year = df[df['year'] == year]

                # 스토리지 비용 라인 그래프
                color_storage = color_storage_list[idx % len(color_storage_list)]
                line, = ax1.plot(df_year['month_only'], df_year['total_storage_cost'], marker='o',
                                 color=color_storage, label=f'Storage Cost {year}')
                self.tab_widgets[table_name]['storage_lines'].append({'line': line, 'year': year})
                self.tab_widgets[table_name]['storage_visibility'][f'Storage Cost {year}'] = True

            ax1.set_xlabel('Month', color='white')
            ax1.set_ylabel('Storage Cost (MXN)', color='white')
            ax1.tick_params(axis='y', labelcolor='white')
            ax1.tick_params(axis='x', colors='white')
            ax1.set_xticks(range(1, 13))  # x축 1부터 12까지 설정
            ax1.set_xticklabels(['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                                 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'])

            # 컨테이너 개수 막대 그래프
            ax2 = ax1.twinx()
            width = 0.35  # 막대 너비 조절

            for idx, year in enumerate(years):
                df_year = df[df['year'] == year]
                color_container = color_container_list[idx % len(color_container_list)]
                # 막대 그래프의 위치를 연도별로 약간씩 조정
                positions = df_year['month_only'] + (idx - len(years)/2) * width
                bars = ax2.bar(positions, df_year['container_count'],
                               width=width, align='center', alpha=0.5, color=color_container,
                               label=f'Container Count {year}')
                self.tab_widgets[table_name]['storage_bars'].append({'bars': bars, 'year': year})
                self.tab_widgets[table_name]['storage_visibility'][f'Container Count {year}'] = True

            ax2.set_ylabel('Container Count', color='white')
            ax2.tick_params(axis='y', labelcolor='white')

            # **[수정된 부분 시작]**
            # 타이틀과 레이아웃 설정
            ax1.set_title(f"This Month's Storage Cost: {current_month_cost:,.2f} MXN", color='white')
            # **[수정된 부분 끝]**

            figure.tight_layout()

            # 배경색 설정
            figure.patch.set_facecolor('#19232D')
            ax1.set_facecolor('#19232D')
            ax2.set_facecolor('#19232D')

            

            # 마우스 오른쪽 클릭 이벤트 핸들러 정의
            def on_canvas_click(event):
                if event.button == 3:  # 오른쪽 클릭
                    # 컨텍스트 메뉴 생성
                    menu = QMenu()

                    visibility = self.tab_widgets[table_name]['storage_visibility']

                    # 스토리지 비용 그래프 메뉴 항목 추가
                    for item in self.tab_widgets[table_name]['storage_lines']:
                        label = f'Storage Cost {item["year"]}'
                        action = QAction(label, menu)
                        action.setCheckable(True)
                        action.setChecked(visibility.get(label, True))

                        # 함수 내에서 현재의 item과 label을 캡쳐하기 위해 기본 인자로 전달
                        def toggle_storage_cost(checked, item=item, label=label):
                            visibility[label] = checked
                            item['line'].set_visible(checked)
                            canvas.draw()

                        action.triggered.connect(toggle_storage_cost)
                        menu.addAction(action)

                    # 컨테이너 개수 그래프 메뉴 항목 추가
                    for item in self.tab_widgets[table_name]['storage_bars']:
                        label = f'Container Count {item["year"]}'
                        action = QAction(label, menu)
                        action.setCheckable(True)
                        action.setChecked(visibility.get(label, True))

                        def toggle_container_count(checked, item=item, label=label):
                            visibility[label] = checked
                            for bar in item['bars']:
                                bar.set_visible(checked)
                            canvas.draw()

                        action.triggered.connect(toggle_container_count)
                        menu.addAction(action)

                    menu.exec_(QCursor.pos())

            # 캔버스에 이벤트 핸들러 연결
            canvas.mpl_connect('button_press_event', on_canvas_click)

            # 캔버스 그리기
            canvas.draw()

            # 더블클릭 이벤트 핸들러 추가
            def on_double_click(event):
                if event.dblclick:
                    self.show_storage_cost_analysis(table_name)
            
            # 이벤트 연결
            canvas.mpl_connect('button_press_event', on_double_click)

        else:
            # 데이터가 없는 경우 처리
            figure.clear()
            ax = figure.add_subplot(111)
            ax.text(0.5, 0.5, 'No data available', horizontalalignment='center', verticalalignment='center',
                    transform=ax.transAxes, color='white')
            ax.axis('off')
            figure.patch.set_facecolor('#19232D')
            ax.set_facecolor('#19232D')
            canvas.draw()

    def show_container_counts_chart(self, table_name, figure, canvas):
        try:
            conn = connect_db()
            today = datetime.today().date()
            tomorrow = today + timedelta(days=1)  # 내일 날짜 계산

            # 긴급 컨테이너 - CEDROS 조회
            query_emergency_cedros = f"""
            SELECT COUNT(*) as count
            FROM {table_name}
            WHERE ([urgentcargo] LIKE '%y%' OR [urgentcargo] LIKE '%Y%')
            AND [f.dest] = 'CEDROS'
            AND (
                (date([eta]) >= '{today}' AND [eta] IS NOT NULL)
                OR
                ([eta] IS NULL AND date([etaport]) >= '{today}')
            )
            """
            df_emergency_cedros = pd.read_sql(query_emergency_cedros, conn)
            emergency_count_cedros = df_emergency_cedros["count"].iloc[0]

            # 오늘 입고 예정 컨테이너 - CEDROS 조회
            query_incoming_today_cedros = f"""
            SELECT COUNT(*) as count
            FROM {table_name}
            WHERE (date([eta]) = '{today}') AND [eta] IS NOT NULL
            AND [f.dest] = 'CEDROS'
            """
            df_incoming_today_cedros = pd.read_sql(query_incoming_today_cedros, conn)
            incoming_today_count_cedros = df_incoming_today_cedros["count"].iloc[0]

            # 내일 입고 예정 컨테이너 - CEDROS 조회
            query_incoming_tomorrow_cedros = f"""
            SELECT COUNT(*) as count
            FROM {table_name}
            WHERE (date([eta]) = '{tomorrow}') AND [eta] IS NOT NULL
            AND [f.dest] = 'CEDROS'
            """
            df_incoming_tomorrow_cedros = pd.read_sql(query_incoming_tomorrow_cedros, conn)
            incoming_tomorrow_count_cedros = df_incoming_tomorrow_cedros["count"].iloc[0]

            conn.close()

            # 데이터를 딕셔너리로 정리
            data = {
                'CEDROS': {
                    'URGENT': emergency_count_cedros,
                    'ETA TODAY': incoming_today_count_cedros,
                    'ETA TOMORROW': incoming_tomorrow_count_cedros
                }
            }

            # 그래프 그리기
            figure.clear()
            ax = figure.add_subplot(111)

            labels = ['URGENT', 'ETA TODAY', 'ETA TOMORROW']
            cedros_values = [data['CEDROS'][label] for label in labels]

            x = range(len(labels))
            bar_width = 0.2  # 막대 너비 조정

            # 색상 설정
            colors = ['mediumturquoise']  # 원하는 색상으로 변경 가능
            bars = ax.bar(x, cedros_values, width=bar_width, color=colors, label='CEDROS')

            ax.set_ylabel('Container Count', color='white')
            ax.set_title('CEDROS Container Status', color='white')
            ax.set_xticks(x)
            ax.set_xticklabels(labels, rotation=0)
            ax.legend()

            ax.tick_params(axis='x', colors='white')
            ax.tick_params(axis='y', colors='white')
            figure.patch.set_facecolor('#19232D')  # 배경색 설정
            ax.set_facecolor('#19232D')

            canvas.draw()

            # 인터랙티브 툴팁 추가
            annot = ax.annotate("", xy=(0, 0), xytext=(0, 20), textcoords="offset points",
                                bbox=dict(boxstyle="round", fc="w"),
                                arrowprops=dict(arrowstyle="->"))
            annot.set_visible(False)

            def update_annot(bar):
                x = bar.get_x() + bar.get_width() / 2
                y = bar.get_height()
                annot.xy = (x, y)
                idx = bars.index(bar)
                text = f"{labels[idx]}: {int(y)}"
                annot.set_text(text)
                annot.get_bbox_patch().set_facecolor('#ffffff')
                annot.get_bbox_patch().set_alpha(0.9)

            def hover(event):
                vis = annot.get_visible()
                if event.inaxes == ax:
                    for bar in bars:
                        if bar.contains(event)[0]:
                            update_annot(bar)
                            annot.set_visible(True)
                            canvas.draw_idle()
                            break
                    else:
                        if vis:
                            annot.set_visible(False)
                            canvas.draw_idle()

            canvas.mpl_connect("motion_notify_event", hover)

            # **더블 클릭 이벤트 핸들러 추가**
            def on_bar_double_click(event):
                if event.dblclick:
                    if event.inaxes == ax:
                        for bar, label in zip(bars, labels):
                            if bar.contains(event)[0]:
                                # 막대를 더블 클릭하면 해당 데이터를 표시
                                self.show_cedros_data(table_name, label)
                                break

            canvas.mpl_connect('button_press_event', on_bar_double_click)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while loading data: {str(e)}")

    def show_cedros_data(self, table_name, category):
        """
        선택한 카테고리에 해당하는 CEDROS의 데이터를 가져와서 표시하는 메서드
        """
        try:
            conn = connect_db()
            today = datetime.today().date()
            tomorrow = today + timedelta(days=1)

            if category == 'URGENT':
                query = f"""
                SELECT *
                FROM {table_name}
                WHERE ([urgentcargo] LIKE '%y%' OR [urgentcargo] LIKE '%Y%')
                AND [f.dest] = 'CEDROS'
                AND (
                    (date([eta]) >= '{today}' AND [eta] IS NOT NULL)
                    OR
                    ([eta] IS NULL AND date([etaport]) >= '{today}')
                )
                """
            elif category == 'ETA TODAY':
                query = f"""
                SELECT *
                FROM {table_name}
                WHERE (date([eta]) = '{today}') AND [eta] IS NOT NULL
                AND [f.dest] = 'CEDROS'
                """
            elif category == 'ETA TOMORROW':
                query = f"""
                SELECT *
                FROM {table_name}
                WHERE (date([eta]) = '{tomorrow}') AND [eta] IS NOT NULL
                AND [f.dest] = 'CEDROS'
                """
            else:
                conn.close()
                QMessageBox.warning(self, "Warning", "Unknown category.")
                return

            df = pd.read_sql(query, conn)
            conn.close()

            if not df.empty:
                self.data_window = DataDisplayWindow(df, self)
                self.data_window.setModal(False)
                self.data_window.show()
            else:
                QMessageBox.information(self, "Information", f"No data for {category}.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while loading data: {str(e)}")

    def plot_bar_chart(self, data, title, x_label, y_label, figure, canvas):
        # 기존 차트 지우기
        figure.clear()
        ax = figure.add_subplot(111)

        # 가로 막대 그래프 그리기 (막대 높이를 줄이기 위해 height 값을 0.3으로 설정)
        bars = ax.barh(list(data.keys()), list(data.values()), color='cyan', height=0.1)
        ax.set_title(title, color='white')
        ax.set_xlabel(x_label, color='white')
        ax.set_ylabel(y_label, color='white')
        ax.tick_params(axis='x', colors='white')
        ax.tick_params(axis='y', colors='white')
        figure.patch.set_facecolor('#19232D')  # qdarkstyle 배경색과 맞추기
        ax.set_facecolor('#19232D')

        canvas.draw()

        # 인터랙티브 툴팁 추가
        annot = ax.annotate("", xy=(0, 0), xytext=(0, 15), textcoords="offset points",
                            bbox=dict(boxstyle="round", fc="w"),
                            arrowprops=dict(arrowstyle="->"))
        annot.set_visible(False)

        def update_annot(bar):
            x = bar.get_width()
            y = bar.get_y() + bar.get_height() / 2
            annot.xy = (x, y)
            text = f"총 {int(x)}개의 컨테이너"
            annot.set_text(text)
            annot.get_bbox_patch().set_facecolor('#ffffff')
            annot.get_bbox_patch().set_alpha(0.9)

        def hover(event):
            vis = annot.get_visible()
            if event.inaxes == ax:
                for bar in bars:
                    if bar.contains(event)[0]:
                        update_annot(bar)
                        annot.set_visible(True)
                        canvas.draw_idle()
                        break
                else:
                    if vis:
                        annot.set_visible(False)
                        canvas.draw_idle()

        canvas.mpl_connect("motion_notify_event", hover)

    def plot_chart(self, df, title, x_label, y_label, figure, canvas, interactive=False, current_month_cost=None):
        # 기존 차트 지우기
        figure.clear()
        ax = figure.add_subplot(111)
        line, = ax.plot(df.iloc[:, 0], df.iloc[:, 1], marker='o', linestyle='-', color='cyan')
        ax.set_title(title, color='white')
        ax.set_xlabel(x_label, color='white')
        ax.set_ylabel(y_label, color='white')
        ax.tick_params(axis='x', colors='white')
        ax.tick_params(axis='y', colors='white')
        figure.patch.set_facecolor('#19232D')  # qdarkstyle 배경색과 맞추기
        ax.set_facecolor('#19232D')
        figure.autofmt_xdate(rotation=45)


        canvas.draw()

        if interactive:
            # 인터랙티브 툴팁 추가
            annot = ax.annotate("", xy=(0, 0), xytext=(-90, -50), textcoords="offset points",
                                bbox=dict(boxstyle="round", fc="w"),
                                arrowprops=dict(arrowstyle="->"))
            annot.set_visible(False)

            def update_annot(ind):
                pos = line.get_xydata()[ind["ind"][0]]
                annot.xy = pos
                text = f"{x_label}: {pos[0]}\n{y_label}: {pos[1]:,.2f}"
                annot.set_text(text)
                annot.get_bbox_patch().set_facecolor('#ffffff')
                annot.get_bbox_patch().set_alpha(0.9)

            def hover(event):
                vis = annot.get_visible()
                if event.inaxes == ax:
                    cont, ind = line.contains(event)
                    if cont:
                        update_annot(ind)
                        annot.set_visible(True)
                        canvas.draw_idle()
                    else:
                        if vis:
                            annot.set_visible(False)
                            canvas.draw_idle()

            canvas.mpl_connect("motion_notify_event", hover)

    def show_modality_donut_chart_by_port(self, table_name, modality, port, figure, canvas):
        try:
            import numpy as np

            conn = connect_db()
            today = datetime.today().date()

            # modality와 destinationport별 총 데이터 개수 확인
            query_modality = f"""
            SELECT COUNT(*) as count
            FROM {table_name}
            WHERE [modality] = '{modality}' 
            AND [destinationport] = '{port}'
            """
            df_modality = pd.read_sql(query_modality, conn)
            modality_count = df_modality['count'].iloc[0]

            if modality_count == 0:
                # 데이터 없음 처리
                figure.clear()
                ax = figure.add_subplot(111)
                ax.text(0.5, 0.5, 'No data', horizontalalignment='center', verticalalignment='center',
                        transform=ax.transAxes, color='white')
                ax.set_title(f"{modality}", color='white', fontsize=12)
                ax.axis('off')
                figure.patch.set_facecolor('#19232D')
                ax.set_facecolor('#19232D')
                canvas.draw()
                conn.close()
                return

            # '입항 전' 계산 (etaport가 오늘 이후인 것들)
            query_pre_arrival = f"""
            SELECT COUNT(*) as count
            FROM {table_name}
            WHERE [modality] = '{modality}' 
            AND [destinationport] = '{port}'
            AND [etaport] IS NOT NULL 
            AND [etaport] != ''
            AND date([etaport]) > '{today}'
            """
            df_pre_arrival = pd.read_sql(query_pre_arrival, conn)
            pre_arrival_count = df_pre_arrival['count'].iloc[0]

            # '당일 반출' 계산 (terminalappointment가 오늘인 것들)
            query_same_day = f"""
            SELECT COUNT(*) as count
            FROM {table_name}
            WHERE [modality] = '{modality}' 
            AND [destinationport] = '{port}'
            AND [terminalappointment] IS NOT NULL 
            AND [terminalappointment] != ''
            AND date([terminalappointment]) = '{today}'
            """
            df_same_day = pd.read_sql(query_same_day, conn)
            same_day_count = df_same_day['count'].iloc[0]

            # '잔량' 계산 (terminalappointment가 오늘 이후인 것들)
            query_remaining = f"""
            SELECT COUNT(*) as count
            FROM {table_name}
            WHERE [modality] = '{modality}' 
            AND [destinationport] = '{port}'
            AND [terminalappointment] IS NOT NULL 
            AND [terminalappointment] != ''
            AND date([terminalappointment]) > '{today}'
            """
            df_remaining = pd.read_sql(query_remaining, conn)
            remaining_count = df_remaining['count'].iloc[0]

            conn.close()

            # 데이터 준비
            counts = {
                'Pre-arrival': pre_arrival_count,
                'Today export': same_day_count,
                'Remaining': remaining_count
            }

            # 도넛 차트 생성
            figure.clear()
            ax = figure.add_subplot(111)

            labels = list(counts.keys())
            sizes = list(counts.values())

            total = sum(sizes)
            if total == 0:
                # 데이터 없음 처리
                ax.text(0.5, 0.5, 'No data', horizontalalignment='center', verticalalignment='center',
                        transform=ax.transAxes, color='white')
            else:
                # 색상 지정 (원하는 색상으로 변경 가능)
                colors = ['#ff9999', '#66b3ff', '#99ff99']

                # 도넛 차트의 두께를 조정 (width 값을 더 작게 설정)
                wedge_width = 0.14  # 기본 께
                wedges, _ = ax.pie(
                    sizes,
                    startangle=90,
                    wedgeprops=dict(width=wedge_width),  # 두께 설정
                    colors=colors
                )

                # 각 웨지에 고유 ID 부여
                for i, wedge in enumerate(wedges):
                    wedge.set_gid(i)
                    wedge.category = labels[i]  # 카테고리 정보를 추가

                # 도넛 차트 중심에 텍스트 추가
                inner_labels = [f"{label}\n{size}" for label, size in zip(labels, sizes)]
                inner_text = '\n\n'.join(inner_labels)
                ax.text(0, 0, inner_text, ha='center', va='center', fontsize=8, color='white')

                # 인터랙티브 툴팁 및 강조 효과 추가
                annot = ax.annotate("", xy=(0, 0), xytext=(20, 20), textcoords="offset points",
                                    bbox=dict(boxstyle="round", fc="w"),
                                    arrowprops=dict(arrowstyle="->"))
                annot.set_visible(False)

                # 원래의 웨지 속성을 저장하기 위한 딕셔너리
                original_wedge_props = {}

                def update_annot(wedge, idx):
                    angle = (wedge.theta2 - wedge.theta1) / 2. + wedge.theta1
                    x = np.cos(np.deg2rad(angle)) * 0.7
                    y = np.sin(np.deg2rad(angle)) * 0.7
                    annot.xy = (x, y)
                    text = f"{labels[idx]}: {sizes[idx]}"
                    annot.set_text(text)
                    annot.get_bbox_patch().set_facecolor('#ffffff')
                    annot.get_bbox_patch().set_alpha(0.9)

                def hover(event):
                    is_found = False
                    for wedge in wedges:
                        if wedge.contains_point((event.x, event.y)):
                            idx = wedge.get_gid()
                            update_annot(wedge, idx)
                            annot.set_visible(True)

                            # 웨지 강조 효과 적용
                            if wedge not in original_wedge_props:
                                # 원래의 속성 저장
                                original_wedge_props[wedge] = {
                                    'facecolor': wedge.get_facecolor(),
                                    'linewidth': wedge.get_linewidth(),
                                    'edgecolor': wedge.get_edgecolor(),
                                    'alpha': wedge.get_alpha()
                                }
                            # 강조를 위해 웨지의 속성 변경
                            wedge.set_edgecolor('yellow')
                            wedge.set_alpha(0.7)  # 투명도 조절

                            is_found = True
                        else:
                            # 강조 효과를 적용한 웨지를 원래 상태로 복구
                            if wedge in original_wedge_props:
                                wedge.set_facecolor(original_wedge_props[wedge]['facecolor'])
                                wedge.set_linewidth(original_wedge_props[wedge]['linewidth'])
                                wedge.set_edgecolor(original_wedge_props[wedge]['edgecolor'])
                                wedge.set_alpha(original_wedge_props[wedge]['alpha'])
                                del original_wedge_props[wedge]

                    if not is_found:
                        annot.set_visible(False)

                    canvas.draw_idle()

                canvas.mpl_connect("motion_notify_event", hover)

                # **더블 클릭 이벤트 핸들러 추가**
                def on_wedge_double_click(event):
                    if event.dblclick:
                        if event.inaxes == ax:
                            for wedge in wedges:
                                if wedge.contains_point((event.x, event.y)):
                                    category = wedge.category
                                    # 웨지를 더블 클릭하면 해당 데이터를 표시
                                    self.show_modality_data(table_name, modality, port, category)
                                    break

                canvas.mpl_connect('button_press_event', on_wedge_double_click)

            ax.set_title(f"{modality}", color='white', fontsize=12)

            ax.axis('equal')  # 원형으로 그리기
            figure.patch.set_facecolor('#19232D')  # qdarkstyle 배경색과 맞춤
            ax.set_facecolor('#19232D')

            canvas.draw()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while creating the donut chart: {str(e)}")

    def show_modality_data(self, table_name, modality, port, category):
        """
        선택한 카테고리에 해당하는 데이터를 가져와서 표시하는 메서드
        """
        try:
            conn = connect_db()
            today = datetime.today().date()

            if category == 'Pre-arrival':
                query = f"""
                SELECT *
                FROM {table_name}
                WHERE [modality] = '{modality}' 
                AND [destinationport] = '{port}'
                AND [etaport] IS NOT NULL 
                AND [etaport] != ''
                AND date([etaport]) > '{today}'
                """
            elif category == 'Today export':
                query = f"""
                SELECT *
                FROM {table_name}
                WHERE [modality] = '{modality}' 
                AND [destinationport] = '{port}'
                AND [terminalappointment] IS NOT NULL 
                AND [terminalappointment] != ''
                AND date([terminalappointment]) = '{today}'
                """
            elif category == 'Remaining':
                query = f"""
                SELECT *
                FROM {table_name}
                WHERE [modality] = '{modality}' 
                AND [destinationport] = '{port}'
                AND [terminalappointment] IS NOT NULL 
                AND [terminalappointment] != ''
                AND date([terminalappointment]) > '{today}'
                """
            else:
                conn.close()
                QMessageBox.warning(self, "Warning", "Unknown category.")
                return

            df = pd.read_sql(query, conn)
            conn.close()

            if not df.empty:
                self.data_window = DataDisplayWindow(df, self)
                self.data_window.setModal(False)
                self.data_window.show()
            else:
                QMessageBox.information(self, "Information", f"No data for {category}.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while loading data: {str(e)}")

    def update_dual_donut_chart(self, table_name, combo_box, truck_figure, truck_canvas, rail_figure, rail_canvas):
        selected_port = combo_box.currentText()
        if not selected_port:
            return

        # 두 도넛 차트를 업데이트 (TRUCK과 RAIL)
        self.show_modality_donut_chart_by_port(table_name, "TRUCK", selected_port, truck_figure, truck_canvas)
        self.show_modality_donut_chart_by_port(table_name, "RAIL", selected_port, rail_figure, rail_canvas)
    
    def show_combined_storage_cost_chart(self):
        combined_data = pd.DataFrame()
        for table_name in self.tab_info.values():
            df = StorageUtils.get_monthly_storage_data(table_name)
            combined_data = pd.concat([combined_data, df], ignore_index=True)

        if not combined_data.empty:
            # 'month'를 datetime으로 변환하고 연도와 월 추출
            combined_data['month'] = pd.to_datetime(combined_data['month'])
            combined_data['year'] = combined_data['month'].dt.year
            combined_data['month_only'] = combined_data['month'].dt.month

            # 2024년 이후의 데이터만 사용
            combined_data = combined_data[combined_data['year'] >= 2024]

            if combined_data.empty:
                # 데이터가 없는 경우 처리
                self.combined_storage_figure.clear()
                ax = self.combined_storage_figure.add_subplot(111)
                ax.text(0.5, 0.5, 'No data after 2024.', horizontalalignment='center', verticalalignment='center',
                        transform=ax.transAxes, color='white')
                ax.axis('off')
                self.combined_storage_figure.patch.set_facecolor('#19232D')
                ax.set_facecolor('#19232D')
                self.combined_storage_canvas.draw()
                return

            # 모든 연도에 대해 1월부터 12월까지의 월을 생성
            years = combined_data['year'].unique()
            years.sort()
            full_months = pd.DataFrame()
            for year in years:
                months = pd.DataFrame({
                    'month': pd.date_range(start=f'{year}-01-01', end=f'{year}-12-01', freq='MS')
                })
                months['year'] = year
                months['month_only'] = months['month'].dt.month
                full_months = pd.concat([full_months, months], ignore_index=True)

            # 원래 데이터프레임을 모든 달을 포함하는 데이터프레임과 병합
            combined_data = pd.merge(full_months, combined_data, on=['year', 'month', 'month_only'], how='left')
            combined_data['total_storage_cost'] = combined_data['total_storage_cost'].fillna(0)
            combined_data['container_count'] = combined_data['container_count'].fillna(0)

            # 월별로 스토리지 비용과 컨테이너 개수를 합산
            combined_data = combined_data.groupby(['year', 'month_only']).agg({
                'total_storage_cost': 'sum',
                'container_count': 'sum'
            }).reset_index()

            # 이번 달 계산
            current_month = datetime.now().month
            current_year = datetime.now().year
            if current_year >= 2024:
                current_month_cost = combined_data[
                    (combined_data['year'] == current_year) & 
                    (combined_data['month_only'] == current_month)
                ]['total_storage_cost'].sum()
            else:
                current_month_cost = 0

            # 차트 그리기
            self.combined_storage_figure.clear()
            ax1 = self.combined_storage_figure.add_subplot(111)

            # 색상 리스트
            color_storage_list = ['cyan', 'yellow', 'magenta', 'green']
            color_container_list = ['gray', 'orange', 'purple', 'blue']

            # 라인과 바 객체를 저장할 리스트 초기화
            self.combined_storage_lines = []
            self.combined_storage_bars = []

            # 그래프의 가시성 상태를 저장하는 딕셔너리 초기화
            if 'combined_storage_visibility' not in self.__dict__:
                self.combined_storage_visibility = {}

            # 스토리지 비용 라인 그래프
            for idx, year in enumerate(years):
                df_year = combined_data[combined_data['year'] == year]
                color_storage = color_storage_list[idx % len(color_storage_list)]
                line, = ax1.plot(df_year['month_only'], df_year['total_storage_cost'], 
                            marker='o', color=color_storage, label=f'Storage Cost {year}')
                self.combined_storage_lines.append({'line': line, 'year': year})
                self.combined_storage_visibility[f'Storage Cost {year}'] = True

            # 컨테이너 개수 막대 그래프
            ax2 = ax1.twinx()
            width = 0.35  # 막대 너비 조절
            
            for idx, year in enumerate(years):
                df_year = combined_data[combined_data['year'] == year]
                color_container = color_container_list[idx % len(color_container_list)]
                
                # 막대 위치 조정: 연도별로 위치를 조금씩 이동
                positions = df_year['month_only'] + (idx - len(years)/2) * width
                
                bars = ax2.bar(positions, df_year['container_count'],
                            width=width, align='center', alpha=0.5, 
                            color=color_container,
                            label=f'Container Count {year}')
                self.combined_storage_bars.append({'bars': bars, 'year': year})
                self.combined_storage_visibility[f'Container Count {year}'] = True

            # 축 레이블 및 설정
            ax1.set_xlabel('Month', color='white')
            ax1.set_ylabel('Storage Cost (MXN)', color='white')
            ax1.tick_params(axis='y', labelcolor='white')
            ax1.tick_params(axis='x', colors='white')
            
            ax1.set_xticks(range(1, 13))
            ax1.set_xticklabels(['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                                'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'])

            ax2.set_ylabel('Container Count', color='white')
            ax2.tick_params(axis='y', labelcolor='white')

            # 차트 제목
            ax1.set_title(f"This Month's Total Storage Cost: {current_month_cost:,.2f} MXN", color='white')

            # 레이아웃 및 배경색 설정
            self.combined_storage_figure.tight_layout()
            self.combined_storage_figure.patch.set_facecolor('#19232D')
            ax1.set_facecolor('#19232D')
            ax2.set_facecolor('#19232D')

            # 마우스 오른쪽 클릭 이벤트 핸들러
            def on_canvas_click(event):
                if event.button == 3:  # 오른쪽 클릭
                    menu = QMenu()
                    visibility = self.combined_storage_visibility

                    # 스토리지 비용 그래프 메뉴 항목
                    for item in self.combined_storage_lines:
                        label = f'Storage Cost {item["year"]}'
                        action = QAction(label, menu)
                        action.setCheckable(True)
                        action.setChecked(visibility.get(label, True))

                        def toggle_storage_cost(checked, item=item, label=label):
                            visibility[label] = checked
                            item['line'].set_visible(checked)
                            self.combined_storage_canvas.draw()

                        action.triggered.connect(toggle_storage_cost)
                        menu.addAction(action)

                    # 컨테이너 개수 그래프 메뉴 항목
                    for item in self.combined_storage_bars:
                        label = f'Container Count {item["year"]}'
                        action = QAction(label, menu)
                        action.setCheckable(True)
                        action.setChecked(visibility.get(label, True))

                        def toggle_container_count(checked, item=item, label=label):
                            visibility[label] = checked
                            for bar in item['bars']:
                                bar.set_visible(checked)
                            self.combined_storage_canvas.draw()

                        action.triggered.connect(toggle_container_count)
                        menu.addAction(action)

                    menu.exec_(QCursor.pos())
                elif event.dblclick:  # 더블 클릭
                    # 전체 데이터에 대한 분석 창 열기
                    analysis_window = StorageCostAnalysisWindow("all_tables", parent=self, tab_info=self.tab_info)
                    analysis_window.show()

            # 캔버스에 이벤트 핸들러 연결
            self.combined_storage_canvas.mpl_connect('button_press_event', on_canvas_click)

            # 그리기
            self.combined_storage_canvas.draw()

        else:
            # 데이터가 없는 경우 처리
            self.combined_storage_figure.clear()
            ax = self.combined_storage_figure.add_subplot(111)
            ax.text(0.5, 0.5, 'No data available', horizontalalignment='center', 
                    verticalalignment='center', transform=ax.transAxes, color='white')
            ax.axis('off')
            self.combined_storage_figure.patch.set_facecolor('#19232D')
            ax.set_facecolor('#19232D')
            self.combined_storage_canvas.draw()    

    def reload_data(self):
        """
        데이터베이스에서 데이터를 다시 로드하고 차트 및 표시를 업데이트합니다.
        """
        # 로딩 애니메이션 시작 (필요에 따라 로딩 애니메이션 추가 가능)
        # self.loading_label.show()
        # self.loading_movie.start()

        # 각 탭의 데이터를 업데이트
        for table_name, tab in self.tab_widgets.items():
            # 업데이트 시간 라벨 갱신
            tab["update_label"].setText(f"Last updated: {self.get_last_update_time(table_name)}")

            # 차트 갱신
            self.show_storage_cost_chart(table_name, tab["storage_figure"], tab["storage_canvas"])
            self.show_chart(table_name, tab["analysis_figure"], tab["analysis_canvas"])

            if tab["combo_box"]:
                # 콤보박스의 현재 선택된 포트에 따라 도넛 차트 업데이트
                self.update_dual_donut_chart(
                    table_name,
                    tab["combo_box"],
                    tab["truck_figure"],
                    tab["truck_canvas"],
                    tab["rail_figure"],
                    tab["rail_canvas"]
                )

            # VESSEL DELAY 도넛 차트 갱신
            if "month_combo_box" in tab and "delay_donut_figure" in tab and "delay_donut_canvas" in tab:
                selected_month = tab["month_combo_box"].currentText()
                self.show_vessel_delay_donut_chart(table_name, selected_month, tab["delay_donut_figure"],
                                                   tab["delay_donut_canvas"])

        # 전체 스토리지 비용 차트 갱신
        self.show_combined_storage_cost_chart()

        # 로딩 애니메이션 종료 (필요에 따라 로딩 애니메이션 추가 가능)
        # self.loading_movie.stop()
        # self.loading_label.hide()

        QMessageBox.information(self, "Data Updated", "Data has been successfully updated.")

    def get_unique_origins(self):
        try:
            conn = connect_db()
            origins_set = set()
            cursor = conn.cursor()
            for table_name in self.tab_info.values():
                query = f"SELECT DISTINCT origin FROM {table_name}"
                cursor.execute(query)
                results = cursor.fetchall()
                for row in results:
                    origin = row[0]
                    if origin:
                        standardized_origin = Mapping.standardize_origin_name(origin)
                        origins_set.add(standardized_origin)
            conn.close()
            return sorted(origins_set)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while getting the origin list: {str(e)}")
            return []

    def get_shipping_lines_for_origin(self, origin):
        try:
            conn = connect_db()
            cursor = conn.cursor()
            
            shipping_lines = []
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'table%'")
            tables = cursor.fetchall()
            
            for (table_name,) in tables:
                cursor.execute(f"SELECT DISTINCT shippingline FROM [{table_name}] WHERE origin = ?", (origin,))
                results = cursor.fetchall()
                shipping_lines.extend([sl[0] for sl in results if sl[0]])
            
            conn.close()
            
            # 결과를 반환하기 전에 표준화 적용
            standardized_shipping_lines = [Mapping.standardize_shipping_line_name(sl) for sl in shipping_lines]
            return sorted(set(standardized_shipping_lines))  # 중복 제거 후 정렬하여 반환
            
        except Exception as e:
            print(f"Error getting shipping lines: {str(e)}")
            return []

    def calculate_analyses(self, origin, shipping_lines, destination_ports, start_date, end_date):
        try:
            conn = connect_db()
            data_frames = []
            cursor = conn.cursor()

            # Origin과 Shipping Line 이름을 표준화
            standardized_origin = Mapping.standardize_origin_name(origin)
            standardized_shipping_lines = [Mapping.standardize_shipping_line_name(sl) for sl in shipping_lines]

            # 플레이스홀더 생성
            shipping_placeholders = ','.join(['?' for _ in standardized_shipping_lines])
            port_placeholders = ','.join(['?' for _ in destination_ports])
            
            # 파라미터 리스트 생성
            params = [standardized_origin] + standardized_shipping_lines + destination_ports + [start_date, end_date]

            for table_name in self.tab_info.values():
                query = f"""
                    SELECT origin, shippingdate, initialeta, etaport, unloadingterminal, eta, shippingline,
                           julianday(etaport) - julianday(initialeta) AS vessel_delay
                    FROM {table_name}
                    WHERE TRIM(origin) = ? 
                    AND TRIM(shippingline) IN ({shipping_placeholders})
                    AND TRIM(destinationport) IN ({port_placeholders})
                    AND DATE(etaport) BETWEEN ? AND ?
                """
                cursor.execute(query, params)
                rows = cursor.fetchall()
                if rows:
                    columns = [desc[0] for desc in cursor.description]
                    df = pd.DataFrame(rows, columns=columns)

                    # 데이터프레임의 origin과 shippingline 컬럼을 표준화
                    df['origin'] = df['origin'].apply(Mapping.standardize_origin_name)
                    df['shippingline'] = df['shippingline'].apply(Mapping.standardize_shipping_line_name)

                    data_frames.append(df)
            conn.close()

            # 빈 데이터프레임 처리
            data_frames = [df for df in data_frames if not df.empty and not df.isna().all().all()]

            if not data_frames:
                return "No data found for the selected conditions."

            data = pd.concat(data_frames, ignore_index=True)

            # 날짜 열 처리
            date_columns = ['shippingdate', 'initialeta', 'etaport', 'unloadingterminal', 'eta']
            for col in date_columns:
                if col in data.columns:
                    data[col] = pd.to_datetime(data[col], errors='coerce')

            # 유효한 날짜 데이터만 사용
            data = data.dropna(subset=date_columns)

            if data.empty:
                return "No valid date data found."

            # 분석 수행
            results = {}

            # 1. 평균 Vessel Delay
            if 'vessel_delay' in data.columns:
                vessel_delay_avg = data['vessel_delay'].mean()
                results['Average Vessel Delay'] = vessel_delay_avg
            else:
                results['Average Vessel Delay'] = 'N/A'

            # 2. shippingdate - etaport 평균 리드타임
            if 'shippingdate' in data.columns and 'etaport' in data.columns:
                data['leadtime1'] = (data['etaport'] - data['shippingdate']).dt.days
                leadtime1_avg = data['leadtime1'].mean()
                results['Average Lead Time (shippingdate to etaport)'] = leadtime1_avg
            else:
                results['Average Lead Time (shippingdate to etaport)'] = 'N/A'

            # 3. etaport - unloadingterminal 평균 리드타임
            if 'etaport' in data.columns and 'unloadingterminal' in data.columns:
                data['leadtime2'] = (data['unloadingterminal'] - data['etaport']).dt.days
                leadtime2_avg = data['leadtime2'].mean()
                results['Average Lead Time (etaport to unloadingterminal)'] = leadtime2_avg
            else:
                results['Average Lead Time (etaport to unloadingterminal)'] = 'N/A'

            # 4. unloadingterminal - eta 평균 리드타임
            if 'unloadingterminal' in data.columns and 'eta' in data.columns:
                data['leadtime3'] = (data['eta'] - data['unloadingterminal']).dt.days
                leadtime3_avg = data['leadtime3'].mean()
                results['Average Lead Time (unloadingterminal to eta)'] = leadtime3_avg
            else:
                results['Average Lead Time (unloadingterminal to eta)'] = 'N/A'

            # 5. shippingdate - eta 평균 리드타임
            if 'shippingdate' in data.columns and 'eta' in data.columns:
                data['leadtime4'] = (data['eta'] - data['shippingdate']).dt.days
                leadtime4_avg = data['leadtime4'].mean()
                results['Average Lead Time (shippingdate to eta)'] = leadtime4_avg
            else:
                results['Average Lead Time (shippingdate to eta)'] = 'N/A'

            return results

        except Exception as e:
            return f"An error occurred during analysis: {str(e)}"

    def show_origin_analysis(self, table_name, start_date, end_date):
        try:
            conn = connect_db()
            
            # 날짜 범위에 해당하는 데이터 조회 쿼리
            query = f"""
            SELECT [origin], [eta], COUNT(*) as count
            FROM {table_name}
            WHERE date([eta]) BETWEEN ? AND ?
            GROUP BY [origin]
            ORDER BY count DESC
            """
            df = pd.read_sql(query, conn, params=(start_date, end_date))
            
            # 전체 데이터 조회를 위한 쿼리 (검증용)
            verification_query = f"""
            SELECT [origin], [eta], [container], [division]
            FROM {table_name}
            WHERE date([eta]) BETWEEN ? AND ?
            ORDER BY [eta]
            """
            verification_df = pd.read_sql(verification_query, conn, params=(start_date, end_date))
            
            conn.close()

            if df.empty:
                QMessageBox.information(self, "Information", "No data for the selected period.")
                return

            # 차트 생성 (기존 코드)
            plt.figure(figsize=(10, 6))
            bars = plt.bar(df['origin'], df['count'])
            plt.title(f'Origin Analysis ({start_date} ~ {end_date})')
            plt.xlabel('Origin')
            plt.ylabel('Count')
            plt.xticks(rotation=45, ha='right')

            # 데이터 검증 버튼 추가
            verify_button = QPushButton("Verify Data", self)
            verify_button.clicked.connect(lambda: self.show_verification_data(verification_df, start_date, end_date))
            
            # 버튼을 차트 아래에 추가
            layout = QVBoxLayout()
            layout.addWidget(plt.gcf().canvas)
            layout.addWidget(verify_button)

            # 새 창에 표시
            dialog = QDialog(self)
            dialog.setWindowTitle("Origin Analysis")
            dialog.setLayout(layout)
            dialog.resize(800, 600)
            dialog.exec_()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")

    def show_verification_data(self, df, start_date, end_date):
        """
        선택된 날짜 범위의 상세 데이터를 표시하는 창을 엽니다.
        """
        try:
            if not df.empty:
                # DataDisplayWindow를 사용하여 데이터 표시
                window = DataDisplayWindow(df, self)
                window.setWindowTitle(f"Data Verification ({start_date} ~ {end_date})")
                window.setModal(False)
                window.show()
                
                # 날짜 범위 정보 추가
                info_label = QLabel(f"Selected Date Range: {start_date} ~ {end_date}")
                info_label.setStyleSheet("color: white;")
                window.layout().insertWidget(0, info_label)
                
                # 통계 ���보 추가
                stats_text = f"""
                Total Containers: {len(df)}
                Unique Origins: {df['origin'].nunique()}
                Date Range: {df['eta'].min()} ~ {df['eta'].max()}
                """
                stats_label = QLabel(stats_text)
                stats_label.setStyleSheet("color: white;")
                window.layout().insertWidget(1, stats_label)
            else:
                QMessageBox.information(self, "Information", "No data available for verification.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while showing verification data: {str(e)}")

    def show_origin_analysis_dialog(self):
        try:
            dialog = QDialog(self)
            dialog.setWindowTitle("Origin Analysis")
            layout = QVBoxLayout()

            # 날짜 선택 위젯
            date_layout = QHBoxLayout()
            
            # Start Date
            start_date_label = QLabel("Start Date (etaport):", dialog)
            self.start_date_edit = QDateEdit(dialog)
            self.start_date_edit.setCalendarPopup(True)
            self.start_date_edit.setDate(QDate.currentDate())
            date_layout.addWidget(start_date_label)
            date_layout.addWidget(self.start_date_edit)
            
            # End Date
            end_date_label = QLabel("End Date (etaport):", dialog)
            self.end_date_edit = QDateEdit(dialog)
            self.end_date_edit.setCalendarPopup(True)
            self.end_date_edit.setDate(QDate.currentDate())
            date_layout.addWidget(end_date_label)
            date_layout.addWidget(self.end_date_edit)
            
            layout.addLayout(date_layout)

            # Origins 리스트
            origins_label = QLabel("Origins:", dialog)
            layout.addWidget(origins_label)
            self.origins_list = QListWidget(dialog)
            self.origins_list.setSelectionMode(QAbstractItemView.MultiSelection)
            layout.addWidget(self.origins_list)

            # Shipping Lines 리스트
            shipping_lines_label = QLabel("Shipping Lines:", dialog)
            layout.addWidget(shipping_lines_label)
            self.shipping_lines_list = QListWidget(dialog)
            self.shipping_lines_list.setSelectionMode(QAbstractItemView.MultiSelection)
            layout.addWidget(self.shipping_lines_list)

            # 버튼을 위한 수평 레이아웃
            button_layout = QHBoxLayout()
            
            # Verify Data 버튼 추가
            verify_button = QPushButton("Verify Data", dialog)
            verify_button.clicked.connect(self.handle_verify_button_click)
            
            # Close 버튼
            close_button = QPushButton("Close", dialog)
            close_button.clicked.connect(dialog.close)
            
            # 버튼들을 수평 레이아웃에 추��
            button_layout.addWidget(verify_button)
            button_layout.addWidget(close_button)
            
            # 메인 레이아웃의 맨 아래에 버튼 레이아웃 추가
            layout.addLayout(button_layout)
            
            dialog.setLayout(layout)
            dialog.exec_()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")

    def handle_verify_button_click(self):
        """Verify 버튼 클릭 핸들러"""
        try:
            # 선택된 날짜 가져오기
            start_date = self.start_date_edit.date().toPyDate()
            end_date = self.end_date_edit.date().toPyDate()
            
            # 선택된 Origins 가져오기
            selected_origins = [item.text() for item in self.origins_list.selectedItems()]
            
            # 선택된 Shipping Lines 가져오기
            selected_shipping_lines = [item.text() for item in self.shipping_lines_list.selectedItems()]
            
            # 데이터 검증 함수 호출
            self.verify_origin_analysis_data(
                start_date,
                end_date,
                selected_origins,
                selected_shipping_lines
            )
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while handling verify button: {str(e)}")

    def verify_origin_analysis_data(self, start_date, end_date, selected_origins, selected_shipping_lines):
        try:
            if not selected_origins or not selected_shipping_lines:
                QMessageBox.warning(self, "Warning", "Please select at least one Origin and one Shipping Line.")
                return

            conn = connect_db()
            
            # 선택된 Origin과 Shipping Line을 표준화
            standardized_origins = [Mapping.standardize_origin_name(origin) for origin in selected_origins]
            standardized_shipping_lines = [Mapping.standardize_shipping_line_name(sl) for sl in selected_shipping_lines]

            # 선택된 조건에 맞는 데이터 조회
            query = f"""
            SELECT [origin], [etaport], [container], [division], [shippingline]
            FROM {self.current_table_name}
            WHERE date([etaport]) BETWEEN ? AND ?
            AND [origin] IN ({','.join(['?']*len(standardized_origins))})
            AND [shippingline] IN ({','.join(['?']*len(standardized_shipping_lines))})
            ORDER BY [etaport]
            """
            
            # 쿼리 파라미터 설정
            params = [start_date, end_date] + standardized_origins + standardized_shipping_lines
            
            # 데이터 조회
            df = pd.read_sql(query, conn, params=params)
            conn.close()

            if not df.empty:
                # 데이터 표시 창 생성
                window = DataDisplayWindow(df, self)
                window.setWindowTitle(f"Origin Analysis Data Verification ({start_date} ~ {end_date})")
                
                # 통계 정보 추가
                stats_text = f"""
                Date Range: {start_date} ~ {end_date}
                Total Containers: {len(df)}
                Unique Origins: {df['origin'].nunique()}
                Selected Origins: {', '.join(selected_origins)}
                Selected Shipping Lines: {', '.join(selected_shipping_lines)}
                """
                stats_label = QLabel(stats_text)
                stats_label.setStyleSheet("color: white;")
                window.layout().insertWidget(0, stats_label)
                
                window.setModal(False)
                window.show()
            else:
                QMessageBox.information(self, "Information", "No data available for the selected criteria.")
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while verifying data: {str(e)}")

    def calculate_storage_cost(self, df):
        
        # 현재 날짜 기준 설정
        current_date = datetime.now()
        
        # 데이터프레임의 날짜를 datetime으로 변환
        df['etaport'] = pd.to_datetime(df['etaport'])
        
        # 현재 연도 데이터와 미래 데이터(2025년 이후) 필터링
        current_df = df[df['etaport'].dt.year == current_date.year]
        future_df = df[df['etaport'].dt.year >= 2025]  # 2025년 이후 모든 데이터
        
        # 두 데이터프레임 합치기
        combined_df = pd.concat([current_df, future_df])
        
        if combined_df.empty:
            return pd.DataFrame()
        
        # 시작일: 현재 날짜 기준
        start_date = current_date - timedelta(days=30)
        
        # 종료일: 데이터의 가장 마지막 날짜 + 1년
        # (미래 데이터의 스토리지 비용을 1년치 더 계산하기 위함)
        max_future_date = future_df['etaport'].max() if not future_df.empty else current_date
        end_date = max_future_date + timedelta(days=365)
    
    def show_storage_cost_analysis(self, table_name):
        """스토리지 비용 분석 창을 표시"""
        analysis_window = StorageCostAnalysisWindow(table_name, self)
        analysis_window.show()
 

if __name__ == "__main__":
    app = QApplication(sys.argv)

    # qdarkstyle 적용
    app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())

    # 로딩 화면으로 시작
    loading_screen = LoadingScreen()
    loading_screen.show()

    sys.exit(app.exec_())