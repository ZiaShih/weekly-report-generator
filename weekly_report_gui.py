import sys
import os
import pandas as pd
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QPushButton, QLabel, QFileDialog, 
                            QMessageBox, QTableWidget, QTableWidgetItem, 
                            QHeaderView, QTextEdit, QComboBox, QLineEdit)
from PyQt6.QtCore import Qt
from weekly_report_generator import WeeklyReportGenerator

class WeeklyReportGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("周报生成器")
        self.setMinimumSize(800, 600)
        
        # 创建主窗口部件
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # 创建主布局
        layout = QVBoxLayout()
        main_widget.setLayout(layout)
        
        # 创建文件选择区域
        file_layout = QHBoxLayout()
        self.excel_path = QLineEdit()
        self.excel_path.setPlaceholderText("请选择Excel文件...")
        self.excel_path.setReadOnly(True)
        browse_btn = QPushButton("浏览...")
        browse_btn.clicked.connect(self.browse_excel)
        file_layout.addWidget(QLabel("Excel文件:"))
        file_layout.addWidget(self.excel_path)
        file_layout.addWidget(browse_btn)
        layout.addLayout(file_layout)
        
        # 创建期数和日期输入区域
        info_layout = QHBoxLayout()
        self.issue_input = QLineEdit()
        self.issue_input.setPlaceholderText("如：1")
        self.date_input = QLineEdit()
        self.date_input.setPlaceholderText("如：2025年1月7日")
        info_layout.addWidget(QLabel("期数:"))
        info_layout.addWidget(self.issue_input)
        info_layout.addWidget(QLabel("日期:"))
        info_layout.addWidget(self.date_input)
        layout.addLayout(info_layout)
        
        # 创建数据表格
        self.table = QTableWidget()
        self.table.setColumnCount(11)
        self.table.setHorizontalHeaderLabels([
            "工号", "姓名", "工作类型", "项目名称/入池机构-项目名称", 
            "项目阶段", "上周三至本周二工作内容", "本周三至下周二工作计划",
            "问题反馈", "通过简历数量", "面试人员数量", "面试通过人员数量"
        ])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.table)
        
        # 创建按钮区域
        button_layout = QHBoxLayout()
        load_btn = QPushButton("加载数据")
        load_btn.clicked.connect(self.load_data)
        pdf_btn = QPushButton("下载 PDF 文件")
        pdf_btn.clicked.connect(self.download_pdf)
        word_btn = QPushButton("下载 Word 文件")
        word_btn.clicked.connect(self.download_word)
        button_layout.addWidget(load_btn)
        button_layout.addWidget(pdf_btn)
        button_layout.addWidget(word_btn)
        layout.addLayout(button_layout)
        
        # 创建状态栏
        self.statusBar().showMessage("就绪")
    
    def browse_excel(self):
        """浏览并选择Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Excel文件",
            "",
            "Excel文件 (*.xlsx *.xls)"
        )
        if file_path:
            self.excel_path.setText(file_path)
    
    def load_data(self):
        """加载Excel数据到表格"""
        excel_path = self.excel_path.text()
        if not excel_path:
            QMessageBox.warning(self, "警告", "请先选择Excel文件")
            return
        try:
            # 读取Excel数据
            df = pd.read_excel(excel_path)
            # 动态设置表头
            self.table.setColumnCount(len(df.columns))
            self.table.setHorizontalHeaderLabels([str(col) for col in df.columns])
            # 设置表格行数
            self.table.setRowCount(len(df))
            # 填充表格数据
            for i, row in df.iterrows():
                for j, value in enumerate(row):
                    item = QTableWidgetItem(str(value))
                    self.table.setItem(i, j, item)
            self.statusBar().showMessage("数据加载成功")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载数据失败: {str(e)}")
    
    def download_pdf(self):
        """下载PDF文件"""
        excel_path = self.excel_path.text()
        issue = self.issue_input.text()
        date_str = self.date_input.text()
        if not excel_path:
            QMessageBox.warning(self, "警告", "请先选择Excel文件")
            return
        if not issue or not date_str:
            QMessageBox.warning(self, "警告", "请填写期数和日期")
            return
        try:
            save_path, _ = QFileDialog.getSaveFileName(
                self,
                "保存周报PDF",
                "",
                "PDF文件 (*.pdf)"
            )
            if save_path:
                generator = WeeklyReportGenerator(excel_path, save_path, issue, date_str)
                generator.run()
                QMessageBox.information(self, "成功", "PDF文件下载成功！")
                self.statusBar().showMessage("PDF文件下载成功")
                os.system(f"open '{save_path}'")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"下载PDF失败: {str(e)}")

    def download_word(self):
        """下载Word文件"""
        excel_path = self.excel_path.text()
        issue = self.issue_input.text()
        date_str = self.date_input.text()
        if not excel_path:
            QMessageBox.warning(self, "警告", "请先选择Excel文件")
            return
        if not issue or not date_str:
            QMessageBox.warning(self, "警告", "请填写期数和日期")
            return
        try:
            save_path, _ = QFileDialog.getSaveFileName(
                self,
                "保存周报Word",
                "",
                "Word文件 (*.docx)"
            )
            if save_path:
                from weekly_report_generator import generate_word_report
                generate_word_report(excel_path, save_path, issue, date_str)
                QMessageBox.information(self, "成功", "Word文件下载成功！")
                self.statusBar().showMessage("Word文件下载成功")
                os.system(f"open '{save_path}'")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"下载Word失败: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = WeeklyReportGUI()
    window.show()
    sys.exit(app.exec()) 