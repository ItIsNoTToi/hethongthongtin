import os 
import pandas as pd 
import matplotlib.pyplot as plt

from PyQt5.QtWidgets import QDialog, QFileDialog, QTableWidgetItem, QMessageBox
from PyQt5 import uic, QtWidgets, QtChart
from qgis.core import QgsVectorLayer, QgsProject, QgsField
from PyQt5.QtCore import Qt 
from dbfread import DBF 
from xml.etree import ElementTree as ET
from functools import partial
import csv


FORM_CLASS2, _ = uic.loadUiType(os.path.join(
    os.path.dirname(__file__), 'linked.ui'))

class link_Dialog(QtWidgets.QDialog, FORM_CLASS2):
    def __init__(self, parent=None, linked_data=None):
        """Constructor cho giao diện kết quả."""
        super(link_Dialog, self).__init__(parent)
        self.setupUi(self)
        self.setMinimumSize(850, 600)
        
        self.statisticsFunct = ["Count", "Sum", "Mean"]
        for field in self.statisticsFunct: 
            self.cmb_HTK.addItem(field)

        # check linked_data
        if linked_data is None:
            raise ValueError("linked_data không được truyền hoặc bị None.")
        self.linked_data = linked_data
        self.statistics_data = linked_data

        self.set_statistics() # phai khoi tao self.linked_data truoc 
        #self.set_statistics_field()
        self.btn_TK.clicked.connect(self.statistics) 

        # connect buttons to switch to different page
        self.btn_percentageChart.clicked.connect(self.switch_to_PercentageChart)
        self.btn_donutChart.clicked.connect(self.switch_to_NestedChart)
        self.btn_barChart.clicked.connect(self.switch_to_BarChart)
        self.btn_export.clicked.connect(self.export_data)

        self.show_tableWidget(linked_data)

    def set_statistics(self):
        try: 
            self.cmb_TTTK.clear()
            self.cmb_group.clear()
            for col in self.linked_data.columns:
                self.cmb_TTTK.addItem(col)
                self.cmb_group.addItem(col)
        except ValueError:
            print('khong the doc cac truong du lieu')

    def get_statistics_field(self):
        try:
            return self.cmb_TTTK.currentText()
        except ValueError:
            return ''

    def get_statistics_funct(self):
        try:
            return self.cmb_HTK.currentText()
        except ValueError:
            return ''
        
    def get_group_field(self):
        try:
            return self.cmb_group.currentText()
        except ValueError:
            return ''

    def sum_statistics(self, data):
        """
        Tính tổng các giá trị trong cột được chỉ định, nhóm theo cột khóa.
        """
        group_field = self.get_group_field()  # Cột dùng làm khóa nhóm
        statistics_field = self.get_statistics_field()  # Cột cần tính tổng

        if group_field not in data.columns or statistics_field not in data.columns:
            raise ValueError(f"Cột '{group_field}' hoặc '{statistics_field}' không tồn tại trong dữ liệu.")
        
        # Kiểm tra nếu cột cần tính tổng không phải là số 
        if not pd.api.types.is_numeric_dtype(data[statistics_field]): 
            raise ValueError(f"Cột '{statistics_field}' phải là kiểu số để tính tổng.")

        grouped_data = data.groupby(group_field)[statistics_field].sum().reset_index()
        return grouped_data

    def count_statistics(self, data):
        """
        Đếm số lượng giá trị trong cột được chỉ định, nhóm theo cột khóa.
        """
        group_field = self.get_group_field()  # Cột dùng làm khóa nhóm
        statistics_field = self.get_statistics_field()  # Cột cần đếm số lượng

        if group_field not in data.columns or statistics_field not in data.columns:
            raise ValueError(f"Cột '{group_field}' hoặc '{statistics_field}' không tồn tại trong dữ liệu.")

        grouped_data = data.groupby(group_field)[statistics_field].count().reset_index()
        return grouped_data

    def mean_statistics(self, data):
        """Tính giá trị trung bình trong cột được chỉ định, nhóm theo cột khóa."""
        group_field = self.get_group_field()  # Cột dùng làm khóa nhóm
        statistics_field = self.get_statistics_field()  # Cột cần tính trung bình

        if group_field not in data.columns or statistics_field not in data.columns:
            raise ValueError(f"Cột '{group_field}' hoặc '{statistics_field}' không tồn tại trong dữ liệu.")
        
        # Kiểm tra nếu cột cần tính tích không phải là số 
        if not pd.api.types.is_numeric_dtype(data[statistics_field]): 
            raise ValueError(f"Cột '{statistics_field}' phải là kiểu số để tính tích.")

        grouped_data = data.groupby(group_field)[statistics_field].mean().reset_index()
        grouped_data.columns = ['Category', 'Value'] # Đổi tên cột
        return grouped_data

    def show_tableWidget(self, data): 
        if data is not None and not data.empty:
            self.tb_DLKQ.clear()
            self.tb_DLKQ.setRowCount(0)
            self.tb_DLKQ.setColumnCount(len(data.columns))
            self.tb_DLKQ.setHorizontalHeaderLabels(data.columns)

            for index, row in data.iterrows():
                row_position = self.tb_DLKQ.rowCount()
                self.tb_DLKQ.insertRow(row_position)
                for col, value in enumerate(row):
                    self.tb_DLKQ.setItem(row_position, col, QTableWidgetItem(str(value)))

    def switch_to_PercentageChart(self):
        self.stackedWidget.setCurrentIndex(0)

    def switch_to_NestedChart(self):
        self.stackedWidget.setCurrentIndex(1)

    def switch_to_BarChart(self):
        self.stackedWidget.setCurrentIndex(2)


    def statistics(self):
        try:
            statistics_func = self.get_statistics_field()
            data = pd.DataFrame()
            
            # Tính toán theo hàm thống kê đã chọn
            if statistics_func == "Count":
                data = self.count_statistics(self.linked_data)
            elif statistics_func == "Sum":
                data = self.sum_statistics(self.linked_data)
            elif statistics_func == "Mean":
                data = self.mean_statistics(self.linked_data)
            else:
                QMessageBox.warning(self, "Lỗi", "Hàm thống kê không hợp lệ")
                return

            # Kiểm tra và hiển thị các biểu đồ
            if data is not None and not data.empty:
                self.statistics_data = data
                self.percentageBarChart(data)
                self.nestedDonutChart(data)
                self.barChart(data)
            else:
                QMessageBox.warning(self, "Lỗi", "Tính dữ liệu thống kê thất bại")
        except ValueError as e:
            QMessageBox.warning(self, "Lỗi", str(e))
   

    def export_data(self):
        self.export_to_excel(self.linked_data, "Dữ liệu liệu liên kết")
        self.export_to_excel(self.statistics_data, "Dữ liệu thống kê")

    def export_to_excel(self, data, fileName = "Excel"):
        """
        Xuất dữ liệu DataFrame sang tệp Excel, cho phép người dùng chọn nơi lưu.

        Parameters:
            data: DataFrame
                Dữ liệu cần xuất.
        """
        if data is not None and not data.empty:
            try:
                options = QFileDialog.Options()
                file_path, _ = QFileDialog.getSaveFileName(
                    self, 
                    "Chọn nơi lưu tệp " + fileName, 
                    "", 
                    "Excel Files (*.xlsx);;All Files (*)", 
                    options=options
                )
                if file_path:
                    data.to_excel(file_path, index=False)
                    QMessageBox.information(self, "Thành công", f"Dữ liệu đã được xuất ra tệp {file_path}")
                else:
                    QMessageBox.warning(self, "Lỗi", "Đã huỷ chọn nơi lưu tệp")
            except Exception as e:
                QMessageBox.warning(self, "Lỗi", f"Xuất dữ liệu thất bại: {str(e)}")
        else:
            QMessageBox.warning(self, "Lỗi", "Không có dữ liệu để xuất")


    def percentageBarChart(self, grouped_data):
        """
        Tạo biểu đồ cột phần trăm dựa trên dữ liệu thống kê.

        Parameters:
            grouped_data: DataFrame
                Dữ liệu đã được nhóm với các cột 'Category' và 'Value'.
                Ví dụ: {"Category": ["Python", "Java", "C#", "Javascript"], "Value": [12, 10, 34, 90]}.
        """
        # Khởi tạo biểu đồ
        self.percentage_bar_chart = QtChart.QChart()

        # Tạo QBarSeries
        self.percentage_bar_series = QtChart.QBarSeries()

        # Tổng giá trị để tính tỷ lệ phần trăm
        total_value = grouped_data["Value"].sum()

        # Tạo QBarSet cho từng loại dữ liệu với tỷ lệ phần trăm
        bar_set = QtChart.QBarSet("Percentage")
        for index, row in grouped_data.iterrows():
            percentage = (row["Value"] / total_value) * 100 if total_value > 0 else 0
            bar_set.append(percentage)

        self.percentage_bar_series.append(bar_set)

        # Thêm series vào chart
        self.percentage_bar_chart.addSeries(self.percentage_bar_series)
        self.percentage_bar_chart.setTitle("Biểu đồ cột phần trăm")
        self.percentage_bar_chart.setAnimationOptions(QtChart.QChart.SeriesAnimations)

        # Tạo trục x và y
        categories = grouped_data["Category"].tolist()
        axisX = QtChart.QBarCategoryAxis()
        axisX.append(categories)
        self.percentage_bar_chart.addAxis(axisX, Qt.AlignBottom)
        self.percentage_bar_series.attachAxis(axisX)

        axisY = QtChart.QValueAxis()
        axisY.setRange(0, 100)  # Phần trăm từ 0 đến 100
        self.percentage_bar_chart.addAxis(axisY, Qt.AlignLeft)
        self.percentage_bar_series.attachAxis(axisY)

        # Gắn biểu đồ vào giao diện
        self.percentageBarChart_view.setChart(self.percentage_bar_chart)

    def nestedDonutChart(self, grouped_data):
        """
        Tạo biểu đồ hình tròn lồng nhau dựa trên dữ liệu thống kê.

        Parameters:
            grouped_data: DataFrame
                Dữ liệu đã được nhóm với các cột 'Category' và 'Value'.
                Ví dụ: {"Category": ["Python", "Java", "C#", "Javascript"], "Value": [12, 10, 34, 90]}.
        """
        # Khởi tạo biểu đồ
        self.pie_chart = QtChart.QChart()

        # Tạo QPieSeries
        self.pie_series = QtChart.QPieSeries()
        self.pie_series.setLabelsVisible(True)
        self.pie_series.setLabelsPosition(QtChart.QPieSlice.LabelOutside)

        # Tổng giá trị để tính tỷ lệ phần trăm
        total_value = grouped_data["Value"].sum()

        # Thêm các lát cắt vào biểu đồ
        for index, row in grouped_data.iterrows():
            label = row["Category"]
            value = row["Value"]
            percentage = (value / total_value) * 100 if total_value > 0 else 0
            slice_item = self.pie_series.append(f"{label} ({percentage:.2f}%)", value)

            # Tùy chỉnh lát cắt
            if isinstance(slice_item, QtChart.QPieSlice):
                slice_item.setLabelVisible(True)
                slice_item.setLabelPosition(QtChart.QPieSlice.LabelOutside)

        # Thêm series vào chart
        self.pie_chart.addSeries(self.pie_series)
        self.pie_chart.setTitle("Biểu đồ hình tròn")
        self.pie_chart.legend().setVisible(True)
        self.pie_chart.legend().setAlignment(Qt.AlignBottom)

        # Gắn biểu đồ vào giao diện
        self.nestedChart_view.setChart(self.pie_chart)

    def barChart(self, grouped_data):
        """
        Tạo biểu đồ cột dựa trên dữ liệu thống kê.

        Parameters:
            grouped_data: DataFrame
                Dữ liệu đã được nhóm với các cột 'Category' và 'Value'.
                Ví dụ: {"Category": ["Python", "Java", "C#", "Javascript"], "Value": [12, 10, 34, 90]}.
        """
        # Khởi tạo biểu đồ
        self.bar_chart = QtChart.QChart()

        # Tạo QBarSeries
        self.bar_series = QtChart.QBarSeries()

        # Tạo QBarSet cho từng loại dữ liệu
        bar_set = QtChart.QBarSet("Value")
        for index, row in grouped_data.iterrows():
            bar_set.append(row["Value"])

        self.bar_series.append(bar_set)

        # Thêm series vào chart
        self.bar_chart.addSeries(self.bar_series)
        self.bar_chart.setTitle("Biểu đồ cột")
        self.bar_chart.setAnimationOptions(QtChart.QChart.SeriesAnimations)

        # Tạo trục x và y
        categories = grouped_data["Category"].tolist()
        axisX = QtChart.QBarCategoryAxis()
        axisX.append(categories)
        self.bar_chart.addAxis(axisX, Qt.AlignBottom)
        self.bar_series.attachAxis(axisX)

        axisY = QtChart.QValueAxis()
        self.bar_chart.addAxis(axisY, Qt.AlignLeft)
        self.bar_series.attachAxis(axisY)

        # Gắn biểu đồ vào giao diện
        self.barChart_view.setChart(self.bar_chart)
