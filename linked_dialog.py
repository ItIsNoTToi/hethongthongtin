import os 
import pandas as pd 
import matplotlib.pyplot as plt

from PyQt5.QtWidgets import QDialog, QFileDialog, QTableWidgetItem, QMessageBox 
from PyQt5 import uic, QtWidgets 
from qgis.core import QgsVectorLayer, QgsProject, QgsField
from PyQt5.QtCore import Qt 
from dbfread import DBF 
from xml.etree import ElementTree as ET

FORM_CLASS2, _ = uic.loadUiType(os.path.join(
    os.path.dirname(__file__), 'linked.ui'))

class link_Dialog(QtWidgets.QDialog, FORM_CLASS2):
    def __init__(self, parent=None, matched_data=None):
        """Constructor cho giao diện kết quả."""
        super(link_Dialog, self).__init__(parent)
        self.setupUi(self)

        self.matched_data = matched_data
        self.show_tableWidget()

    def show_tableWidget(self): 
        if self.matched_data is not None and not self.matched_data.empty:
            self.tableWidget_3.setRowCount(0)
            self.tableWidget_3.setColumnCount(len(self.matched_data.columns))
            self.tableWidget_3.setHorizontalHeaderLabels(self.matched_data.columns)

            for index, row in self.matched_data.iterrows():
                row_position = self.tableWidget_3.rowCount()
                self.tableWidget_3.insertRow(row_position)
                for col, value in enumerate(row):
                    self.tableWidget_3.setItem(row_position, col, QTableWidgetItem(str(value)))
