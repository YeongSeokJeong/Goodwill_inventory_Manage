import sys
from PyQt5.QtWidgets import *
import pandas as pd
import numpy as np
from PyQt5 import QtCore


class MyWindow(QWidget):
    
    def __init__(self):
        super().__init__()
        self.setupUI()
        self.input_file = ''
        self.output_file = ''

    def setupUI(self):
        self.resize(506,207)
        self.setWindowTitle("Inventory Management")

        self.pushButton1 = QPushButton("Input File Open")
        self.pushButton1.setGeometry(QtCore.QRect(30, 40, 151, 31))
        self.pushButton1.clicked.connect(self.pushButtonClicked1)
        
        self.textbox1 = QLineEdit(self)
        self.textbox1.setReadOnly(True)
        
        self.pushButton2 = QPushButton("Output File Open")
        self.pushButton2.setGeometry(QtCore.QRect(30, 90, 151, 31))
        self.pushButton2.setObjectName("pushButton2")
        self.pushButton2.clicked.connect(self.pushButtonClicked2)
        
        self.textbox2 = QLineEdit(self)
        self.textbox2.setReadOnly(True)
        
        self.pushButton3 = QPushButton("Do it!")
        self.pushButton3.clicked.connect(self.checked_file)
        self.label = QLabel()
        
#         self.push()
        
        layout = QGridLayout()
        layout.addWidget(self.pushButton1,0,0)
        layout.addWidget(self.textbox1,0,1)
        layout.addWidget(self.pushButton2,1,0)
        layout.addWidget(self.textbox2,1,1)
        layout.addWidget(self.pushButton3,2,0)
        layout.addWidget(self.label,2,1)

        self.setLayout(layout)

    def pushButtonClicked1(self):
        fname = QFileDialog.getOpenFileName(self)
        self.input_file = fname[0]
        self.textbox1.setText(fname[0])
        
    def pushButtonClicked2(self):
        fname = QFileDialog.getOpenFileName(self)
        self.output_file = fname[0]
        self.textbox2.setText(fname[0])
        
    def pushButtonClicked4(self):
        sys.exit()
        
    def checked_file(self):
        if self.input_file == '' and self.output_file == '':
            QMessageBox.warning(self,'Warning!',
                                 "Please check input,output file",
                               QMessageBox.Ok,QMessageBox.Ok)
            return
        elif self.input_file == '':
            QMessageBox.warning(self,'Warning!',
                                 "Please check input file",
                               QMessageBox.Ok,QMessageBox.Ok)
            return
        elif self.output_file == '':
            QMessageBox.warning(self,'Warning!',
                                 "Please check output file",
                               QMessageBox.Ok,QMessageBox.Ok)
            return
        self.arrange_inventory()
        
    def arrange_inventory(self):
        input_file = pd.read_excel(self.input_file,
                           sheet_name="Transferrred Stock")
        touch_data = input_file.iloc[3:].values
        title = input_file.iloc[2].to_list()
        st_name_qty = []
        for i in range(len(touch_data)):
            item = [touch_data[i,0].lower(),touch_data[i,1]]
            st_name_qty.append(item)
        output_file = pd.read_excel(self.output_file,
                           sheet_name = "Inventory List")
        data = output_file.iloc[2:,1:-1]
        columns = output_file.iloc[1,1:-1].to_list()
        data = data.astype({"Unnamed: 6":float,
                           "Unnamed: 7":float})
        item_name = data["Unnamed: 3"].to_list()
        item_name = [name.lower() for name in item_name if type(name) != float]
        for i in range(len(st_name_qty)):
            try:
                num = item_name.index(st_name_qty[i][0])
            except:
                QMessageBox.warning(self,'Warning!',
                                 "Please check " + st_name_qty[i][0],
                               QMessageBox.Ok,QMessageBox.Ok)
            data.iloc[num,5] -= st_name_qty[i][1]
        data.columns = columns
        display(data)
        dir_add = QFileDialog.getExistingDirectory(self,"Select Directory")
        self.label.setText("Success!")
        data.to_excel(dir_add + "/new_data.xlsx")
        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec_()