import openpyxl, random, os, warnings
import matplotlib.pyplot as plt
import pandas as pd
from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt

from matplotlib import font_manager,rc
import matplotlib
font_path = "C:/Windows/Fonts/HYHWPEQ.TTF"
font_name = font_manager.FontProperties(fname=font_path).get_name()
matplotlib.rc('font',family=font_name)

class exceler:
    def __init__(self,
                 filename : str = 'test',
                 numlimit : int = 52):
        self.filename = filename
        self.numlimit = numlimit

    def readxl(self):
        with warnings.catch_warnings(record=True):
            warnings.simplefilter("always")
            self.df = pd.read_excel(self.filename, engine="openpyxl")
        numlist = []
        
        for i in self.df:
            if i == '신고일 주':
                for j in range(len(self.df[i])):
                    try:
                        numlist.append(int(float(self.df[i][j])))
                    except Exception as e:
                        pass
        
        self.data = [len([0 for j in numlist if i==j]) for i in range(1, self.numlimit)]
        self.length = len(self.data)
        plt.plot(self.data, label='신고자 수')

    def image_generate(self):
        for i in range(1,len(self.filename)):
            if self.filename[len(self.filename)-i] == '/':
                address = len(self.filename)-i+1
                imagename = self.filename[len(self.filename)-i:-5]
                break
        plt.ylabel('(명)')
        plt.xlabel('주(week)')
        plt.legend()
        plt.savefig(self.filename[:address]+imagename+'.png')
        plt.cla()

import sys
from PyQt5.QtWidgets import *


class MyApp(QWidget):

    def __init__(self):
        super().__init__()
        f = open(os.getcwd()+'\\readme.txt', encoding='utf-8')
        self.text = f.read()
        self.initUI()

    def initUI(self):
        label1 = QLabel('파일 위치', self)
        label2 = QLabel(self.text, self)
        label3 = QLabel('진행 상황', self)
        self.label4 = QLabel('대기 중', self)
        label6 = QLabel('표 길이', self)
        
        self.slider1 = QSlider(Qt.Horizontal, self)
        self.label5 = QLabel(str(self.slider1.value()), self)
        self.slider1.move(300, 30)
        min_1 = 50
        max_1 = 150
        self.slider1.setRange(min_1, max_1)
        self.slider1.valueChanged.connect(self.value_changed)
        self.slider1.setSingleStep(2)
        
        label1.move(7,6)
        label1.resize(110, 20)
        label2.move(7, 60)
        label3.move(7, 35)
        self.label4.move(180, 35)
        self.label4.setStyleSheet("Color : gray")
        self.label5.move(390, 34)
        self.label5.resize(60, 15)
        label6.move(255, 34)
        
        self.qle = QLineEdit(self)
        self.qle.move(65, 6)
        self.qle.resize(300, 20)
        self.qle.setReadOnly(True)

        btn1 = QPushButton('찾아보기(B)...',self)
        btn1.setShortcut('B')
        btn1.clicked.connect(self.openfile)
        btn2 = QPushButton('그래프 생성(G)',self)
        btn2.setShortcut('G')
        btn2.clicked.connect(self.imagecreate)

        btn1.move(370, 5)
        btn2.move(460, 5)

        self.bar = QProgressBar(self)
        self.bar.setAlignment(Qt.AlignCenter)
        self.bar.setValue(0)
        self.bar.resize(110, 20)
        self.bar.move(65, 30)

        self.setWindowTitle('그래프생성기')
        self.setGeometry(500,500,560,193)
        self.show()

    def value_changed(self, value):
        try:
            self.label5.setText(str(value))
        except Exception as e:
            print(e)

    def openfile(self):
        try:
            plt.cla()
            self.filename = QFileDialog.getOpenFileName()
            self.qle.setText(self.filename[0])
            self.gened = exceler(filename=self.filename[0])
            self.gened.readxl()
            min_1 = self.gened.length
            max_1 = self.gened.length*2
            self.slider1.setRange(min_1, max_1)
        except Exception as e:
            self.label4.setText('오류')
            self.label4.setStyleSheet("Color : red")
    def imagecreate(self):
        try:
            self.gened.numlimit = self.slider1.value()
            self.label4.setText('진행 중')
            self.label4.setStyleSheet("Color : blue")
            self.bar.setValue(50)
            self.gened.image_generate()
            self.bar.setValue(75)
            self.label4.setText('완료')
            self.label4.setStyleSheet("Color : green")
            self.bar.setValue(100)
        except Exception as e:
            self.label4.setText('오류')
            self.label4.setStyleSheet("Color : red")
            
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())
