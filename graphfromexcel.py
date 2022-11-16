import openpyxl, random, os, warnings
import matplotlib.pyplot as plt
import pandas as pd
from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt, QRect
from PyQt5.QtGui import QStandardItemModel, QFont
import traceback
import numpy as np

import matplotlib
import matplotlib.font_manager as fm
font_path = "C:/Windows/Fonts/batang.ttc"
font_name = fm.FontProperties(fname=font_path).get_name()
matplotlib.rc('font',family=font_name)

class exceler:
    def __init__(self,
                 filename : str = 'test'):
        self.filename = filename

    def readxl(self):
        with warnings.catch_warnings(record=True):
            warnings.simplefilter("always")
            self.df = pd.read_excel(self.filename, engine="openpyxl")
        self.numlist = []
        
        for i in self.df:
            if i == '신고일 주':
                for j in range(len(self.df[i])):
                    self.numlist.append(int(float(self.df[i][j])))
        self.length = max(set(self.numlist))


    def setdata(self, numlimit):
        self.data = [len([0 for j in self.numlist if i==j]) for i in range(1, numlimit)]
        for i in range(1,len(self.filename)):
            if self.filename[len(self.filename)-i] == '/':
                self.address = len(self.filename)-i+1
                self.imagename = self.filename[len(self.filename)-i:-5]
                break

    def image_generate(self, data):
        plt.figure(figsize=(5.4, 5.8))
        for i in range(len(data)):
            if data[i] != 0:
                zero_del = i+1
        zero_del = np.array([True if j<zero_del else False for j in [i for i in range(len(data))]
                             ])
        
        plt.plot(data, zero_del, color='white', linewidth=2.2, label='신고자 수')
        plt.ylabel('(명)')
        plt.xlabel('주(week)')
        plt.legend()
        plt.savefig(self.filename[:self.address]+self.imagename+'.png')
        plt.cla()

import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QIcon
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

class PlotCanvas(FigureCanvas):

    def __init__(self, parent=None, width=5, height=4, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi)

        FigureCanvas.__init__(self, self.fig)
        self.setParent(parent)

        FigureCanvas.setSizePolicy(self,
                QSizePolicy.Expanding,
                QSizePolicy.Expanding)
        FigureCanvas.updateGeometry(self)
        self.plot()


    def plot(self, data=[], title=''):
        self.fig.clear()
        ax = self.figure.add_subplot(111)
        if data == []:
            ax.set_xlim(0, 50)
            ax.set_ylim(0, 40)
            #ax.set_xticks([0.04, 0.08, 0.12])
            ax.set_yticks([0, 10, 20, 30, 40])
        ax.set_ylabel('(명)')
        ax.set_xlabel('주(week)')
        ax.plot(data)
        ax.set_title(title)
        self.draw()



class MyApp(QWidget):

    def __init__(self):
        super().__init__()
        f = open(os.getcwd()+'\\readme1.txt', encoding='utf-8')
        self.icon_add = os.getcwd()+'\\icon\\xlsx.png'
        self.text = f.read()
        self.data = []
        self.initUI()

    def initUI(self):

        #사용 규칙 텍스트
        label2 = QLabel(self.text, self)
        label2.move(560, 360)


        #그래프 표시
        self.model1 = PlotCanvas(self)
        self.model1.move(10, 10)
        self.model1.resize(540,580)

        #표 길이 슬라이더
        label6 = QLabel('표 길이', self)
        label6.move(255, 630)
        
        self.slider1 = QSlider(Qt.Horizontal, self)
        self.slider1.move(300, 625)
        
        self.label5 = QLabel(str(self.slider1.value()), self)
        self.label5.move(390, 628)
        self.label5.resize(60, 15)
        min_1 = 50
        max_1 = 150
        self.slider1.setRange(min_1, max_1)
        self.slider1.valueChanged.connect(self.value_changed)
        self.slider1.setSingleStep(2)

        #파일 위치 관련
        self.qle = QLineEdit(self)
        self.qle.move(65, 600)
        self.qle.resize(300, 20)
        self.qle.setReadOnly(True)

        label1 = QLabel('파일 위치', self)
        label1.move(7,600)
        label1.resize(110, 20)
        
        btn1 = QPushButton('찾아보기(B)...',self)
        btn1.setShortcut('B')
        btn1.clicked.connect(self.openfile)
        btn1.move(370, 600)

        #통합 그래프용 리스트
        label7 = QLabel('표 통합하기', self)
        label7.move(561, 11)
        
        self.treev = QTreeWidget(self)
        self.treev.setColumnCount(3)
        
        colist = ["이름", "길이", "순서"]
        self.treev.setHeaderLabels(colist)
        self.treev.setAlternatingRowColors(True);
        self.treev.header().setSectionResizeMode(QHeaderView.Stretch)
        self.treev.move(560, 30)
        self.treev.resize(390, 255)

        btn3 = QPushButton('통합 그래프 이미지 생성(C)',self)
        btn3.setShortcut('C')
        btn3.clicked.connect(self.imagecreate)
        
        btn4 = QPushButton('선택한 항목 제거(X)',self)
        btn4.setShortcut('X')
        btn4.clicked.connect(self.delsel)

        btn5 = QPushButton('한 칸 아래로(↓)', self)
        btn5.setShortcut('Down')
        btn5.clicked.connect(self.treedown)

        btn6 = QPushButton('한 칸 위로(↑)', self)
        btn6.setShortcut('Up')
        btn6.clicked.connect(self.treeup)

        btn7 = QPushButton('통합 그래프 미리보기(V)',self)
        btn7.setShortcut('V')
        btn7.clicked.connect(self.graphview)
        
        btn3.resize(192, 30)
        btn4.resize(192, 30)
        btn5.resize(95, 30)
        btn6.resize(95, 30)
        btn7.resize(192, 30)
        
        btn3.move(560, 323)
        btn4.move(757, 323)
        btn5.move(757, 290)
        btn6.move(854, 290)
        btn7.move(560, 290)
        

        #진행 바(꾸미기용)
        self.label3 = QLabel('진행 상황', self)
        self.label3.move(7, 628)
        
        self.bar = QProgressBar(self)
        self.bar.setAlignment(Qt.AlignCenter)
        self.bar.setValue(0)
        self.bar.resize(110, 20)
        self.bar.move(65, 625)

        self.label4 = QLabel('대기 중', self)
        self.label4.move(180, 628)
        self.label4.setStyleSheet("Color : gray")

        #실행
        self.setWindowTitle('그래프생성기')
        self.setGeometry(500,300,960,653)
        self.setFixedSize(960,653)
        self.show()

    def treedown(self):
        try:
            sel = self.treev.selectedItems()[0]
            thenum = int(sel.data(2, 0))-1
            if thenum != len(self.data)-1:
                self.data[thenum], self.data[thenum+1] = self.data[thenum+1], self.data[thenum]
            self.treev_update()
        except Exception as e:
            pass

    def treeup(self):
        try:
            sel = self.treev.selectedItems()[0]
            thenum = int(sel.data(2, 0))-1
            if thenum != 0:
                self.data[thenum], self.data[thenum-1] = self.data[thenum-1], self.data[thenum]
            self.treev_update()
        except Exception as e:
            pass
        
    def delsel(self):
        try:
            sel = self.treev.selectedItems()[0]
            if type(sel)==type(QTreeWidgetItem()):
                seld = [sel.data(0,0), sel.data(1,0), sel.data(2, 0)]
            for i, j in enumerate(self.data):
                if j[0:2] == seld[0:2] and j[3] == seld[2]:
                    del self.data[i]
                    break
            self.treev_update()
            self.model1.fig.clear()
            self.qle.setText('')
            self.model1.plot()
            
        except Exception as e:
            print(e)
        
    def treev_add(self, name, length, data):
        for i in self.data:
            if i[0] == name:
                del self.data[-1]
                break
        self.data.append([name, str(length), data, None])
        self.treev_update()
            
    def treev_update(self):
        self.treev.clear()
        for i, j in enumerate(self.data):
            self.data[i][3] = str(i+1)
            self.item_ = QTreeWidgetItem(self.treev, [*j[0:2], j[3]])
            self.item_.setIcon(0, QIcon(self.icon_add))

    def treev_sum(self):
        sumdata = []
        for i in self.data:
            sumdata+=list(i)[2]
        return sumdata
    
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
            self.slider1.setValue(self.gened.length)
            
            self.gened.setdata(self.slider1.value()+2)
            plt_imagename = '통합 그래프 '+self.gened.imagename[1:]+'~'+self.gened.imagename[1:]
            self.model1.plot(data=self.gened.data, title=plt_imagename)
            
            self.treev_add(self.gened.imagename[1:], self.gened.length, self.gened.data)
            self.label4.setText('대기 중')
            self.label4.setStyleSheet('Color : gray')
            self.bar.setValue(0)
            
        except Exception as e:
            traceback.print_exc()
            self.label4.setText('오류')
            self.label4.setStyleSheet("Color : red")
            print(e)
            
    def graphview(self):
        try:
            self.model1.fig.clear()
            self.gened.data = self.treev_sum()
            self.gened.length = len(self.gened.data)
            min_1 = self.gened.length
            max_1 = self.gened.length*2

            proto_value = self.slider1.value()
            self.slider1.setRange(min_1, max_1)
            self.slider1.setValue(proto_value)
            
            self.gened.data+=[0 for i in range(int(self.slider1.value()-self.gened.length))]
            self.gened.imagename = '통합 그래프 '+str(self.data[0][0])+'~'+str(self.data[-1][0])
            for i in range(len(self.gened.data)):
                if self.gened.data[i] != 0:
                    zero_del = i+1
            zero_del = np.array([True if j<zero_del else False for j in [i for i in range(len(self.gened.data))]
                             ])
        
            self.model1.plot(self.gened.data, zero_del, color='white', linewidth=2.2, label='신고자 수')

            #self.model1.plot(data=self.gened.data, title=self.gened.imagename)
        except Exception as e:
            print(e)
    
    def imagecreate(self):
        try:
            plt.cla()
            self.label4.setText('진행 중')
            self.label4.setStyleSheet("Color : blue")
            
            sumdata = self.treev_sum()
            
            if sumdata != []:
                self.bar.setValue(25)
                plt.plot(sumdata, label='신고자 수')
                self.bar.setValue(50)
                self.gened.length=self.slider1.value()
                self.gened.setdata(self.slider1.value()+2)
                self.gened.imagename = '통합 그래프 '+str(self.data[0][0])+'~'+str(self.data[-1][0])
                self.bar.setValue(75)
                self.gened.image_generate(sumdata)
                self.label4.setText('완료')
                self.label4.setStyleSheet("Color : green")
                self.bar.setValue(100)
            else:
                raise Exception("데이터가 없습니다.")
        except Exception as e:
            print(e)
            self.label4.setText('오류')
            self.label4.setStyleSheet("Color : red")

            
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())

