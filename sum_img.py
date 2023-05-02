import sys
from PyQt5.QtWidgets import *
import os
from PyQt5.QtCore import *
import datetime
from PyQt5 import QtWidgets
import cv2
import numpy as np
import time
from imutils.perspective import four_point_transform
from imutils.contours import sort_contours
import matplotlib.pyplot as plt
import pytesseract

today_data = datetime.datetime.now()
f_path =os.getcwd()+'/'+today_data.strftime("%Y%m")

class QListWidget(QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent=None)
        self.setAcceptDrops(True)
        

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
            imgFiles = []
            for url in event.mimeData().urls():
                if str(url).find('.jpg')>0:
                    if url.isLocalFile():                   
                        imgFiles.append(str(url.toLocalFile()))
                    else :
                        imgFiles.append(str(url.toString()))
                if str(url).find('.tif')>0:
                    if url.isLocalFile():                   
                        imgFiles.append(str(url.toLocalFile()))
                    else :
                        imgFiles.append(str(url.toString()))
                if str(url).find('.pjp')>0:
                    if url.isLocalFile():                   
                        imgFiles.append(str(url.toLocalFile()))
                    else :
                        imgFiles.append(str(url.toString()))  
                if str(url).find('.xbm')>0:
                    if url.isLocalFile():                   
                        imgFiles.append(str(url.toLocalFile()))
                    else :
                        imgFiles.append(str(url.toString()))      
                if str(url).find('.jxl')>0:
                    if url.isLocalFile():                   
                        imgFiles.append(str(url.toLocalFile()))
                    else :
                        imgFiles.append(str(url.toString()))
                if str(url).find('.svgz')>0:
                    if url.isLocalFile():                   
                        imgFiles.append(str(url.toLocalFile()))
                    else :
                        imgFiles.append(str(url.toString()))
                if str(url).find('.jpeg')>0:
                    if url.isLocalFile():                   
                        imgFiles.append(str(url.toLocalFile()))
                    else :
                        imgFiles.append(str(url.toString()))
                if str(url).find('.ico')>0:
                    if url.isLocalFile():                   
                        imgFiles.append(str(url.toLocalFile()))
                    else :
                        imgFiles.append(str(url.toString()))
                if str(url).find('.tiff')>0:
                    if url.isLocalFile():                   
                        imgFiles.append(str(url.toLocalFile()))
                    else :
                        imgFiles.append(str(url.toString()))
                if str(url).find('.gif')>0:
                    if url.isLocalFile():                   
                        imgFiles.append(str(url.toLocalFile()))
                    else :
                        imgFiles.append(str(url.toString()))
                if str(url).find('.svg')>0:
                    if url.isLocalFile():                   
                        imgFiles.append(str(url.toLocalFile()))
                    else :
                        imgFiles.append(str(url.toString()))
                if str(url).find('.jfif')>0:
                    if url.isLocalFile():                   
                        imgFiles.append(str(url.toLocalFile()))
                    else :
                        imgFiles.append(str(url.toString()))
                if str(url).find('.webp')>0:
                    if url.isLocalFile():                   
                        imgFiles.append(str(url.toLocalFile()))
                    else :
                        imgFiles.append(str(url.toString()))
                if str(url).find('.png')>0:
                    if url.isLocalFile():                   
                        imgFiles.append(str(url.toLocalFile()))
                    else :
                        imgFiles.append(str(url.toString()))
                if str(url).find('.bmp')>0:
                    if url.isLocalFile():                   
                        imgFiles.append(str(url.toLocalFile()))
                    else :
                        imgFiles.append(str(url.toString()))
                if str(url).find('.pjpeg')>0:
                    if url.isLocalFile():                   
                        imgFiles.append(str(url.toLocalFile()))
                    else :
                        imgFiles.append(str(url.toString()))
                if str(url).find('.avif')>0:
                    if url.isLocalFile():                   
                        imgFiles.append(str(url.toLocalFile()))
                    else :
                        imgFiles.append(str(url.toString()))
            self.addItems(imgFiles)
        else:
            super().dropEvent(event)

   

class QPushButton(QPushButton):
    def __init__(self, parent = None):
        super().__init__(parent)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

class QLabel(QLabel):
    def __init__(self, parent = None):
        super().__init__(parent)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

class QLineEdit(QLineEdit):
    def __init__(self, parent = None):
        super().__init__(parent)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        #layout_2.addWidget(layout_1)

class QComboBox(QComboBox):
    def __init__(self, parent = None):
        super().__init__(parent)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

class StWidgetForm(QGroupBox):
    """
    위젯 베이스 클래스
    """
    def __init__(self):
        QGroupBox.__init__(self)
        self.box = QBoxLayout(QBoxLayout.TopToBottom)
        self.setLayout(self.box)

class Therad1(QThread):
    def __init__(self, parent) :
        super().__init__(parent)
        self.parent = parent

    def run(self):
        self.bar_count= int((self.parent.e/(self.parent.count-1))*100)
        self.parent.bar.setValue(self.bar_count)

class Therad2(QThread):
    def __init__(self, parent) :
        super().__init__(parent)
        self.parent = parent
    
    def run(self):
        cv2.setUseOptimized(False)
        self.count = self.parent.list_w.count()
        os.makedirs(f_path, exist_ok=True)
        m = today_data.strftime('%m월')
        name = self.parent.name_widget.text()
        self.bar_count = 0
        
        class_h = self.parent.class_widget.currentText()
        base_x =30
        base_y =30
        base_x_t = 30
        if self.parent.checkbox_value == 'true':
            base_img = np.zeros((2480,3508,3),np.uint8)
        else :
            base_img = np.zeros((3508,2480,3),np.uint8)
            #base_img_[i] = np.zeros((3508,2480,3),np.uint8)
        num = 1
        if self.count < 2 :
            self.mes = 2
            self.parent.stk_w.setCurrentIndex(2)
            #return self.mes
            #QMessageBox.information(self,"확인","이미지를 확인해주세요")
        else :
            self.parent.stk_w.setCurrentIndex(0)
            for self.e in range(0,self.count) :
                item = self.parent.list_w.item(self.e).text()

                image_gray = np.fromfile(item,np.uint8)
                img_gray_ = cv2.imdecode(image_gray, cv2.IMREAD_GRAYSCALE)
               
                img_gray=cv2.resize(img_gray_,dsize=(0,0), fx=0.5,fy=0.5,interpolation=cv2.INTER_AREA)
                img_ = cv2.imdecode(image_gray, cv2.IMREAD_COLOR)
                img=cv2.resize(img_,dsize=(0,0), fx=0.5,fy=0.5,interpolation=cv2.INTER_AREA)
              
                blur = cv2.GaussianBlur(img_gray, ksize=(5,5), sigmaX=0)
                
                edged = cv2.Canny(blur, 10, 250)
                kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (7,7))
                closed = cv2.morphologyEx(edged, cv2.MORPH_CLOSE, kernel)
                contours, _ = cv2.findContours(closed.copy(),cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
                
                contours_xy = np.array(contours)
                contours_xy1 = np.array(contours)
                contours_xy.shape
                time.sleep(0.5)
                x_min, x_max = 0,0
                y_min, y_max = 0,0
                value = list()
                value1 = list()
                for i in range(len(contours_xy)):
                    for j in range(len(contours_xy[i])):
                        value.append(contours_xy[i][j][0][0]) #네번째 괄호가 0일때 x의 값
                        value1.append(contours_xy1[i][j][0][1]) #네번째 괄호가 0일때 x의 값
                        x_min = min(value)
                        x_max = max(value)+20
                        y_min = min(value1)
                        y_max = max(value1)+20
    
                x = x_min
                y = y_min
                w = x_max-x_min
                h = y_max-y_min
                plt.imshow(edged)
                time.sleep(0.5)
                img_trim = img[y:y+h, x:x+w]
                width, height, channel = img_trim.shape
                #img_trim_resize = cv2.resize(img_trim,dsize=(0,0), fx=0.5,fy=0.5,interpolation=cv2.INTER_AREA)
                time.sleep(0.5)
                if height < 440 :
                    img_trim = cv2.resize(img_trim,dsize=(0,0), fx=2.5,fy=2.5,interpolation=cv2.INTER_AREA)
                    width, height, channel = img_trim.shape
                    if width > 2440 :
                        img_trim = cv2.resize(img_trim,dsize=(0,0), fx=0.7,fy=0.7,interpolation=cv2.INTER_AREA)
                time.sleep(0.5)
                
                if self.parent.checkbox_value == 'true':
                    if  height+base_y < 3508:
                        base_img[base_x:(width+base_x), base_y:(height+base_y)] = img_trim

                        base_y = height+base_y+20
                    
                        if base_x_t > width+base_x+20 :
                            base_x_t = base_x_t
                        else :
                            base_x_t = width+base_x+20 
                        time.sleep(0.5)
                    elif  width+base_x_t < 2480:
                        base_y = 30
                        base_x = base_x_t
                        base_img[base_x:(width+base_x), base_y:(height+base_y)] = img_trim
                        if base_x_t > width+base_x+20 :
                            base_x_t = base_x_t
                            
                        else :
                            base_x_t = width+base_x+20 

                        base_y = height+base_y+20
                    
                        time.sleep(0.5)
                    else :
                        f_name = f_path + '/'+class_h+'_'+name+'_'+m+'_영수증_'+str(num)+'.jpg'
                        result,encode_img = cv2.imencode(f_name,base_img)
                        if result :
                            with open(f_name,mode='w+b') as f:
                                encode_img.tofile(f)
                    
                        num = num+1
                        base_img = np.zeros((2480,3508,3),np.uint8)
                        
                        base_x = 30
                        base_y = 30
                        base_x_t = 30
                        base_img[base_x:(width+base_x), base_y:(height+base_y)] = img_trim
                        if base_x_t > width+base_x+20 :
                            base_x_t = base_x_t
                            base_y = height+base_y+20
                        else :
                            base_x_t = width+base_x+20 
                            base_y = height+base_y+20
                        time.sleep(0.5)
                else :
                    if  height+base_y < 2480:
                        base_img[base_x:(width+base_x), base_y:(height+base_y)] = img_trim

                        base_y = height+base_y+20
                    
                        if base_x_t > width+base_x+20 :
                            base_x_t = base_x_t
                        else :
                            base_x_t = width+base_x+20 
                        time.sleep(0.5)
                    elif  width+base_x_t < 3508:
                        base_y = 30
                        base_x = base_x_t
                        base_img[base_x:(width+base_x), base_y:(height+base_y)] = img_trim
                        if base_x_t > width+base_x+20 :
                            base_x_t = base_x_t
                            
                        else :
                            base_x_t = width+base_x+20 

                        base_y = height+base_y+20
                    
                        time.sleep(0.5)
                    else :
                        f_name = f_path + '/'+class_h+'_'+name+'_'+m+'_영수증_'+str(num)+'.jpg'
                        result,encode_img = cv2.imencode(f_name,base_img)
                        if result :
                            with open(f_name,mode='w+b') as f:
                                encode_img.tofile(f)
                    
                        num = num+1
                        if self.parent.checkbox_value == 'true':
                            base_img = np.zeros((2480,3508,3),np.uint8)
                        else :
                            base_img = np.zeros((3508,2480,3),np.uint8)
                        base_x = 30
                        base_y = 30
                        base_x_t = 30
                        base_img[base_x:(width+base_x), base_y:(height+base_y)] = img_trim
                        if base_x_t > width+base_x+20 :
                            base_x_t = base_x_t
                            base_y = height+base_y+20
                        else :
                            base_x_t = width+base_x+20 
                            base_y = height+base_y+20
                        time.sleep(0.5)
                self.bar_count= int((self.e/(self.count-1))*100)
                self.parent.bar.setValue(self.bar_count)
                    #bar_count =Therad1(self)
                    #bar_count.run
                time.sleep(1)
               
            f_name = f_path + '/'+class_h+'_'+name+'_'+m+'_영수증_'+str(num)+'.jpg'
            result,encode_img = cv2.imencode(f_name,base_img)
            if result :
                with open(f_name,mode='w+b') as f:
                    encode_img.tofile(f)
           
            self.parent.stk_w.setCurrentIndex(1)
            time.sleep(0.5)
            #return self.mes
            #QMessageBox.information(self,"확인","변환완료")

        
class Main(QDialog,object):
    def __init__(self):
        super().__init__()
        self.main()
    def main(self):
               
        self.setWindowTitle("3Frame")
      
        self.widget_layout = QVBoxLayout()
       
        self.list_w = QListWidget()
        
        self.main_layout = QtWidgets.QHBoxLayout()
        self.name_widget =  QLineEdit("")
        self.name_laber = QLabel("이     름 ")

        #self.combovla = ['NEW BUSINESS DIV', 'DESIGN DEPT', 'INTERACTION DESIGN DEPT', 'ANIMATION DEPT',
                      #   'NEW MEDIA DEPT', 'VFX DEPT', 'CGI DIV', 'MATTE DEPT', 'MANAGEMENT', 'DIRECTOR DEPT',
                      #   'BUSINESS &MARKETING DEPT']
        self.class_widget =  QtWidgets.QComboBox()
        self.class_widget.addItems(['NEW BUSINESS DIV', 'DESIGN DEPT', 'INTERACTION DESIGN DEPT', 'ANIMATION DEPT',
                         'NEW MEDIA DEPT', 'VFX DEPT', 'CGI DIV', 'MATTE DEPT', 'MANAGEMENT', 'DIRECTOR DEPT',
                         'BUSINESS & MARKETING DEPT'])
        self.class_laber = QLabel("부 서 명 ") 
        
        self.search_layout_top = QtWidgets.QHBoxLayout()
        self.search_layout = QtWidgets.QVBoxLayout()
        self.name_layout = QtWidgets.QHBoxLayout()
        self.name_layout2 = QtWidgets.QHBoxLayout()
        self.class_layout = QtWidgets.QHBoxLayout()
        self.btn_layout = QtWidgets.QHBoxLayout()
        self.btn_layout2 = QtWidgets.QHBoxLayout()
        self.stk_w = QStackedWidget()
        self.stack_layout = QtWidgets.QStackedWidget()

        self.page_1 = QWidget()
        
        self.verticalLayout_1 = QHBoxLayout(self.page_1)
        self.page_2 = QWidget()
        
        self.verticalLayout_2 = QHBoxLayout(self.page_2)
        self.page_3 = QWidget()
        
        self.verticalLayout_3 = QHBoxLayout(self.page_3)




        name_group = QGroupBox()
        name_box = QBoxLayout(QBoxLayout.TopToBottom)
        name_group.setLayout(name_box)
        name_group.setTitle("")
        self.search_layout_top.addWidget(name_group)

        search_group = QGroupBox()
        srarch_box = QBoxLayout(QBoxLayout.TopToBottom)
        search_group.setLayout(srarch_box)
        search_group.setTitle("이미지")
        self.search_layout.addWidget(search_group)
      
        self.name_layout2.addWidget(self.name_laber,stretch=1)
        self.name_layout2.addWidget(self.name_widget,stretch=15)

        self.class_layout.addWidget(self.class_laber,stretch=1)
        self.class_layout.addWidget(self.class_widget,stretch=15)

        name_box.addLayout(self.class_layout,stretch=2)
        name_box.addLayout(self.name_layout2,stretch=2)

        self.bar=QProgressBar(self)
        self.bar.setValue(0)
        self.bar_label1 = QLabel("변환 완료")
        self.bar_label2 = QLabel("이미지를 확인해 주세요")
       
        self.name_layout.addLayout(self.search_layout_top,stretch=1)
        srarch_box.addWidget(self.list_w,stretch=19)

 
        self.list_w.setDragDropMode(QAbstractItemView.InternalMove)

        group = QGroupBox()
        box = QBoxLayout(QBoxLayout.TopToBottom)
        group.setLayout(box)
        group.setTitle("메뉴")
        self.main_layout.addWidget(group)

        button_1 = QPushButton("위")
        button_2 = QPushButton("아래")
        button_3 = QPushButton("삭제")
        self.button_4 = QPushButton("합치기")
        self.checkbox = QCheckBox("세로")
        self.checkbox_value = 'true'
        self.checkbox.stateChanged.connect(self.checkbox_changed)
        self.btn_layout2.addWidget(button_1)
        self.btn_layout2.addWidget(button_2)
        self.btn_layout2.addWidget(button_3)
        self.btn_layout.addWidget(self.button_4,stretch=9)
        self.btn_layout.addWidget(self.checkbox,stretch=1)
        box.addLayout(self.btn_layout2,stretch=2)
        box.addLayout(self.btn_layout,stretch=2)
        self.verticalLayout_1.addWidget(self.bar,stretch=2)
        self.verticalLayout_2.addWidget(self.bar_label1,stretch=2)
        self.verticalLayout_3.addWidget(self.bar_label2,stretch=2)
        self.stk_w.addWidget(self.page_1)
        self.stk_w.addWidget(self.page_2)
        self.stk_w.addWidget(self.page_3)
        box.addWidget(self.stk_w,stretch=2)
       
        self.widget_layout.addLayout(self.name_layout, stretch=1)
        self.widget_layout.addLayout(self.search_layout, stretch=19)
        self.widget_layout.addLayout(self.main_layout, stretch=1)
        
        
        button_1.clicked.connect(self.up_img)
        button_2.clicked.connect(self.down_img)
        button_3.clicked.connect(self.del_img)
        self.button_4.clicked.connect(self.marge_img)

        self.setLayout(self.widget_layout)
        self.move(0,0)
     
    def checkbox_changed(self, state):
        if state == 0 :
            self.checkbox_value = 'true'
        else :
            self.checkbox_value = 'false' 

    def del_img(self):
        self.list_w.takeItem(self.list_w.currentRow())
 
    def up_img(self):
        row = self.list_w.currentRow()
        if row > 0:
            item = self.list_w.takeItem(row)
            self.list_w.insertItem(row-1,item)
            self.list_w.setCurrentRow(row-1)

    def down_img(self):
        row = self.list_w.currentRow()
        if row < self.list_w.count()-1:
            item = self.list_w.takeItem(row)
            self.list_w.insertItem(row+1,item)
            self.list_w.setCurrentRow(row+1)
    
    def marge_img(self):
        self.mes = 0
        img =Therad2(self)
        img.start()
        
        
        

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main = Main()
    main.show()
    sys.exit(app.exec_())
