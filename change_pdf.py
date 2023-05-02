import openpyxl as op
import os
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import sys
from PyQt5 import QtCore, QtWidgets
import datetime
import xlwings as xw

today_data = datetime.datetime.now()
f_path =os.getcwd()+'/'+today_data.strftime("%Y%m")

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
    pop_msg=pyqtSignal(str)
    def __init__(self, parent) :
        super().__init__(parent)
        self.parent = parent
    
    def run(self):
        app = xw.App(visible=False)
        f_path =os.getcwd()+'/'+today_data.strftime("%Y%m")
       # xl = Dispatch("Excel.Application")
       
        select_list =self.parent.sheet_list.selectedItems()
        #wb = xl.Workbooks.Open(self.parent.filepath) #info가 튜플이므로 인덱싱으로 접근(0번째는 파일경로)
        wb= xw.Book(self.parent.filepath)

        #excelinfo를 입력받아 for문 실행
        for info in select_list:
            select_sheet_name=info.text()
            #ws = wb.Worksheets(select_sheet_name) #튜플의 2번째 요소는 시트명임. 
            ws = wb.sheets[select_sheet_name]
            #ws.Select() #위 설정한 시트 선택
            #wb.ActiveSheet.ExportAsFixedFormat(0, f_path+"/"+select_sheet_name+".pdf") #파일명, 시트명으로 pdf 파일명 저장
            current_work_dir = f_path
            pdf_path = os.path.join(current_work_dir, select_sheet_name+".pdf") 
            report_sheet = wb.sheets[select_sheet_name]
            report_sheet.api.ExportAsFixedFormat(0, pdf_path)



            
        #wb.Close(False) #workbook 닫기, True일 경우 그 상태를 저장한다.
        p_msg='T'
        self.pop_msg.emit(p_msg)
        app.kill()
        #xl.Quit()  # excel application 닫기

    
class Main(QDialog,object):
    def __init__(self):
        super().__init__()
        self.main()
    def main(self):
               
        self.filepath = "" 
        self.sheet_name=[]     
           
        self.setWindowTitle("PDF 변환기")
     
        self.widget_layout = QHBoxLayout()
       
        self.file_where = QLineEdit()
        
        self.main_layout = QtWidgets.QHBoxLayout()

        self.search_layout = QtWidgets.QVBoxLayout()
       
        search_group = QGroupBox()
        srarch_box = QBoxLayout(QBoxLayout.TopToBottom)
        file_box=QBoxLayout(QBoxLayout.LeftToRight)
        search_group.setTitle("엑셀 파일")
        self.pushButton = QPushButton("File Open")
        self.pushButton.clicked.connect(self.pushButtonClicked)
        self.trans_btn1 = QPushButton("변환")
        self.trans_btn2 = QPushButton("초기화")

        btn_box=QBoxLayout(QBoxLayout.LeftToRight)
        btn_box.addWidget(self.trans_btn1,stretch=1)
        btn_box.addWidget(self.trans_btn2,stretch=1)

        file_box.addWidget(self.file_where,stretch=1)
        file_box.addWidget(self.pushButton,stretch=1)
        search_group.setLayout(srarch_box)


        self.trans_btn3 = QPushButton("전체선택")
        self.trans_btn4 = QPushButton("전체해제")
        btn_box2=QBoxLayout(QBoxLayout.LeftToRight)
        btn_box3=QBoxLayout(QBoxLayout.TopToBottom)
        btn_box2.addWidget(self.trans_btn3)
        btn_box2.addWidget(self.trans_btn4)
        btn_box3.addLayout(btn_box2)
        srarch_box.addLayout(file_box,stretch=1)
        srarch_box.addLayout(btn_box3,stretch=3)
        srarch_box.addLayout(btn_box,stretch=1)
        self.search_layout.addWidget(search_group)

        self.widget_layout.addLayout(self.search_layout, stretch=1)
        
        
        self.sheet_list = QListWidget()
        self.sheet_list.setSelectionMode(QtWidgets.QAbstractItemView.MultiSelection)
        #self.sheet_list.itemChanged.connect(self.pushButtonClicked)
        self.sheet_list.itemClicked.connect(self.chkItemClicked)


        self.widget_layout.addWidget(self.sheet_list, stretch=1)
        self.widget_layout.addLayout(self.main_layout, stretch=1)
     
        self.trans_btn3.clicked.connect(lambda: self.select(Qt.Checked))
        self.trans_btn4.clicked.connect(lambda: self.select(Qt.Unchecked))
        self.setLayout(self.widget_layout)
        self.move(0,0)

        self.trans_btn1.clicked.connect(self.mkpdf)
        self.trans_btn2.clicked.connect(self.ClearAll2)


        self.pdf_mk =Therad1(parent=self)
        self.pdf_mk.pop_msg.connect(self.pop_m)


        

    def select(self, state):
        for row in range(self.sheet_list.count()):
            itm = self.sheet_list.item(row)
            itm.setSelected(state)
            itm.setCheckState(state)
    

    def chkItemClicked(self) :
        print(self.sheet_name)
        select_list =self.sheet_list.selectedItems()
        for i in select_list:
            print(i.text())
        if self.sheet_list.currentItem().checkState() == 2:
            
            self.sheet_list.currentItem().setCheckState(QtCore.Qt.Unchecked)
        elif self.sheet_list.currentItem().checkState() == 0:
            self.sheet_list.currentItem().setCheckState(QtCore.Qt.Checked)

    def ClearAll(self):
        self.sheet_name=[]  
       
        self.sheet_list.clear()
        self.excelInfo()

    def ClearAll2(self):
        self.sheet_name=[]  
        
        self.sheet_list.clear()
        self.file_where.setText("")

    def pushButtonClicked(self):
        fname = QFileDialog.getOpenFileName(self)
        self.file_where.setText(fname[0])
        f_path=fname[0]
        self.filepath=f_path
        self.ClearAll()
        
           
    def excelInfo(self):
        #result = [] #빈 리스트 생성
       
        wb = op.load_workbook(self.filepath) #openpyxl workbook 생성
        ws_list = wb.sheetnames #해당 workbook의 시트명을 리스트로 받음
        self.filename = os.path.basename(self.filepath)
        self.filename = self.filename.replace(".xlsx","")
        for sht in ws_list: #시트명 리스트를 for문을 통해 반복
            temp_tuple = (sht) #파일경로, 파일명, sht를 튜플에 저장
            self.sheet_name.append(temp_tuple) #위 튜플을 빈 리스트에 추가
        #print(result)
        self.AddItem()
        return self.sheet_name # 튜플로 이루어진 리스트 리턴
    def AddItem(self):
        
        for data in self.sheet_name:
            item = QtWidgets.QListWidgetItem()
            item.setCheckState(QtCore.Qt.Unchecked)
            item.setText(data)
            self.sheet_list.addItem(item)
            
            

    def createF(self,dir1):
        try:
            if not os.path.exists(dir1):
                os.makedirs(dir1)
        except OSError:
            print("폴더생성이 되지 않았습니다.")

    def mkpdf(self):
        self.createF(os.getcwd()+'/'+today_data.strftime("%Y%m"))
        self.pdf_mk.start()

    

    @pyqtSlot(str)
    def pop_m(self,msg):
        if msg == 'T' :
            QMessageBox.information(self,"확인","변환완료")
        else :
            QMessageBox.question(self,"확인","파일을 확인해주세요")

  
if __name__ == '__main__':
    app = QApplication(sys.argv)
    main = Main()
    main.show()
    sys.exit(app.exec_())
