import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtCore import Qt
from PyQt5 import QtWidgets
from PyQt5.QtGui import QIcon
import win32serviceutil
import subprocess
import os
import time
class QLabel(QLabel):
    def __init__(self, parent = None):
        super().__init__(parent)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

class QLineEdit(QLineEdit):
    def __init__(self, parent = None):
        super().__init__(parent)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)


   
class Main(QDialog,object):
  
    def __init__(self):
        
        super().__init__()
        screen_rect =QDesktopWidget().availableGeometry()
        self.wi,self.he = screen_rect.width(), screen_rect.height()
        self.widget_layout = QVBoxLayout(self)    
        
        ff = open("C:/Program Files/Zabbix Agent/zabbix_agentd.conf",'r')
        for line in ff :
            if 'server=127' in line:
                pass
            elif '# Server =' in line :
                pass
            elif 'Server=' in line:
                server_name = line
        ip_name = server_name[server_name.find('=')+1:len(server_name)-1]
        print(ip_name)

        
        
        self.main_widget = QtWidgets.QWidget()
        self.main_layout = QtWidgets.QHBoxLayout()
        self.main_widget.setLayout(self.main_layout)
        self.old_zabbix_ip_layout = QtWidgets.QHBoxLayout()
        self.new_zabbix_ip_layout = QtWidgets.QHBoxLayout()
        self.zabbix_ip_layout = QtWidgets.QVBoxLayout()
        self.old_zabbix_ip_name = QLabel("Zabbix 이전 IP")
        self.old_zabbix_ip =  QLabel(ip_name)
        self.new_zabbix_ip_name = QLabel("Zabbix 새 IP")
        self.new_zabbix_ip =  QLineEdit()
        button_1 = QPushButton("교체")

        print(ip_name)
        self.old_zabbix_ip_layout.addWidget(self.old_zabbix_ip_name)
        self.old_zabbix_ip_layout.addWidget(self.old_zabbix_ip)
        self.new_zabbix_ip_layout.addWidget(self.new_zabbix_ip_name)
        self.new_zabbix_ip_layout.addWidget(self.new_zabbix_ip)
        self.zabbix_ip_layout.addLayout(self.old_zabbix_ip_layout)
        self.zabbix_ip_layout.addLayout(self.new_zabbix_ip_layout)
        self.widget_layout.addLayout(self.zabbix_ip_layout)
        self.widget_layout.addWidget(button_1)
        self.show()
        self.old_ip = self.old_zabbix_ip.text()
       # self.new_ip = self.new_zabbix_ip.text()
        #self.replace_in_file("C:/Program Files/Zabbix Agent/zabbix_agentd.conf", lines, self.new_ip)
        button_1.clicked.connect(self.change_ip)
    

    def change_ip(self):
        #self.replace_in_file("C:/Program Files/Zabbix Agent/zabbix_agentd.conf", self.old_ip, self.new_ip)
        file_path = "C:/Program Files/Zabbix Agent/zabbix_agentd.conf"
        #file_path = "./zabbix_agentd.conf"
        old_str = self.old_ip
        new_str = self.new_zabbix_ip.text()
        service_name = 'Zabbix Agent'
        win32serviceutil.StopService(service_name)
        time.sleep(1)
        #win32serviceutil.StopService(service_name)
        fr = open(file_path, 'r')
        lines = fr.readlines()
        fr.close()
        ck_line = ''
        if old_str == '':
            ck_line ='Server=\n'
        else :
            ck_line='Server='+str(old_str)+'\n'
        # old_str -> new_str 치환
        fw = open(file_path, 'w')
        for line in lines:
            if old_str == '':
                if line == ck_line:

                    fw.write(line.replace('Server=', 'Server='+new_str,-1))
                else :
                    fw.write(line)
            else :    
                if line == ck_line :
                    fw.write(line.replace(old_str, new_str,-1))
                else :
                    fw.write(line)
        fw.close()
        time.sleep(1)
        win32serviceutil.StartService(service_name)
        time.sleep(1)
        subprocess.run('zabbix_port_open.bat')
        QMessageBox.information(self,"확인","변환완료")
# 호출: file1.txt 파일에서 comma(,) 없애기


if __name__ == '__main__':
    
    app = QApplication(sys.argv)
    main = Main()
    #main.conn()
    main.show()
    #main.update()
    
    sys.exit(app.exec_())
