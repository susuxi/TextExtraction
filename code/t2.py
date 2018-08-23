# -*- coding: utf-8 -*-
"""
Created on Sun Apr 15 01:12:59 2018

@author: 晨
"""

from PyQt5.QtWidgets import (QApplication, QWidget,QPushButton,QDesktopWidget,
                             QLabel,QLineEdit,QGridLayout,QFileDialog,QTextEdit)

import time
import os
import sys
import ai


class window(QWidget):
    
    def __init__(self):
        super().__init__()
        self.initUI()
    
    def initUI(self):
        chosebt1 = QPushButton('选择文件夹',self)
        chosebt2 = QPushButton('选择文件夹',self)
        startbt = QPushButton('开始识别',self)
        savebt = QPushButton('导出excel',self)

        
        imgadress = QLabel('图片地址:')
        exceladress = QLabel('excel存入:')
        
        self.imgEdit = QLineEdit()
        self.excelEdit = QLineEdit()
        self.result = QTextEdit()
        
        
        
        
        grid = QGridLayout()
        grid.setSpacing(10)

        grid.addWidget(imgadress, 1, 0)
        grid.addWidget(self.imgEdit, 1, 1)
        grid.addWidget(chosebt1, 1, 2)
        grid.addWidget(exceladress, 2, 0)
        grid.addWidget(self.excelEdit, 2, 1)
        grid.addWidget(chosebt2, 2, 2)
        grid.addWidget(self.result,3,0,5,4)
        grid.addWidget(startbt,1,3,1,1)
        grid.addWidget(savebt,2,3,1,1)
        
        
      
        result = []
        self.setLayout(grid) 
        
        self.setGeometry(500,500,450,300)
        self.center()
        self.setWindowTitle('工商图片识别')
        self.show()
        
        
        chosebt1.clicked.connect(self.Imgchosedic)
        chosebt2.clicked.connect(self.Excchosedic)
        startbt.clicked.connect(lambda:self.recognize(result))
        savebt.clicked.connect(lambda:self.export(result))


    
    def Imgchosedic(self):
        dic_name = QFileDialog.getExistingDirectory(self,'选择文件夹','E:\\',QFileDialog.ShowDirsOnly
                                                | QFileDialog.DontResolveSymlinks)
        self.imgEdit.setText(dic_name)
    
    def Excchosedic(self):
        dic_name = QFileDialog.getExistingDirectory(self,'选择文件夹','E:\\',QFileDialog.ShowDirsOnly
                                                | QFileDialog.DontResolveSymlinks)
        self.excelEdit.setText(dic_name)
    
    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def recognize(self,result):
        start = time.time()
        
        self.result.setText("-----------识别开始-------------")
        QApplication.processEvents()
        
        imgdir = self.imgEdit.text()
        i = 1
        for file in os.listdir(imgdir):  
            file_path = os.path.join(imgdir,file) 
            self.result.append(file+': ')
            QApplication.processEvents()
            if os.path.splitext(file)[1]!='.png' and os.path.splitext(file)[1]!='.jpg':
                self.result.append('文件类型错误，仅支持png或jpg格式图片')
            else:           
                com_name,reg_id = ai.preprocess(file_path,i,result)
                self.result.append(com_name+' '+reg_id+'\n')
            
            #滚轮保持在最底端
            self.scroll()
            QApplication.processEvents()
        end = time.time()
        self.result.append("-----------识别完成-------------")
        t = end-start
        #print(t)
        self.result.append("共用时"+str(t)+"s")
        
    
    def export(self,result):
        exportdir = self.excelEdit.text()
        if(not os.path.exists(exportdir)):
            self.result.append("导出地址为空或不存在！！\n")
            QApplication.processEvents()
        else:
            self.result.append(" ")
            self.result.append("-----------正在导出-------------\n")
            self.scroll()
            QApplication.processEvents()
            ai.create_excel(result,exportdir)
            self.result.append("-----------导出成功-------------\n")      
            
    def scroll(self):
        cursor = self.result.textCursor()
        pos = len(self.result.toPlainText())
        cursor.setPosition(pos)
        self.result.setTextCursor(cursor)
                   

if __name__ == '__main__':
    
    app = QApplication(sys.argv)

    ex = window()
    sys.exit(app.exec_())