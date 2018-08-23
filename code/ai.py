# -*- coding: utf-8 -*-
"""
Created on Fri May 18 16:54:36 2018

@author: 晨
"""
import cv2
import numpy as np
#import time
#import os
import pytesseract
import xlwt
import xlrd




"""
一般图片去水印
"""
def rm_watermark(image):

    alp_channel = image[:,:,3]
    rgb_channel = image[:,:,:3]
    
    white_background = np.ones_like(rgb_channel,dtype = np.uint8) * 255
    
    alp_factor = alp_channel[:,:,np.newaxis].astype(np.float32) / 255.0
    alp_factor = np.concatenate((alp_factor,alp_factor,alp_factor),axis=2)
    
    
    base = rgb_channel.astype(np.float32) * alp_factor
    white = white_background.astype(np.float32) * (1 - alp_factor)
    final_image = base + white
    return final_image.astype(np.uint8)


"""
图片预处理+识别
"""
#start = time.time()
def preprocess(filename,i,result):
    image = cv2.imread(filename,cv2.IMREAD_UNCHANGED)
    if image.shape[2] == 3:
        #print("error: can't recognize")
        return 'error:','image can not regconize'
    else:
        img = rm_watermark(image)

        #图像二值化处理，低于阈值的像素点灰度值置为0；高于阈值的值置为参数3
        ret,thresh = cv2.threshold(img,127,255,cv2.THRESH_BINARY_INV)
        g_img = cv2.cvtColor(thresh,cv2.COLOR_BGR2GRAY) #灰度化
    
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (20, 5)) #膨胀处理核函数
        dilation = cv2.dilate(g_img, kernel, iterations = 1)
    
    
     #框定文字选区
    
    image,contours, hierarchy = cv2.findContours(dilation,cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE)
    flag = 0
    for j in range(len(contours)-2,len(contours)):
        
        x,y,w,h = cv2.boundingRect(contours[j])
        newimg = img[y:y+h,x:x+w]
        ret2,thresh2 = cv2.threshold(newimg,135,255,cv2.THRESH_BINARY)
        #图片切割结果查看，测试用
        #newdir = ("D:\\B-AI\\train\\")
        #if not os.path.isdir(newdir):
        #    os.makedirs(newdir)
        #cv2.imwrite(newdir+str(i)+"_"+str(j)+".jpg",thresh2)
        # print(newdir+str(i)+"_"+str(j))
        word = pytesseract.image_to_string(thresh2,lang = 'chi_sim',config ='--psm 7')
        if flag == 0:
            com_name = word[7:]
        else :
            reg_id = word[8:]
        flag = flag + 1
    result.append([com_name,reg_id]) 
    return com_name,reg_id

#图片识别
    
"""
输出为excel表格
"""
def create_excel(result,exportdir):
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet('info',cell_overwrite_ok=True)
    sheet.write(0,0,'公司名称')
    sheet.write(0,1,'企业注册号')
    for i in range(0,len(result)):
        sheet.write(i+1,0,result[i][0])
        sheet.write(i+1,1,result[i][1])
    first_col=sheet.col(0)
    sec_col=sheet.col(1)
    first_col.width = 256*32
    sec_col.width = 256*27
    book.save(exportdir+'/result.xls')
    
    
"""
单项识别数据追加excel表格
"""    
    
def append_excel(com_name,reg_id,cur_img):  
    book = xlrd.open_workbook(r'E:\cut\test1.xls')
    print(book.get_sheet(0))
    #sheet.write(cur_img,0,com_name)    
    #sheet.write(cur_img,1,reg_id)
    book.save(r'E:\cut\test1.xls')
    
    
    
"""
测试主程序
"""  
#img_dir = "D:\\B-AI\\train\\"
#amount = 6
#result = []
#for i in range(1,6):
#    filename = img_dir+str(i)+".png"
#    print(filename)
#    result,img = preprocess(filename,i,result)
#create_excel(result)


#end = time.time()
#print(end-start)
#cv2.namedWindow("Image")  #创建窗口并显示图像
#cv2.imshow("Image",img)
#cv2.waitKey(0)
    #释放窗口
#cv2.destroyAllWindows() 

