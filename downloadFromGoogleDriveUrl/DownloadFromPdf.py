import PyPDF2
import urllib.request
import os
import requests
import pdfplumber
import xlwt
import xlrd
import time 

if __name__ == "__main__":
    
    # 定义保存Excel的位置
    workbook = xlwt.Workbook()  #定义workbook
    sheet = workbook.add_sheet('Sheet1')  #添加sheet
    i = 0 # Excel起始位置

    pdf = pdfplumber.open("Archivio PSPT 2015 - 2015.pdf")
    print('开始读取数据')
    for pageNum in range(len(pdf.pages)):
        # 获取当前页面的全部文本信息，包括表格中的文字
        #print(page.extract_text())
        #print(page.extract_text().replace('\n',' ').split(' '))
        #print(page.extract_tables())
        pageAll = pdf.pages[pageNum].extract_text().replace('\n',' ').split(' ')
        #print(pageAll[0:6])
        print("正在写入第"+str(pageNum)+"页数据...")
        if pageNum == 0:
            pdfTitle = pageAll[0:6]
            pdfData = pageAll[6:]
            pdfdata = []
            for column in range(0, len(pdfData), 5):
                #对列表按行数分组，这里一行有5个数组
                pdfdata.append(pdfData[column : column+5])
                
            for j in range(len(pdfTitle)):
                sheet.write(i, j, pdfTitle[j])
            i += 1
            for row in pdfdata:
                for j in range(len(row)):
                    sheet.write(i, j, row[j])
                i += 1
        else:
            pdfData = pageAll
            pdfdata = []
            for column in range(0, len(pdfData), 5):
                pdfdata.append(pdfData[column : column+5])

            for row in pdfdata:
                for j in range(len(row)):
                    sheet.write(i, j, row[j])
                i += 1    
    pdf.close()
    
    # 保存Excel表
    workbook.save('C:/Users/WKN/Downloads/Archivio PSPT 2015 - 2015.xls')
    print('写入excel成功')
    
    excelFile = xlrd.open_workbook('C:/Users/WKN/Downloads/Archivio PSPT 2015 - 2015.xls',formatting_info=True)
    sheet1 = excelFile.sheet_by_name("Sheet1")
    for rowNum in range(576,sheet1.nrows):
        #print(rowNum)
        if rowNum == 0:
            continue
        else:
            rowValue = sheet1.row_values(rowNum)[0:5]
            if rowValue[3] == 'K':
                GoogleDriveURLId = rowValue[4].split('/')[-2]
                downloadUrl = "http://drive.google.com/uc?export=download"+"&id="+GoogleDriveURLId
                #print("%02d"%(int(rowValue[1])))
                imgName = rowValue[0] + "%02d"%(int(rowValue[1])) + \
                          "%02d"%(int(rowValue[2])) + "." + \
                          rowValue[3] + ".1.2" + ".jpg"
                print("Try downloading file: {}".format(imgName))
                filepath = 'D:/工作/天文图像集/Kfilter/2015/'+ imgName
                try:
                    opener=urllib.request.build_opener()
                    opener.addheaders=[('User-Agent','Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36')]
                    urllib.request.install_opener(opener)
                    urllib.request.urlretrieve(downloadUrl, filename = filepath)
                    time.sleep(5)
                except Exception as e:
                    opener=urllib.request.build_opener()
                    opener.addheaders=[('User-Agent','Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36')]
                    urllib.request.install_opener(opener)
                    urllib.request.urlretrieve(downloadUrl, filename = filepath)
                    time.sleep(5)
                    print("Error occurred when downloading file, error message:")
                    print(e)
        
    
