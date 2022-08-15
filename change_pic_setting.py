from PIL import Image
#import tempfile
import os
import pandas as pd
import shutil


def picFileCopy(xls_file_path,src_dir_path,to_dir_path):
    '''
    目的: 找到包含指定關鍵字的檔案,並另複制到其他資料夾

    設定參數:
    filename.xls :新增一個xls檔案,只要old那欄要抓取檔案的名稱。(不用含副檔名.jpg)
    src_dir_path:照片存檔的資料夾路徑
    to_dir_path:目的資料夾
    '''

    #src_dir_path = 'C:/Users/runra/Desktop/test2/'        # 源文件夾
    #to_dir_path = 'C:/Users/runra/Desktop/test3/'         # 存放複製文件的文件夾
    #key = 'a1'                 # 源文件夾中的文件包含字符key則覆制到to_dir_path文件夾中
    

    if not os.path.exists(to_dir_path):
        print("to_dir_path not exist,so create the dir")
        os.mkdir(to_dir_path, 1)
    if os.path.exists(src_dir_path):
        print("src_dir_path exist")

        file=pd.read_excel(xls_file_path+"filename.xls")
        for key in file['old']:
            for file in os.listdir(src_dir_path):
                # is file
                if os.path.isfile(src_dir_path+'/'+file):
                    if key in file:
                        print('找到包含"'+key+'"關鍵字的文件,絕對路徑為----->'+src_dir_path+'/'+file)
                        print('複製到----->'+to_dir_path+file)
                        shutil.copy(src_dir_path+'/'+file, to_dir_path+'/'+file)# 移動用move函数



def listPicFileInfo(src_dir_path):
    '''
    目的: 得到指定資料夾圖片檔的明細

    設定參數:
    src_dir_path:照片存檔的資料夾路徑
    
    '''
    #src_dir_path = "C:/Users/runra/Desktop/test/"   #記得改用 / 反斜線
    li=os.listdir(src_dir_path)

    for ims in li:
        im = Image.open(src_dir_path+ims)
        size = os.path.getsize(src_dir_path+ims)
        nx, ny = im.size
        f = im.format
        
        if im.info.get('dpi'):
            x_dpi, y_dpi = im.info['dpi']
            print("檔名: {:<15s}  像素: {}*{}  格式: {}  圖片尺寸: {:.1f} KB  DPI(x_dpi*y_dpi): {}*{}  ".format(
                                                    ims,nx,ny,f,size/1024,str(x_dpi),str(y_dpi)))
        else:
            print("檔名: {:<15s}  像素: {}*{}  格式: {}  圖片尺寸: {:.1f} KB  DPI(x_dpi*y_dpi): {}  ".format(
                                                    ims,nx,ny,f,size/1024,"No DPI data. Invalid Image header"))
        im.close()

def changeDpi(src_dir_path,newDpi):
    '''
    目的: 改資料夾內jpg圖片的dpi,並另存新檔後加dpi規格說明

    設定參數:
    src_dir_path:照片存檔的資料夾路徑
    newDip:要另存設定的的新dpi
    '''
    #src_dir_path = "C:/Users/runra/Desktop/test/"   #記得改用 / 反斜線
    li=os.listdir(src_dir_path)
    for ims in li:
        im = Image.open(src_dir_path+ims)
        nx, ny = im.size
        f = im.format
        print("原始照片尺寸: ",nx,"x",ny,"圖片格式: ",f)
        fileName=ims.split(".")

        if im.info.get('dpi'):
            x_dpi, y_dpi = im.info['dpi']
            if x_dpi != y_dpi:
                print('Inconsistent DPI image data')
            print("原本x_dpi: " + str(x_dpi),"原本y_dpi: " + str(y_dpi))
        else:
            print('No DPI data. Invalid Image header')
            #print(im.info)
        
        #newDpi=300  #設定新的dpi
        newFileName=fileName[0]+"_"+str(newDpi)+"dpi"+".jpg"        #另存新檔
        #newFileName=fileName[0]+".jpg"                             #直接覆蓋
        im.save(f"{src_dir_path}/"+newFileName,dpi=(newDpi,newDpi),quality=100)
        im.close()
    print("資料夾jpg檔案轉換成",newDpi,"dpi結束!")


def changePicSize(src_dir_path,max_width):
    '''
    目的: 等比例放大/縮小圖片

    設定參數:
    src_dir_path:照片存檔的資料夾路徑
    max :輸入指定寬度
    '''
    #src_dir_path = "C:/Users/runra/Desktop/test/"   #記得改用 / 反斜線
    li=os.listdir(src_dir_path)
    for ims in li:
        im = Image.open(src_dir_path+ims)
        width,height = im.size
        #max = 200                     # 指定寬最大的數值
        scale = height/width           # 設定 scale 為 height/width
        w = max_width                  # 設定調整後的寬度為最大的數值
        h = int(max_width*scale)       # 設定調整後的高度為 max 乘以 scale ( 使用 int 去除小數點 )
        im2 = im.resize((w, h))        # 調整尺寸
        
        fileName=ims.split(".")
        newFileName=fileName[0]+"_"+"width"+str(max_width)+".jpg"   #另存新檔
        #newFileName=fileName[0]+".jpg"                             #直接覆蓋
        im2.save(f"{src_dir_path}/"+newFileName,quality=100) 
        im.close()
        print(ims,"圖片放大/縮小結束!")



def changeFileName(xls_file_path,src_dir_path):
    '''
    目的: 更改檔名

    設定參數:
    xls_file_path:放old,new檔名xls檔的位置,不要跟圖片檔同一個資料夾
    src_dir_path:照片存檔的資料夾路徑
    filename.xls :新增一個xls檔案,內放二欄'old','new'檔名資料,old欄要寫含副檔名的完整名稱
    '''
    #xls_file_path="C:/Users/runra/Desktop/test/"        #放old,new檔名xls檔的位置
    #src_dir_path = "C:/Users/runra/Desktop/test/pic/"   #記得改用 / 反斜線

    file=pd.read_excel(xls_file_path+"filename.xls")
    #print(file.head())
    old=set(file['old'].tolist())
    file.set_index("old" , inplace=True)

    li=os.listdir(src_dir_path)
    for ims in li:
        if ims in old:
            fileName=ims.split(".") 
            new=file.loc[ims].new
            os.rename(f"{src_dir_path}/"+ims,f"{src_dir_path}/"+new)   
        else:
            print(ims,"未改名!")
        
    print("圖片更名結束!")



def changePNG_to_JPG(src_dir_path):
    '''
    目的: PNG另存jpg

    設定參數:
    src_dir_path:照片存檔的資料夾路徑
    
    '''
    #src_dir_path = "C:/Users/runra/Desktop/test/pic/"   #最後要加雙斜線,不然會報錯

    li=os.listdir(src_dir_path)
    for ims in li:
        fileName=ims.split(".")
        im = Image.open(src_dir_path+ims)
        newFileName=fileName[0]+".jpg"   #另存新檔
        im.save(f"{src_dir_path}/"+newFileName,quality=100) 
        im.close()
    print("PNG圖片另存jpg結束!")





#主功能區---------------------------------------------------------------
'''備註
    在檔案夾內開新文字檔，將以下文字貼入，存成.bat，即可列出資料夾內所有檔名
    @echo off
    dir /b /on >list.txt
'''


xls_file_path="C:/Users/runra/Desktop/test3/"       #變更檔名時，放old,new檔名xls的位置
#src_dir_path = "C:/Users/runra/Desktop/test/pic/"  #記得改用 / 反斜線
src_dir_path = "C:/Users/runra/Desktop/test2/"      #源文件位置,記得改用 / 反斜線
to_dir_path= "C:/Users/runra/Desktop/test3/"        #另存新檔的位置



picFileCopy(xls_file_path,src_dir_path,to_dir_path)       #copy指定關鍵字檔案
#listPicFileInfo(src_dir_path)                 #得到指定資料夾圖片檔的明細
#hangeDpi(src_dir_path,300)                    #改變指定資料夾圖片dpi
#changePicSize(src_dir_path,450)               #變更指定資料夾內,指定圖片寬度，等比例放大/縮小
#changeFileName(xls_file_path,src_dir_path)    #大量更改檔名
#changePNG_to_JPG(src_dir_path)                #png 另存jpg


