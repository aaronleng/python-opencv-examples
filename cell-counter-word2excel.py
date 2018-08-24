import cv2
import numpy as np
import xlwt
import os
import zipfile
import shutil


# 定义解压函数
def un_zip(file_name):
    """unzip zip file"""
    zip_file = zipfile.ZipFile(file_name)
    if os.path.isdir(file_name + "_files"):
        pass
    else:
        os.mkdir(file_name + "_files")
    for names in zip_file.namelist():
        zip_file.extract(names,file_name + "_files/")
    zip_file.close()

# 从word文档中读取图片
os.chdir(r'D:\python\opencv\data')  # 首先改变目录到文件的目录
shutil.copyfile("6M45℃（CPAs=1h）1.docx","s4.docx")
os.rename('s4.docx','test.ZIP')  # 重命名为zip文件
un_zip('test.Zip') # 解压文件
os.remove('/python/opencv/data/test.ZIP')  # 删除压缩文件
pic_number=len(os.listdir('/python/opencv/data/test.Zip_files/word/media'))  # 需处理的图片数目


# 使用 blob 算法和 canny 算法 进行细胞识别
j=1
number=[]
while (j< pic_number+1):
    # Read image
    img = cv2.imread("/python/opencv/data/test.Zip_files/word/media/image%d.png"%j,)
    # cv2.imshow('s',img)
    # 尺寸裁剪为600*600
    im = cv2.resize(img, (600, 600), interpolation=cv2.INTER_CUBIC)
    # cv2.imwrite('2.jpg', im)
    # 灰度化
    gray = cv2.cvtColor(im, cv2.COLOR_BGR2GRAY)

    # 高斯滤波
    gauss = cv2.GaussianBlur(gray, (3, 3), 0)
    # canny边缘检测算法
    im = cv2.Canny(gauss, 11, 35)


    # Setup SimpleBlobDetector parameters.
    params = cv2.SimpleBlobDetector_Params()

    # Change thresholds
    params.minThreshold = 10
    params.maxThreshold = 200

    params.filterByColor = True
    params.blobColor =0

    # Filter by Area.
    params.filterByArea = True
    params.minArea = 20
    params.maxArea = 160

    # Filter by Circularity
    params.filterByCircularity = True
    params.minCircularity = 0.1


    # Filter by Convexity
    params.filterByConvexity = True
    params.minConvexity = 0.1

    # Filter by Inertia
    params.filterByInertia = True
    params.minInertiaRatio = 0.1
    params.maxInertiaRatio = 1
    # Create a detector with the parameters
    detector = cv2.SimpleBlobDetector_create(params)

    # Detect blobs.
    keypoints = detector.detect(im)
    count = len(keypoints) # 细胞数
    number.append(count)   # 将细胞数存储在数组中

    # Draw detected blobs as red circles.
    # cv2.DRAW_MATCHES_FLAGS_DRAW_RICH_KEYPOINTS ensures
    # the size of the circle corresponds to the size of blob

    im_with_keypoints = cv2.drawKeypoints(im, keypoints, np.array([]), (0, 0, 255),
                                          cv2.DRAW_MATCHES_FLAGS_DRAW_RICH_KEYPOINTS)  #用红圈标记细胞

    font = cv2.LINE_AA
    result = str(count)
    output_image = cv2.putText(im_with_keypoints, result, (0, 30), font, 1, (0, 255, 0), 2) #标记细胞数
    cv2.imwrite('/python/opencv/cropImg/%d.jpg' % j, output_image)
    # cv2.namedWindow('result',cv2.WINDOW_NORMAL)
    # cv2.imshow('result',output_image)
    j += 1

shutil.rmtree('/python/opencv/data/test.Zip_files') # 删除多余文件
# 将数据保存到excel文档
book = xlwt.Workbook() #创建一个Excel
sheet1 = book.add_sheet('hello') #在其中创建一个名为hello的sheet

n = 0 #行序号
while (n<pic_number):
    x = number[n]
    sheet1.write(n,0,x) #在新sheet中的第1行第n+1列写入读取到的x值
    n = n+1 #列号递增
# sheet1.write(0,0,'细胞数') #在新sheet中的第1行第n列写入读取到的x值
book.save('/python/opencv/data/cell_number.xls') # 创建保存文件

cv2.waitKey(0)
