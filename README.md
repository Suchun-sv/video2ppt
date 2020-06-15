# video2ppt
capture the image of the video and generate pptx file.

# TODO
installer 成 exe
添加图形化界面

# 效果展示(懒得放图了)
*.mp4 -> tempPicture/ -> tempPicture_reduce/ ->*.pptx

# Start
## 1. 选定mp4的路径
abc = video2pptx(".//英语课_15周.mp4")
## 2. 隔指定时间(秒s)截图到指定文件夹
abc.capFrame(".//tempPicture",60)
## 3. 利用相似度计算算法计算相邻两张图片相似度
similar_score = abc.calcSimilar(".\\tempPicture")
print(similar_score)
## 4. 指定阈值，将相似度小于阈值的图片挑选出来
abc.copyPictureBySimilar(0.93,".\\tempPicture",similar_score=similar_score)
## 5. 从指定文件夹所有图片生成pptx
abc.createPPtx("英语课_15周",pictureFile=r"./tempPicture_reduce/")

# requirements
```
import cv2
import tqdm
from PIL import Image
import os
import shutil
from pptx import Presentation
from pptx.util import Inches
```