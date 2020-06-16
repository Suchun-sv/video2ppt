# 致谢
[@amberOoO](https://github.com/amberOoO)
# video2ppt
capture the image of the video and generate pptx file.

# TODO
installer 成 exe(300MB+)
添加图形化界面(搁置)

# 效果展示
*.mp4 -> tempPicture/ -> tempPicture_reduce/ ->*.pptx

# Start
## 1.transfer.py(命令行参数方式,除videoPath必须指定之外，其他按需更改)
```
  --videoPath VIDEOPATH, -v VIDEOPATH
                        videoPath
  --time_interval TIME_INTERVAL, -t TIME_INTERVAL
                        time_interval
  --pictureFolder PICTUREFOLDER, -p PICTUREFOLDER
                        pictureFoder
  --reducePictureFolder REDUCEPICTUREFOLDER, -r REDUCEPICTUREFOLDER
                        reducePictureFolder
  --pptName PPTNAME, -m PPTNAME
                        pptName
  --threshold THRESHOLD, -th THRESHOLD
                        threshold
  --debug               debug
  --simple, -s          simple mode,run in the videopath
```
> 推荐使用方式
cd 目录/
mv *.mp4 目录/
python transfer.py -v *mp4 -s (自动配置临时截图文件夹路径，并在生成pptx后清除)

## 2.transfer.py(配置文件方式)
### 1) 配置config.ini
### 2) python transfer.py

## 3.preScript.ipynb(调用函数方式)
### 1) 选定mp4的路径
abc = video2pptx(".//英语课_15周.mp4")
### 2) 隔指定时间(秒s)截图到指定文件夹
abc.capFrame(".//tempPicture",60)
### 3) 利用相似度计算算法计算相邻两张图片相似度
similar_score = abc.calcSimilar(".\\tempPicture")
print(similar_score)
### 4) 指定阈值，将相似度小于阈值的图片挑选出来
abc.copyPictureBySimilar(0.93,".\\tempPicture",similar_score=similar_score)
### 5) 从指定文件夹所有图片生成pptx
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
