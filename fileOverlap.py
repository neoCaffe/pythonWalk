''' fileOverlap.py 檔案重覆之處理
    Author: neoCaffe  '''
import os, hashlib

#--- 找出重覆之檔案 nPath 資料夾 / fTypes 要搜尋的類型
def findOverlap( nPath, fTypes ):
    overlap = dict()  # 重覆之檔 key: hash / value: filePath
    imgFiles = []     # 所有圖檔之名稱
    f_tree = os.walk(nPath)
    
    for dirname,subdir,files in f_tree:
        # 取得 符合之檔案，存入 imgFiles 串列中
        for file in files:  
            ext = file.split('.')[-1]
            if ext in filetypes:
                tmp = dirname +'/'+file
                imgFiles.append(tmp)
      
        # 如果這一層有檔案 
        if len(files)>0:
            #--- 逐一檢查，如果新來之檔案 hash 不存在，則加入 
            for img in imgFiles:
                imghsh = hashlib.md5(open(img,'rb').read()).digest()
                if imghsh not in overlap:
                    overlap[imghsh] = os.path.abspath(img) 

    return imgFiles, overlap

#--- 流程 主軸 -----
# 指定搜尋之目錄 (或者預設為當前目錄)
pathHere = os.getcwd() # 當前目錄位置
path = input('從哪個資料夾 開始搜尋 ? ') or pathHere
print(f'搜尋資料夾: {path} (包含子目錄)圖檔')

# 要篩選的檔案類型
filetypes = ['jpg', 'png', 'bmp', 'jpeg']  
iFile, ioverlap = findOverlap( path, filetypes )

print(f'圖檔數量: {len(iFile)} 重覆者: {len(ioverlap)}')

if len(ioverlap)!=0:
    print("找到下列重覆的檔案：")
    hshList = list(ioverlap.keys())
    fList = list(ioverlap.values())
    for i in range(len(ioverlap)):
        print(hshList[i],fList[i])
   
# 把結果存檔
f = open( pathHere+'\overlap.txt','w',encoding='utf-8' )
print(f'圖檔數量: {len(iFile)} 重覆者: {len(ioverlap)}',file=f)
print("找到下列重覆的檔案：",file=f)
for i in range(len(fList)):
    print(f'{hshList[i]}    {fList[i]}',file=f)
f.close()

