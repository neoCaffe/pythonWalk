''' fileContentTxt.py  
   內容搜尋，子目錄下所有 txt py files
   neoCaffe 2021-09-10 '''

import os

#---  搜尋內容 txt 
# 參數 nPath 資料夾 / fTypes 要搜尋的類型
def findTxt( nPath, fTypes , keyword):
    allTxt = []              # 文字檔 之檔名
    result = []              # 符合之 檔名+行數
    contents = dict()        # key:找到的位置 / value:該段落文字
    f_tree = os.walk(nPath)  # 是一個generator
    # os.walk 傳回的是一個 generator
    #print(f'return a generator: {type(f_tree)}')
    
    for dirname,subdir,files in f_tree:
        # 一層一層向下
        print(f'本層資料夾總檔案數量: {len(files)}')
        txtFiles = []    # 這一層的 txt files
        # 取得 符合之檔案(文字檔)，存入 txtFiles 串列中
        for file in files:  
            ext = file.split('.')[-1]
            if ext in filetypes:
                tmp = os.path.abspath(file)
                txtFiles.append(tmp)
                allTxt.append(tmp)
        # txtFiles >0 表示這一層有符合的檔案 (文字檔)
        if len(txtFiles) > 0:
            # 逐一打開檔案
            for file in txtFiles:
                try:
                    # 開啟文字檔的方式
                    f = open(file,'r',encoding='utf-8')
                    # 讀取內容
                    lines = f.readlines()
                    f.close()
                    n = 0
                    # 逐行檢查
                    for line in lines:
                        n += 1
                        if keyword in line:
                            tmp = f'檔案: {file} 第{n}行找到< {keyword} >'
                            result.append(tmp)  # 存入結果
                            # dict 新增一筆 key:找到的位置 / value:該行文字
                            contents[tmp] = line
                except:
                    err = f'檔案: {file} 讀取錯誤'
                    result.append(err)
                                 
    return allTxt, result, contents   # 傳回
#---------------------------------------

#---------------------------------------
#--- 流程 主軸 -----
# 指定搜尋之目錄 (或者預設為當前目錄)
pathHere = os.getcwd() # 當前目錄位置
path = input('從哪個資料夾 開始搜尋 ? ') or pathHere
print(f'搜尋資料夾: {path} (含子目錄)')
tKey = input('keyword: ') or 'Python'

filetypes = ['txt', 'py']  # 要篩選的檔案類型

tFile, tResult, tCont = findTxt( path, filetypes, tKey )

''' 搜尋結果報告 '''
print(f'文字檔數量: {len(tFile)}')
if len(tResult) != 0:
    print(f'keyword:< {tKey} > 出現次數: {len(tResult)}\n出現位置:')
    for pos, text in tCont.items():
        print(pos,text)
 
''' 搜尋結果存檔  '''
f = open( pathHere+'/result_txt.txt','w',encoding='utf-8' )
print(f'文字檔數量: {len(tFile)}' ,file=f)

if len(tResult) != 0:
    print(f'keyword:< {tKey} > 出現次數: {len(tResult)}\n出現位置:',file=f)
    k = 0
    for pos, text in tCont.items():
        k += 1
        print(k,pos,text,file=f)

f.close()

