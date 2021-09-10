''' fileContentDoc.py  
   內容搜尋，子目錄下所有doc files( 舊版word )
   neoCaffe 2021-09-10 '''
import os,sys 
from win32com import client

#---------------------------------
#---  搜尋內容 .doc 舊版word檔 
# 參數 nPath 資料夾 / fTypes 要搜尋的類型
def findDoc( nPath, fileTypes, keyword):
    allDoc = []              # doc檔 之檔名
    result = []              # 符合之 檔名+行數
    contents = dict()        # key:找到的位置 / value:該段落文字
    f_tree = os.walk(nPath)  # 是一個generator
       
    for dirname,subdir,files in f_tree:
        # 一層一層向下
        print(f'本層資料夾總檔案數量: {len(files)}')
        docFiles = []    # 這一層的 doc files
        # 取得 符合之檔案(.doc檔)，存入 docFiles 串列中
        for file in files:  
            ext = file.split('.')[-1]
            if ext in filetypes:
                tmp = os.path.abspath(file)
                docFiles.append(tmp)
                allDoc.append(tmp)
        # docFiles >0 表示這一層有符合的檔案 (doc檔)
        if len(docFiles) > 0:
            # 逐一打開檔案
            for file in docFiles:
                try:
                    # 開啟舊版word doc 的方式
                    word = client.gencache.EnsureDispatch('Word.Application')
                    word.Visible = 0
                    word.DisplayAlerts = 0
                    doc = word.Documents.Open(file)
                    paras = doc.Paragraphs
                    n = 0               
                    for p in paras:
                        n += 1  # 段落數
                        # 如果該段落有此一keyword
                        if keyword in p.Range.Text:
                            tmp = f'檔案: {file} 第{n}段落找到< {keyword} >'
                            result.append(tmp)  # 存入 result list
                            # dict 新增一筆 key:找到的位置 / value:該段落文字
                            contents[tmp] = p.Range.Text
                    doc.Close()   # 關檔步驟，不可少        
                except:
                    err = f'檔案: {file} 讀取錯誤'
                    result.append(err)
                    
       
        
    return allDoc, result, contents  # 傳回
#---------------------------------------

#---------------------------------------
#--- 流程 主軸 -----
# 指定搜尋之目錄 (或者預設為當前目錄)
pathHere = os.getcwd() # 當前目錄位置
path = input('從哪個資料夾 開始搜尋 ? ') or pathHere
print(f'搜尋資料夾: {path} (含子目錄)')
tKey = input('keyword: ') or 'Python'

#--- 讀取舊版word .doc檔
filetypes = ['doc']    # 要篩選的檔案類型

dFile, dResult, dCont = findDoc( path, filetypes, tKey )    

''' 搜尋結果報告 '''
print(f'doc檔數量: {len(dFile)}')
if len( dResult ) != 0:
    print(f'keyword:< {tKey} > 出現次數: {len(dResult)}\n出現位置:')
    print(len(dCont.items()))
    for pos, text in dCont.items():
        print(pos,text)
        
''' 搜尋結果存檔  '''
f = open( pathHere+'/result_doc.txt','w',encoding='utf-8' )
print(f'doc檔數量: {len(dFile)}' ,file=f)

if len(dResult) != 0:
    print(f'keyword:< {tKey} > 出現次數: {len(dResult)}\n出現位置:',file=f)
    k = 0
    for pos, text in dCont.items():
        k += 1
        print(k,pos,text,file=f)
f.close()


