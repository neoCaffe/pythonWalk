''' fileContentdocx.py 
   內容搜尋，子目錄下所有docx files
   neoCaffe 2021-09-10 '''
import os,sys, docx

#---  搜尋內容 .docx word檔 
# 參數 nPath 資料夾 / fTypes 要搜尋的類型
def findDocx( nPath, fileTypes, keyword):
    allDocx = []            # docx檔 之檔名
    result  = []            # 符合之 檔名+行數
    contents = dict()       # key:找到的位置 / value:該段落文字
    f_tree = os.walk(nPath) # 是一個generator
       
    for dirname,subdir,files in f_tree:
        # 一層一層向下
        print(f'本層資料夾總檔案數量: {len(files)}')
        docxFiles = []    # 這一層的 docx files
        # 取得 符合之檔案(.doc檔)，存入 docFiles 串列中
        for file in files:  
            ext = file.split('.')[-1]
            if ext in filetypes:
                tmp = os.path.abspath(file)
                docxFiles.append(tmp)
                allDocx.append(tmp)
        # docxFiles >0 表示這一層有符合的檔案 (docx檔)
        if len(docxFiles) > 0:
            # 逐一打開檔案
            for file in docxFiles:
                try:
                    # 開啟word docx 的方式
                    doc = docx.Document(file)
                    #print(f'len(doc.paragraphs) {len(doc.paragraphs)}')
                    n = 0
                    for para in doc.paragraphs:
                        n += 1    # 段落數
                        # 如果該段落有此一keyword
                        if keyword in para.text:
                            tmp = f'檔案: {file} 第{n}行找到< {keyword} >'
                            result.append(tmp)   # 存入結果
                            # dict 新增一筆 key:找到的位置 / value:該段落文字
                            contents[tmp] = para.text
                except:
                    err = f'檔案: {file} 讀取錯誤'
                    result.append(err)
                                 
    return allDocx, result, contents  # 傳回
#---------------------------------------

#---------------------------------------
#--- 流程 主軸 -----
# 指定搜尋之目錄 (或者預設為當前目錄)
pathHere = os.getcwd() # 當前目錄位置

path = input('從哪個資料夾 開始搜尋 ? ') or pathHere
print(f'搜尋資料夾: {path} (含子目錄)')
tKey = input('keyword: ') or 'Python'

#--- 讀取word .docx檔
filetypes = ['docx']    # 要篩選的檔案類型

dFile, dResult, dCont = findDocx( path, filetypes, tKey )    

''' 搜尋結果報告 '''
print(f'docx檔數量: {len(dFile)}')
if len( dResult ) != 0:
    print(f'keyword:< {tKey} > 出現次數: {len(dResult)}\n出現位置:')
    #print(len(dCont.items()))
    for pos, text in dCont.items():
        print(pos,text)

''' 搜尋結果存檔  '''
f = open( pathHere+'/result_docx.txt','w',encoding='utf-8' )
print(f'docx檔數量: {len(dFile)}' ,file=f)

if len(dResult) != 0:
    print(f'keyword:< {tKey} > 出現次數: {len(dResult)}\n出現位置:',file=f)
    k = 0
    for pos, text in dCont.items():
        k += 1
        print(k,pos,text,file=f)
f.close()

