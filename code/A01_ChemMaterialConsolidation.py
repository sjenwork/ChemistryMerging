import pandas as pd
import os, re, copy
import numpy as np
import Levenshtein
import datetime
import random   #用來測試
import sys

pd.options.mode.chained_assignment = None  # default='warn'

#-----------------------------------------------------------------------------------------
# 共用函數          開始
#-----------------------------------------------------------------------------------------

def show(massage, level=0, sign='>', nlb=False, nla=False):
    nspace = {0:6, 1:7, 2:8, 3:9, 4:10, 5:11, 6:12, 7:13}
    nsign  = {0:2, 1:2, 2:3, 3:3, 4:3 , 4:3 , 5:4 , 7:4}
    if nla and nlb:
        print('\n' + ' '*nspace[level] + f'{sign}'*nsign[level]  + f' {massage}' + '\n')
    elif nla:
        print(' '*nspace[level] + f'{sign}'*nsign[level]  + f' {massage}' + '\n')
    elif nlb:
        print('\n' + ' '*nspace[level] + f'{sign}'*nsign[level]  + f' {massage}')
    else:
        print(' '*nspace[level] + f'{sign}'*nsign[level]  + f' {massage}')


def AccuComp(a, b):
    res = [1 for i in a if i in b]
    if res == []: 
        res=0
    else:
        res=1
    return res


def FuzzyComp(a, b):
    if type(b) == pd.Series:
        b = b.iloc[0]
    if type(a) == str:
        a = a.split(';')
        b = b.split(';')
    x = []
    for i in b:
        for j in a:
            if i == '' or j == '':
                tmp = 0
            else:
                tmp = Levenshtein.ratio(j,i)
            x.append(tmp)
    res = max(x)
    return res

#-----------------------------------------------------------------------------------------
# 共用函數          結束
#-----------------------------------------------------------------------------------------







#-----------------------------------------------------------------------------------------
# 檔案路徑位置      開始
#-----------------------------------------------------------------------------------------
class Path():
    def __init__(self, basePath, inpFile=None, refFile=None):
        if refFile is None:
            refFile = 'data/指引表.xlsx'
        if inpFile is None:
            inpFile = 'data/202001-03已整併.xls'
        self.basePath = basePath
        self.refFile = refFile
        self.inpFile = inpFile
        self._proPath()
        self._path()

    def _proPath(self):
        self.BasePath = self.basePath 
        #self.BasePath = '/Users/jen/GD_simenvi/SimEnvi/Project/109E10_ChemCloud/01.CodingWork/A03_Merge/'

        self.inpPath   = os.path.join(self.BasePath, os.path.dirname(self.inpFile)   )
        self.refPath = os.path.join(self.BasePath, os.path.dirname(self.refFile)   )
        self.inpName   = os.path.basename(self.inpFile)    
        self.refName = os.path.basename(self.refFile)  

    def _path(self):
        # path for preProsess
        self.fileName_statistic1 = os.path.join(self.inpPath, f'statistic1.xlsx')                       # 不需整併統計
        self.fileName_statistic2 = os.path.join(self.inpPath, f'statistic2.xlsx')                       # 各標準數量統計 
        self.fileName_Final_noco = os.path.join(self.inpPath, f'Final_不需整併.xlsx')

        # path for 指引表
        self.FN_rawRe = os.path.join(self.refPath, self.refName)
        self.FN_resRE = os.path.join(self.refPath, self.refName.replace('.xlsx', '_整理.xls'))
        self.FN_tstRe = os.path.join(self.refPath, self.refName.replace('.xlsx', '_整理測試用.xls'))

        # 移除關鍵字
        self.FN_rmkws = os.path.join(self.inpPath, 'keyword')  

        # path 整併
        self.FN_rawDa = os.path.join(self.inpPath, self.inpName)                                        # 原始檔案
        self.FN_wiCAS = os.path.join(self.inpPath, self.inpName.replace('.xlsx','_withCASmatched.xlsx'))  # 整理完成的資料中，有CAS且可match的資料        => 最終的檔案
        self.FN_proce = os.path.join(self.inpPath, self.inpName.replace('.xlsx','_nonProcessFile.xlsx'))  # 原始檔案 - 已整併，暫存用，完成即刪。
        self.FN_noCAS = os.path.join(self.inpPath, self.inpName.replace('.xlsx','_woutCASmatched.csv' ))  # 剩下的資料                                    => 最終的檔案
        self.FN_check = os.path.join(self.inpPath, self.inpName.replace('.xlsx','_fileForConfirm.csv' ))  # 剩下的資料，若前一步驟完全處理完成，將沒有資料。
        self.FN_bckup = os.path.join(self.inpPath, self.inpName.replace('.xlsx','_fileForBackup_.csv'))  # 即時輸出檔案。備份用，完成即刪。
        self.FN_postP = os.path.join(self.inpPath, self.inpName.replace('.xlsx','_Final_完成整併.xlsx'))  # 處理完成後之後處理

        #self.fileName_mrg_file_test = self.FN_rawDa.replace('.xls', '_test.xls') 

#-----------------------------------------------------------------------------------------
# 檔案路徑位置      結束
#-----------------------------------------------------------------------------------------




#-----------------------------------------------------------------------------------------
# 指引表整理        開始
#-----------------------------------------------------------------------------------------
class referenceTable(Path):
    def __init__(self, basePath, refFile=None, test_run=False):
        super(referenceTable, self).__init__(basePath=basePath, refFile=refFile)
        self.test_run = test_run
        self._read()


    def _colNewName(self):
        nameList = { 'MatchNo(指引表編號)'                  : 'MatchNo'
                   , 'CASNoMatch(對應的Cas No.)'            : 'CAS_m'
                   , 'ChemiChnNameMatch(對應的中文名稱)'    : 'Cname_m'
                   , 'ChemiEngNameMatch(對應的英文名稱)'    : 'Ename_m'
                   , 'ChemiChnAliases(中文別名)'            : 'Cname-a_m'
                   , 'ChemiEngAliases(英文別名)'            : 'Ename-a_m'
                   }
        return nameList

    def _read(self):
        show(f'==============================', 0, sign=' ')
        show(f'Reading Reference Table', 0)
        ## 讀取原始指引表檔案
        if not os.path.isfile(self.FN_resRE):
            show(f'Reading: {self.FN_rawRe}')
            data = pd.read_excel(self.FN_rawRe,  header=[0])
            data = self._processRawData(data)
            data.to_excel(self.FN_resRE)
        else:
            if self.test_run:
                if os.path.isfile(self.FN_tstRe):
                    show(f'{self.FN_tstRe} exists. Reading it!!', 0)
                    data = pd.read_excel(self.FN_tstRe, header=[0], index_col=[0])
                else:
                    show(f'{self.FN_tstRe} does not exists. Creating it!!', 0)
                    data = pd.read_excel(self.FN_resRE , header=[0], index_col=[0])
                    data[:self.test_run].to_excel(self.FN_tstRe)
            else:
                show(f'{self.FN_resRE} exists. Reading it!!')
                data = pd.read_excel(self.FN_resRE,  header=[0], index_col=[0])


        self.reference = data

    def _processRawData(self, data):
        data.rename(self._colNewName(), axis=1, inplace=True)

        data['CAS_m'] = data['CAS_m'].apply(lambda i: i.strftime('%Y-%m-%-d') if type(i) == datetime.datetime else i)
        data['CAS_m'] = data['CAS_m'].apply(lambda i: re.sub('\s+','',i))

        # 將數個CAS對應至同一種matchNo的資料分成不同資料
        matchno = pd.DataFrame(data['CAS_m'].str.split(';').to_list())
        matchno.columns = [f'CAS_m_{i}' for i in range(matchno.columns.size)]
        data = pd.concat([matchno, data], axis=1)
        data.drop('CAS_m', axis=1, inplace=True)
        data = pd.wide_to_long(data, ['CAS_m_'], i='MatchNo', j='casCnt')
        data.dropna(subset=['CAS_m_'], inplace=True)
        data.index = data.index.droplevel(1)
        data.sort_index(inplace=True)
        data.rename({'CAS_m_': 'CAS_m'}, axis=1, inplace=True)
        data = data[list(self._colNewName().values())[1:]]

        # 考慮不同名稱(合併中英文名)，合併為同一個欄位
        data.fillna('', inplace=True)   # 指引表中有一筆資料是空白
        data['name_m_all'] = data['Cname_m'].astype(str) + ';' + data['Cname-a_m'].astype(str) + ';' \
                          +  data['Ename_m'].astype(str) + ';' + data['Ename-a_m'].astype(str)
        # 移除不需要的欄位
        data.drop(['Cname_m', 'Cname-a_m', 'Ename_m', 'Ename-a_m'], axis=1, inplace=True)
        # 把多的空白刪除

        data['name_m_all'] = data['name_m_all'].str.strip()
        data['name_m_all'] = data['name_m_all'].apply(lambda i: re.sub(';\s+', ';', i))
        data['name_m_all'] = data['name_m_all'].apply(lambda i: re.sub('\s+;', ';', i))
        # 把重複的分號移除，僅留下一個分號
        data['name_m_all'] = data['name_m_all'].apply(lambda i: re.sub('[;]+',';',i))
        # 移除句首句尾分號，避免錯誤
        data['name_m_all'] = data['name_m_all'].apply(lambda i: re.sub(';$','',i))
        data['name_m_all'] = data['name_m_all'].apply(lambda i: re.sub('^;','',i))
        # 將所有的英文變成小寫
        data['name_m_all'] = data['name_m_all'].str.lower()
        data['name_m_all'] = data['name_m_all'].apply(lambda i: re.sub("'", "", i))

        
        #data['name_m_all'] = data['name_m_all'].apply(lambda i: i.split(';'))

        # 此部分需要修正方法
        data.drop_duplicates(subset='CAS_m', inplace=True)
        return data


#-----------------------------------------------------------------------------------------
# 指引表整理        結束
#-----------------------------------------------------------------------------------------



#-----------------------------------------------------------------------------------------
# 前處理            開始 
#-----------------------------------------------------------------------------------------
class preProcessor(Path):
    def __init__(self, inpFile, basePath, removeKeyWord='香料'):
        super(preProcessor, self).__init__(basePath=basePath, inpFile=inpFile)
        self.removeKeyWord = removeKeyWord 
        self._filePath()
        self._main()


    def _filePath(self):
        basePath = os.path.dirname(self.FN_rawDa)
        outName = os.path.join(basePath, f'subdata_DATE.xls')
        self.outName = outName

    # ==========================================================
    # 主程式
    # ==========================================================
    def _main(self):
        self._other()
        self._read()
        self._process()
        #self._division()

    # ==========================================================
    # 雜項工作
    # 1. 產生"年月"時間序列
    # ==========================================================
    def _other(self):
        timeSeries = pd.date_range('2017-05', '2020-06',freq='m')
        self.timeSeries = timeSeries


    # ==========================================================
    # Reading file
    # ==========================================================
    def _read(self):
        show('Reading data', 0, nlb=True)
        df = pd.read_excel(self.FN_rawDa)
        self.df = df

    # ==========================================================
    # Processing data
    # ==========================================================
    def _process(self):
        self._processFormat()
        self._statistic()
        #self._removeUseless()

    def _processFormat(self):
        show('Processing data', 0)
        df = self.df
        #df.columns = ['ChemiEngName', 'ChemiChnName', 'CASNo', 'UpdateDate', 'TransId', 'PrimaryTableName', 'IsMatched', 'IsMatchAdd', 'MatchNo']  # for 0521
        df.columns = ['ChemicalChnName', 'ChemicalEngName', 'CASNo', 'PrimaryTableName', 'MatchNo', 'Temp', 'TransId']                              # for 0526
        df.fillna('', inplace=True)
        df = df.applymap(lambda i: re.sub('^-*$', '', i) if type(i) == str else i)

        self.df = df 
        
    def _removeUseless(self):
        df = self.df
        data = df['ChemiChnName']
        idx = data.str.contains(self.removeKeyWord)
        df_useless = df[idx]
        df = df[~idx]
        self.df = df
        self.df_useless = df_useless
        show(f'將不需要整併的化學物質輸出至 {self.fileName_Final_noco}')
        df_useless.to_excel(self.fileName_Final_noco)

    def _statistic(self):
        # ================================================== #
        # Statistic                                          #
        #           計算出每個月份的不需要整併的品項數量     #
        #           目前僅使用"香料"作為關鍵字               #
        #           未來若考慮新增其他關鍵字，必須修改code。 #
        # ================================================== #
        df = self.df
        df = df.set_index('UpdateDate')
        statKeyWord = []
        columns=['Only CAS', 'Only ChiName', 'Only EngName', 'CAS & ChiName', 'CAS & EngName', 'ChiName & EngName', 'All standard']
        statDetStan = pd.DataFrame()
        for itime in self.timeSeries[:]:
            # 找出特定年月資料
            year = itime.year
            month= itime.month
            ind  = (df.index.year == year) & (df.index.month == month)
            data = df.iloc[ind]
            # 找出中文名稱中含有特定關鍵字的index
            idxDataKW = data['ChemiChnName'].str.contains(self.removeKeyWord)
            # 計算資料總數
            lenData = len(data)
            # 含有關鍵字的比數
            lenKw   = idxDataKW.sum()
            statKeyWord.append([lenKw, lenData])

            # 篩選出需要整併的列表
            data2 = data[~idxDataKW]
            tmp = pd.DataFrame(columns=columns, index=[f'{year}-{month}'])
            # 所有三種標準都有值
            tmp['All standard'     ]    = [len(data2[(data2.ChemiChnName!='') & (data2.ChemiEngName!='') & (data2.CASNo!='')])]
            tmp['ChiName & EngName']    = [len(data2[(data2.ChemiChnName!='') & (data2.ChemiEngName!='') & (data2.CASNo=='')])]
            tmp['CAS & EngName'    ]    = [len(data2[(data2.ChemiChnName!='') & (data2.ChemiEngName=='') & (data2.CASNo!='')])]
            tmp['CAS & ChiName'    ]    = [len(data2[(data2.ChemiChnName=='') & (data2.ChemiEngName!='') & (data2.CASNo!='')])]
            tmp['Only CAS'         ]    = [len(data2[(data2.ChemiChnName=='') & (data2.ChemiEngName=='') & (data2.CASNo!='')])]
            tmp['Only ChiName'     ]    = [len(data2[(data2.ChemiChnName!='') & (data2.ChemiEngName=='') & (data2.CASNo=='')])]
            tmp['Only EngName'     ]    = [len(data2[(data2.ChemiChnName=='') & (data2.ChemiEngName!='') & (data2.CASNo=='')])]
            statDetStan = statDetStan.append(tmp)



        # 關鍵字統計
        statKeyWord                     = pd.DataFrame(statKeyWord, columns=[self.removeKeyWord, 'Total'])
        statKeyWord.index               = self.timeSeries
        statKeyWord.fillna('')
        statKeyWord_byYear              = pd.pivot_table(statKeyWord, index=statKeyWord.index.year, aggfunc='sum')

        statKeyWord.loc['Total']        = statKeyWord.sum()
        statKeyWord['Ratio']            = statKeyWord[self.removeKeyWord]/statKeyWord['Total']
        statKeyWord_byYear.loc['Total'] = statKeyWord_byYear.sum()
        statKeyWord_byYear['Ratio']     = statKeyWord_byYear[self.removeKeyWord]/statKeyWord_byYear['Total']

        # 標準比例統計
        statDetStan.index               = self.timeSeries
        statDetStan_byYear              = pd.pivot_table(statDetStan, index=statDetStan.index.year, aggfunc='sum')

        statDetStan.loc['Total']        = statDetStan.sum()
        statDetStan['Total']            = statDetStan.sum(axis=1)
        statDetStan_byYear['Total']     = statDetStan_byYear.sum(axis=1)
        statDetStan_byYear.loc['Total'] = statDetStan_byYear.sum()
        statDetStanRatio                = statDetStan.div(statDetStan['Total'], axis=0)
        statDetStanRatio_byYear         = statDetStan_byYear.div(statDetStan_byYear['Total'], axis=0)
        statDetStan                     = pd.concat([statDetStan, statDetStanRatio], axis=1)
        statDetStan_byYear              = pd.concat([statDetStan_byYear, statDetStanRatio_byYear], axis=1)


        self.statKeyWord                = statKeyWord
        self.statKeyWord_byYear         = statKeyWord_byYear
        self.statDetStan                = statDetStan
        self.statDetStan_byYear         = statDetStan_byYear

        statKeyWord.to_excel(self.fileName_statistic1)
        with pd.ExcelWriter(self.fileName_statistic2) as writer:
            self.statDetStan.to_excel       (writer, sheet_name='by_month')
            self.statDetStan_byYear.to_excel(writer, sheet_name='by_year')


    def _division(self, datePerFile=5000):
        # ==========================================================
        # File division
        # ==========================================================
        df = self.df
        df = df.set_index('UpdateDate')
        for itime in self.timeSeries[:]:
            year = itime.year
            month= itime.month
            ind  = (df.index.year == year) & (df.index.month == month)
            data = df.iloc[ind]

            time = itime.strftime('%Y-%m')
            fileName = self.outName.replace('DATE', time)
            #show(f'writing data to {fileName}', 1)
            show(f'＊{time}: totol {len(data)} chemical material', 3, sign='', nlb=True)

            count = 0
            while True:
                s, e = count, count+datePerFile
                subdata = data.iloc[s:e]
                subfileName = fileName.replace('.xls', f'_{s//datePerFile}.xls')
                show(f'寫入資料： {subfileName}', 4)
                subdata.to_excel(subfileName, index=False)

                count+=datePerFile
                if count>len(data): break
                
        self.data = data
#-----------------------------------------------------------------------------------------
# 前處理            結束 
#-----------------------------------------------------------------------------------------



#-----------------------------------------------------------------------------------------
# 整併檔案          開始
#-----------------------------------------------------------------------------------------
class mergeChem(Path):
    def __init__(self, inpFile, basePath):
        super(mergeChem, self).__init__(inpFile=inpFile, basePath=basePath)
        self._MAIN()

    def _MAIN(self):
        show(f'==============================', sign=' ')
        show(f'整併流程開始。')
        if os.path.isfile(self.FN_proce):
            show('Part I: 已經完成整併"有CAS No"的數據，讀取未整併之檔案', 1)
            show(f'Reading: {self.FN_proce}', 2)
            data = pd.read_excel(self.FN_proce, index_col=[0])
            data.fillna('', inplace=True)
        else:
            show(f'Part I: 讀取原始資料', 1)
            show(f'Step 1: Reading: {self.FN_rawDa}', 2)
            data = self._read()

            show(f'Step 2: 資料整理與清洗', 2)
            data = self._dataOrganize(data)

        show(f'Part II: 資料整併', 1, nlb=True)
        self._dataMerging(data)
        #self._calStatistic()
        #self._printResult()

    def _printResult(self):
        show(f'==============================', sign=' ')
        show(f'統計結果：')
        print(self.statistic.T)
        show(f'==============================', sign=' ')


    ''' -----------------------------------------------------------
        >>>>  IO 
        ----------------------------------------------------------- '''
    def _colNewName(self):
        nameList = {  'CASNo'               : 'CAS'
                    , 'ChemiChnName'        : 'Cname'
                    , 'ChemiEngName'        : 'Ename'
#                    , 'ChemicalFormula'     : 'Formula'
#                    , 'CCCCode'             : 'CCCcode'
                    }
        self.nameList = nameList

    def _read(self):
        self._colNewName()
        data = pd.read_excel(self.FN_rawDa)
#        data = data[['CASNo', 'ChemicalChnName', 'ChemicalEngName', 'ChemicalFormula', 'CCCCode' ]]
        data = data[['CASNo', 'ChemiChnName', 'ChemiEngName' ]]
        data.rename(self.nameList, axis=1, inplace=True)
        show(f'finishing selecting specific column')

        data.index.name = 'index'
        self.dataRaw = copy.deepcopy(data)
        return data

    def _write(self, data, fileName):
        show(f'writing {fileName}', 4)
        data.to_excel(fileName)
        #self.data = data

    ''' -----------------------------------------------------------
        >>>>  資料整理
        ----------------------------------------------------------- '''

    def _readKeyword(self):
        with open(self.FN_rmkws) as f:
            kw = f.readlines()
        kw = [i.strip() for i in kw if i[0] != '#']
        kw = '|'.join(kw)
        self.keyword = kw
        return kw

    def _dataOrganize(self, data):
        data.fillna('', inplace=True)
        data.replace('-', '', inplace=True)

        # 讀取移除的中文關鍵字
        kw = self._readKeyword()
        ind = data['Cname'].str.contains(kw)
        ind = ind.fillna(False)
        data = data[~ind]

        # =====================================================================
        # Main processes

        # 將無意義的符號移除，將括號中的字詞分離，分為不同欄位。
        show(f'將無意義的符號移除，將括號中的字詞分離，分為不同欄位', 2)
        data = self._organizeCAS(data)                                      
        data = self._organizeChiName(data) 
        data = self._organizeEngName(data) 

        self.data = data
        return data


    def _organizeCAS(self, data):
        # 正規化
        casno = data['CAS'].astype(str)
        casno = casno.str.replace('00:00:00', '')
        casno = casno.apply(lambda i: i if re.compile('.*-[0-9][0-9]-.*').match(i) else '')
        casno = casno.str.replace('^0+', '')
        casno = casno.apply(lambda i: re.sub('[^0-9-]','',i))
        casno = casno.apply(lambda i: re.sub('^-*$', '', i))
        casno = casno.str.strip()
        casno = casno.str.replace('0+(?=\d$)', '')
        casno = casno.str.replace('-+','-')
        casno = casno.str.replace('–','-')
        #casno.replace('', np.nan, inplace=True)
        #casno.dropna(inplace=True)
        data['CAS-Rev'] = casno
        return data

    def _organizeChiName(self, data):
        cname = data['Cname'].astype(str)
        cname = cname.apply(lambda i: re.sub('\s+', '', i))
        cname = cname.replace('脂', '酯')
        # 脂 or 酯
        # 光阻劑 光阻液 上光
        data2 = pd.DataFrame()
        # 將括號中的名稱視為獨立名稱，以供後續比對
        # 可能的bug:  括號只有單邊、括號中有括號
        # 括號 ()
        data2['Cname-2'] = cname.apply(lambda i: i[i.find("(") : i.find(")")+1])
        #data2['Cname-2'] = data2['Cname-2'].apply(lambda i: re.sub('\(|\)', '', i))
        # 括號 （）
        data2['Cname-3'] = cname.apply(lambda i: i[i.find("（") : i.find("）")+1])
        #data2['Cname-3'] = data2['Cname-3'].apply(lambda i: re.sub('\（|\）', '', i))
        # 括號 []
        data2['Cname-4'] = cname.apply(lambda i: i[i.find("[") : i.find("]")+1])
        #data2['Cname-4'] = data2['Cname-4'].apply(lambda i: re.sub('\[|\]', '', i))
        # 括號 <>

        data2['Cname-5'] = cname.str.replace('\(.*\)', '')
        data2['Cname-1'] = cname
        #data2['Cname-1'] = data2[['Cname-1']]
        self.data2 = data2
        # 整合所有可能的名稱
        cname_m = data2['Cname-1'] + ';' + data2['Cname-2'] + ';' + data2['Cname-3'] + ';' + data2['Cname-4'] + ';' + data2['Cname-5']
        # 將可能的分隔符號取代為;
        cname_m = cname_m.apply(lambda i: re.sub('\(|\（|\)|\）|\[|\]', '', i))
        cname_m = cname_m.apply(lambda i: re.sub('；|_|/', ';', i))
        # 移除多的分號
        cname_m = cname_m.apply(lambda i: re.sub(';+',';', i))
        # 移除字首及字尾分號
        cname_m = cname_m.apply(lambda i: re.sub(';$','', i))
        cname_m = cname_m.apply(lambda i: re.sub('^;','', i))
        # 所有英文變成小寫
        cname_m = cname_m.str.lower()


        data['Cname-Rev'] = cname_m


        # ============== 後續的部分並非中文修改，而是全部修改

        #data['Cname-Rev'] = data['Cname-Rev'].str.lower()
        # 移除所有百分比相關字串，如"90%'
        data = data.applymap(lambda i: re.sub('[0-9]+%', '', i) if type(i) == str else i)
        # 移除所有#
        data.replace('#', '', inplace=True)
        data = data.applymap(lambda i: re.sub(';$','',i) if type(i) == str else i)
        # ==================================================

        return data

    def _organizeEngName(self, data):
        ename = data['Ename'].astype(str)
        ename = ename.apply(lambda i: re.sub('\s+', ' ', i))
        ename = ename.apply(lambda i: i.lower())
        data['Ename-Rev'] = ename
        return data



    ''' -----------------------------------------------------------
        >>>>  資料整併
        ----------------------------------------------------------- '''
    def _dataMerging(self, data):

        if not os.path.isfile(self.FN_proce):
            show(f'將"有CAS No"的資料挑選出來進行整併', 2)
            data = self._byCAS(data)
            show(f'完成整併"有CAS No"且"可對應到指引表"的化學物種，並給定Match No', 2)
            self._write(data        , fileName = self.FN_proce)
            self._write(self.data_wiCAS, fileName = self.FN_wiCAS)

        show(f'將剩餘("沒有CAS no"或"有CAS no但無法對應"的資料)的化學物種逐一比對', 2)
        data = self._oneByOne(data)
        show(f'==============================', sign=' ')
        #self._write(data            , fileName = self.FN_check)
        #self._write(self.data_noCAS , fileName = self.FN_noCAS)

        self.data = data

    # ------------------------------------------------------------
    #               Method 1
    def _byCAS(self, data):
        casno       = data['CAS-Rev']
        indWithCAS  = casno[casno!='']
        #ref_m2c    = r.reference['CAS_m'].to_dict()        # 轉成dict要注意是否有一個match no可以對應到數個cas no
        ref_c2m     = r.reference.reset_index().set_index('CAS_m')['MatchNo'].to_dict()
        CASmatch    = indWithCAS.map(ref_c2m)
        CASmatch    = CASmatch.dropna().astype(int)
        indCASmatch = list(CASmatch.index)
        indCASunmat = list(indWithCAS.drop(indCASmatch).index)

        data_wiCAS                 = data.loc[indCASmatch]
        data_wiCAS['MatchNoRe']    = CASmatch
        data_wiCAS['ChiScore']     = self._CASmatched_chineseScore(data_wiCAS, r.reference)
        data_wiCAS['EngScore']     = self._CASmatched_englishScore(data_wiCAS, r.reference)
        data_wiCAS['CASscore']     = 1
        data_wiCAS['MatchResult']  = '完成整併' 
        #data_wiCAS.sort_values(by=['ChiScore', 'EngScore'], ascending=False, inplace=True)
    
        #self.ref_m2c       = ref_m2c
        self.ref_c2m        = ref_c2m
        self.indWithCAS     = indWithCAS  
        self.CASmatch       = CASmatch
        self.indCASmatch    = indCASmatch
        self.indCASunmat    = indCASunmat
       

        self.data_wiCAS    =   data_wiCAS

        data.drop(indCASmatch, inplace=True)

        return data

    def _CASmatched_chineseScore(self, data, reference):
        show('Processing Chinese name',3)
        reference = reference.reset_index().set_index('CAS_m')
        reference = reference.loc[data['CAS-Rev'].to_list()]
        reference = reference.reset_index().set_index('MatchNo')
        data = data.set_index('MatchNoRe')
        chiName = pd.concat([data[['Cname-Rev']], reference[['name_m_all']]], axis=1)
        chiName.fillna('', inplace=True)
        score = [FuzzyComp(*j) for i,j in chiName.iterrows()]

        self.reference = reference
        self.chiName = chiName
        return score

    def _CASmatched_englishScore(self, data, reference):
        show('Processing English name', 3)
        reference = reference.reset_index().set_index('CAS_m')
        reference = reference.loc[data['CAS-Rev'].to_list()]
        reference = reference.reset_index().set_index('MatchNo')
        data = data.set_index('MatchNoRe')
        chiName = pd.concat([data[['Ename-Rev']], reference[['name_m_all']]], axis=1)
        chiName.fillna('', inplace=True)
        score = [FuzzyComp(*j) for i,j in chiName.iterrows()]

        self.reference = reference
        self.chiName = chiName
        return score
    #               Method 1 finished
    # ------------------------------------------------------------

    def _oneByOne(self, data):
        '''
        可能的bug：如果中文名裡面是英文，也可以對應到其他英文名的部分
        '''
        reference = r.reference.fillna('')
        #reference = reference.applymap(lambda i: re.sub(';$', '', i))
        #reference = reference.applymap(lambda i: re.sub(';$', '', i) if type(i) == str else i)
        reference['name_m_all'] = reference['name_m_all'].apply(lambda i: i.split(';'))
        dataPro = data[['Cname-Rev', 'Ename-Rev', 'CAS-Rev']].iloc[:40000]

        if os.path.isfile(self.FN_noCAS):
            show(f'"無CAS"之數據整併檔存在。讀取：{self.FN_noCAS}', 3) 
            data_noCAS = pd.read_csv(self.FN_noCAS, index_col=[0])
        else:
            show(f'"無CAS"之數據整併檔不存在。重新建立：{self.FN_noCAS}', 3) 
            data_noCAS = pd.DataFrame()
            self._writeCSVbackup()

        # 迴圈開始
        #randomList = random.sample(list(dataPro.index), 5)
        for datatmp in dataPro.iterrows():
            
            index ,(chiName, engName, CAS) = datatmp

            if index in data_noCAS.index:
                show(f'此資料("{index}")已經整併完成，跳過。', 4)
                continue

            # -----------------------------------------------------------
            # 測試用
            #if index not in randomList: continue
            #if idata >= 100: continue
            #if index not in  [5957, 5998, 7848, 5885, 5882, 283]: continue             # 888, 945
            #if index not in [45072]: continue                                           # 48644
            #if index not in [5478]: continue                                           # 48644, 5478
            # -----------------------------------------------------------
            show(f'Looping for Chemisty species: \tindex="{index}",\t CAS="{CAS}",\t ChineseName="{chiName}",\t EnglishName="{engName}"', 3, nlb=True)


            # 將中文名及英文名用分號(;)區隔: str => List
            chiName = chiName.split(';')
            engName = engName.split(';')

            # --------------------------------
            # 進行比對：
            tmpData = data.loc[[index]]

    # **********************************************************************************************
    # 中文比對                  
    # **********************************************************************************************
        # 精準比對
            show(f'o Accurate comparison for chinese Name ', 4, sign=' ')
            chiRes = reference.apply(lambda i: AccuComp(chiName, i['name_m_all'] ), axis=1)
            chiRes = chiRes[chiRes==1]
            chiRes = chiRes.drop_duplicates()
        # 模糊比對
            if chiRes.empty:
                show(f'o not Found by Accurate comparison !!', 4, sign=' ')
                show(f'o Fuzzy comparison for chinese Name ', 4, sign=' ')
                chiRes = reference.apply(lambda i: FuzzyComp(chiName, i['name_m_all'] ), axis=1)
                chiRes = chiRes.drop_duplicates()
                chiRes = chiRes[chiRes>.5]
                chiRes = chiRes.sort_values(ascending=False)[:5]

            chiResForWrite = chiRes.apply(lambda i: f'{i:.2f}')
    # **********************************************************************************************

    # **********************************************************************************************
    # 英文比對
    # **********************************************************************************************
        # 精準比對
            show(f'o Accurate comparison for english Name', 4, sign=' ')
            engRes = reference.apply(lambda i: AccuComp(engName, i['name_m_all'] ), axis=1)
            engRes = engRes[engRes==1]
            engRes = engRes.drop_duplicates()
        # 模糊比對
            if engRes.empty:
                show(f'o not Found by Accurate comparison !!', 4, sign=' ')
                show(f'o Fuzzy comparison for english Name ', 4, sign=' ')
                engRes = reference.apply(lambda i: FuzzyComp(engName, i['name_m_all'] ), axis=1)
                engRes = engRes.drop_duplicates()
                engRes = engRes[engRes>.5]
                engRes = engRes.sort_values(ascending=False)[:5]

            engResForWrite = engRes.apply(lambda i: f'{i:.2f}')
    # **********************************************************************************************

    # **********************************************************************************************
    # 依據精準比對與模糊比對之結果，決定整併結果
    # **********************************************************************************************
            #   給定初始值
            score_Eng   = ''
            score_Chi   = ''
            matchno     = ''
            matchResult = '未處理'

            if 1 in engRes.values and 1 in chiRes.values:

                show('condition 1: 中文英文名稱皆完全符合', 4)

                # 檢查是否有相同的 match no
                rep_match = [i for i in chiRes.index if i in engRes.index]
                if rep_match != []:
                    show('condition 1-1: 中文對應的match no與英文對應的match no部分一致，找出共同且唯一的match no。', 5)
                    scoreTmp = pd.DataFrame()
                    for imatch in rep_match:
                        scoreTmpChi = chiRes[imatch]
                        scoreTmpEng = engRes[imatch]
                        scoreTmp[imatch] = [scoreTmpChi + scoreTmpEng]
                    matchno = scoreTmp.T[0].idxmax()
                    score_Chi = chiRes[matchno]
                    score_Eng = engRes[matchno]
                    matchResult = '完成整併'
                else:
                    #matchno = 'need to modify'
                    matchno_Chi = chiRes[chiRes==1].index[0]
                    matchno_Eng = engRes[engRes==1].index[0]
                    if matchno_Chi == matchno_Eng:
                        show('condition 1-2: 好像沒有這種情況？？', 5)
                        matchno = matchno_Chi
                        matchResult = '完成整併'
                    else:
                        show('condition 1-3: 中文與英文可能的matchno皆不一致，無法決定match no', 5)
                        matchno = matchno_Chi
                        matchResult = '待確認'
                    score_Eng = engRes[engRes==1].values[0]
                    score_Chi = chiRes[chiRes==1].values[0]

            elif 1 in engRes.values:
                show('condition 2: 英文名稱完全符合', 4)
                matchno = engRes[engRes==1].index[0]
                score_Eng = engRes[engRes==1].values[0]
                score_Chi = FuzzyComp(chiName , reference['name_m_all'][matchno]) 
                matchResult = '完成整併'

            elif 1 in chiRes.values:
                show('condition 3: 中文名稱完全符合', 4)
                matchno = chiRes[chiRes==1].index[0]
                score_Chi = chiRes[chiRes==1].values[0]
                score_Eng = FuzzyComp(engName , reference['name_m_all'][matchno])
                matchResult = '完成整併'

            elif not engRes.empty or not chiRes.empty:
                show('condition 4: 中文名稱及英文名稱部分符合', 4)
                # 給定初始值
                score_Chi = 0
                score_Eng = 0

                # 避免比對結果為空集合
                if not chiRes.empty: score_Chi = chiRes.iloc[0] 
                if not engRes.empty: score_Eng = engRes.iloc[0] 

                # 檢查是否有相同的 match no
                rep_match = [i for i in chiRes.index if i in engRes.index]
                # 避免rep_match重複
                rep_match = list(set(rep_match))
                if rep_match != []:
                    scoreTmp = pd.DataFrame()
                    for imatch in rep_match:
                        scoreTmpChi = chiRes[imatch]
                        scoreTmpEng = engRes[imatch]
                        scoreTmp[imatch] = [scoreTmpChi + scoreTmpEng]
                    matchno = scoreTmp.T[0].idxmax()
                    score_Chi = chiRes[matchno]
                    score_Eng = engRes[matchno]
                    matchResult = '完成整併'

                else:
                    # 中文分數 > 英文分數
                    if score_Chi > score_Eng:
                        matchno = chiRes.index[0]
                        # 重新取得英文名字相似程度的分數
                        if not engRes.empty:
                            score_Eng = FuzzyComp(engName , reference['name_m_all'][matchno]) 
                        matchResult = '完成整併，結果待確認'

                    # 中文分數 < 英文分數
                    elif score_Eng > score_Chi:
                        matchno = engRes.index[0]
                        # 重新取得中文名字相似程度的分數
                        if not chiRes.empty:
                            score_Chi = FuzzyComp(chiName , reference['name_m_all'][matchno]) 
                        matchResult = '完成整併，結果待確認'

                    # 中文分數 = 英文分數
                    elif score_Eng == score_Chi:
                        if score_Eng > 0.6: 
                            ind_Eng = engRes.index[0]
                            ind_Chi = chiRes.index[0]
                            if ind_Eng == ind_Chi:
                                matchno = ind_Eng
                                matchResult = '完成整併'
                            else:
                                matchno = ''
                                matchResult = '無法整併，中英文比對結果不一致'
                        else:
                            matchno = 'need_to_decided'
                    else:
                        matchno = '情況未考慮'
            else:
                show('condition 5: 中文名稱及英文名稱完全不符合', 4)
                matchno = ''
                score_Chi = 0
                score_Eng = 0
                matchResult = '無法整併'
                
            show(f'MatchResult: {matchResult};\t MatchNoRe: {matchno};\t score_Chi: {score_Chi} score_Eng: {score_Eng}',5)

            tmpData['ChiName_MatchList' ] = f'{chiResForWrite.to_dict()}'
            tmpData['EngName_MatchList' ] = f'{engResForWrite.to_dict()}'
            tmpData['MatchResult'       ] = matchResult
            tmpData['MatchNoRe'         ] = matchno
            tmpData['ChiScore'          ] = score_Chi
            tmpData['EngScore'          ] = score_Eng
            tmpData['CASscore'          ] = 0

            data_noCAS = data_noCAS.append(tmpData)
            #tmpData.to_csv(self.FN_bckup, mode='a', header=False)
            tmpData.to_csv(self.FN_noCAS, mode='a', header=False)

        #data = data.drop(data_noCAS.index)
        #self.data_noCAS = data_noCAS
        #self.chiRes = chiRes
        #self.engRes = engRes
        #self.reference = reference
        #self.dataPro = dataPro
        return data
    # finished "def _oneByOne(self, data):"
    # **********************************************************************************************

    def _writeCSVbackup(self):
        head = ['CAS', 'Cname', 'Ename', 'CAS-Rev', 'Cname-Rev', 
                'Ename-Rev', 'ChiName_MatchList', 'EngName_MatchList', 'MatchResult', 'MatchNoRe', 
                'ChiScore', 'EngScore', 'CASscore']
        tmp = pd.DataFrame(columns=head)
        tmp.to_csv(self.FN_noCAS)

    def _calStatistic(self):
        # for statistic
        #statistic
        if not os.path.isfile(self.fileName_statistic):
            statistic = pd.DataFrame({'method': ['count']})
            statistic.set_index('method', inplace=True)
            statistic.loc['count', 'Total'              ] = len(self.dataRaw)
            statistic.loc['count', 'exception'          ] = len(self.dataException)
            statistic.loc['count', 'Total-except'       ] = statistic.loc['count', 'Total'] - statistic.loc['count', 'exception']
            statistic.loc['count', 'With_CAS'           ] = len(self.indWithCAS   )
            statistic.loc['count', 'With_CAS_matched'   ] = len(self.indCASmatch)
            statistic.loc['count', 'With_CAS_unmatched' ] = len(self.indCASunmat)

            statistic.loc['count', 'Without_CAS'        ] = statistic['Total']['count'] - statistic['exception']['count'] - statistic['With_CAS'   ]['count']
            statistic.to_excel(self.fileName_statistic)
        else:
            statistic = pd.read_excel(self.fileName_statistic, index_col=[0])


        # for statistic ratio
        statistic.loc['ratio[%]'] = statistic.loc['count']/statistic['Total']['count']*100
        statistic.loc['ratio_modified[%]'] = statistic.loc['count']/(statistic.loc['count', 'Total-except'])*100
        statistic.loc['ratio_modified[%]'][['Total', 'exception']] = None
        statistic = statistic.applymap(lambda i: '{:.1f}'.format(i))
        self.statistic = statistic
        statistic.to_excel(self.fileName_statistic)

#-----------------------------------------------------------------------------------------
# 整併檔案          結束 
#-----------------------------------------------------------------------------------------



#-----------------------------------------------------------------------------------------
# 後處理          開始 
#-----------------------------------------------------------------------------------------

class postProcessor(Path):
    def __init__(self, inpFile):
        super(postProcessor, self).__init__(inpFile=inpFile)
        self._run()


    def _run(self):
        df0 = pd.read_excel(self.FN_rawDa)
        df1 = pd.read_excel(self.FN_noCAS, index_col=[0])
        df2 = pd.read_excel(self.FN_wiCAS, index_col=[0])
        df = pd.concat([df1, df2], axis=0)
        df0.fillna('', inplace=True)
        df.fillna('', inplace=True)

        df.sort_index(inplace=True)
        df.drop(['CAS-Rev', 'Cname-Rev', 'Ename-Rev'], axis=1, inplace=True)
        df.columns = ['CAS', '中文名', '英文名', '中文名可能配對列表', '英文名可能配對列表', '配對結果', 'MatchNo_Rev', 'ChiScore', 'EngScore', 'CASscore']
        df = pd.concat([df, df0[['IsMatched', 'MatchNo']]], axis=1)
        df['是否一致'] = df['MatchNo_Rev'] == df['MatchNo'] 
        df.to_excel(self.FN_postP)

        self.df = df
        
#-----------------------------------------------------------------------------------------
# 後處理          結束 
#-----------------------------------------------------------------------------------------


if __name__ == '__main__':
    '''
    說明：先跑一次test_run=False，將原始的指引表整理之後輸出，之後才能跑測試run
    '''
    basePath = os.path.dirname(os.getcwd())
    srcFile = 'data/全部_0528.xlsx'

    r = referenceTable(basePath=basePath, test_run = False)    # test_run = integer or False
    m = mergeChem(basePath=basePath, inpFile=srcFile)
