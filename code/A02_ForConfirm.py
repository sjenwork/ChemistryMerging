import pandas as pd
import os, re, copy



class confirm():
    def __init__(self, inFN, kwFile):
        self.inFN = inFN
        self.kwFile = kwFile
        self._readKW()
        self._path()
        self._read()
        self._process()
        #self._statistic()
        self._pivot()

    def _path(self):
        self.BasePath = os.path.dirname(self.inFN)
        self.inName   = os.path.join(self.BasePath, 'NameTable.xlsx')
        self.outName  = os.path.join(self.BasePath, 'statistic.xls')
        self.outName2 = os.path.join(self.BasePath, 'pivot_table.xls')

    def _readKW(self):
        with open(self.kwFile, 'r') as f:
            kw = f.readlines()
        kw = [i for i in kw if i!='\n']
        kw = '|'.join(kw)
        kw = kw.replace('\n','')
        self.kw = kw

    def _nameList(self):
        nameList = pd.read_excel(self.inName, header=None, index_col=0)
        nameList.fillna('', inplace=True)
        nameList = nameList.to_dict()
        return nameList

    def _read(self):
        df = pd.read_excel(self.inFN)
        df.columns = ['CName', 'EName', 'CASNo', 'TableName', 'MatchNo', 'Temp', 'TransId', 'ismatched']
        #df.drop(['Temp', 'TransId'], axis=1, inplace=True)
        df = df.applymap(lambda i: re.sub('^-*$', '', i) if type(i) == str else i)
        df.fillna('', inplace=True)

        self.df = df

    def _process(self):
        CName = self.df['CName'].astype(str)
        indKW = CName.str.contains(self.kw)
        self.indKW = indKW
        df2 = self.df[~indKW]
        self.df2 = df2

    def _statistic(self):
        df2 = self.df2
        columns=['Only CAS', 'Only ChiName', 'Only EngName', 'CAS & ChiName', 'CAS & EngName', 
                  'ChiName & EngName', 'All standard', 'sub-Total', 'Removed', 'Total']
        tmp = pd.DataFrame(columns=columns)
        tmp['All standard'     ]    = [len(df2[(df2.CName!='') & (df2.EName!='') & (df2.CASNo!='')])]
        tmp['ChiName & EngName']    = [len(df2[(df2.CName!='') & (df2.EName!='') & (df2.CASNo=='')])]
        tmp['CAS & EngName'    ]    = [len(df2[(df2.CName!='') & (df2.EName=='') & (df2.CASNo!='')])]
        tmp['CAS & ChiName'    ]    = [len(df2[(df2.CName=='') & (df2.EName!='') & (df2.CASNo!='')])]
        tmp['Only CAS'         ]    = [len(df2[(df2.CName=='') & (df2.EName=='') & (df2.CASNo!='')])]
        tmp['Only ChiName'     ]    = [len(df2[(df2.CName!='') & (df2.EName=='') & (df2.CASNo=='')])]
        tmp['Only EngName'     ]    = [len(df2[(df2.CName=='') & (df2.EName!='') & (df2.CASNo=='')])]
        tmp['Removed'          ]    = indKW.sum()
        tmp['Total'            ]    = len(self.df)
        tmp['sub-Total'        ]    = tmp['Total'] - tmp['Removed']
        tmp.index = ['數量']
        tmp.loc['比例']             = tmp.loc['數量']/tmp['sub-Total']['數量']
        tmp.loc['比例',['Removed','Total']] = ''
    
        tmp.to_excel(self.outName)    
        self.tmp = tmp

    def _pivot(self):
        df2 = self.df2
        df3 = pd.pivot_table(df2, index='Temp', columns=['ismatched'], aggfunc='count')['CASNo']
        self.df3_backup = copy.deepcopy(df3)
        index = df3.index.to_frame()
        index.columns = [1]
        nameList = self._nameList()
        index = index.replace(nameList)
        df3.index = index[1]

        self.index = index
        self.df3 = df3
        self.nameList = nameList
        df3.to_excel(self.outName2)    


if __name__ == '__main__':
    inPath = '/Users/jen/GD_simenvi/SimEnvi/Project/109E10_ChemCloud/01.CodingWork/A03_Merge/data/20200528/'
    #inFile = os.path.join(inPath, '全部_0528.xlsx')
    inFile = os.path.join(inPath, '全部_0528-2.xlsx')   # 新增是否整併之欄位
    #inFile = os.path.join(inPath, '全部_0528_test.xlsx')
    kwPath = os.path.join(inPath, 'keyword')

    c = confirm(inFN = inFile, kwFile=kwPath)
