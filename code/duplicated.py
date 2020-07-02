import pandas as pd
from work import path
import datetime
import re

class duplicate(path):
    def __init__(self):
        super(duplicate, self).__init__()
        self._run()

    def _run(self):
        rawData = pd.read_excel(self.fileName_reference_raw, index_col=[0])
        resData = pd.read_excel(self.fileName_reference_res, index_col=[0])

        rawData['CASNoMatch(對應的Cas No.)'] = rawData['CASNoMatch(對應的Cas No.)'].apply(lambda i: i.strftime('%Y-%m-%-d') if type(i) == datetime.datetime else i)
        rawData['CASNoMatch(對應的Cas No.)'] = rawData['CASNoMatch(對應的Cas No.)'].apply(lambda i: re.sub('\s+','',i))

        ind = resData['CAS_m'].duplicated()
        indCAS = resData[ind]['CAS_m'].values
        

        dup = {i:rawData[rawData['CASNoMatch(對應的Cas No.)'].str.contains(i)] for i in indCAS}
        dup2 = []
        for i in dup.items():
            ind, df = i
            df.index = pd.MultiIndex.from_product([[ind], df.index])
            dup2.append(df)
        dup = pd.concat(dup2, axis =0)
        dup.to_excel('../data/duplicateCAS.xls')
        self.dup = dup

if __name__ == '__main__':
    d = duplicate()
