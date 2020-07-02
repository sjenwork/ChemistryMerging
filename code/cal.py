import pandas as pd

total = pd.read_excel('../data/全部_0528.xlsx', index_col=0)
woCAS = pd.read_csv  ('../data/全部_0528_woutCASmatched.csv', index_col=0)
wiCAS = pd.read_excel('../data/全部_0528_withCASmatched.xlsx', index_col=0)
unPro = pd.read_excel('../data/全部_0528_nonProcessFile.xlsx', index_col=0)


count = pd.DataFrame([len(total), len(woCAS), len(wiCAS), len(unPro)], index=['total', 'without CAS', 'with CAS', 'unProcessed'], columns=['count'])

# with CAS and matched
wiCAS2 = wiCAS[['ChiScore', 'EngScore', 'CASscore']]
count_wiCAS = pd.DataFrame([ len(wiCAS2[wiCAS2.sum(axis=1)==3]), 
                             len(wiCAS2[(wiCAS2.sum(axis=1)<3) & (wiCAS2.sum(axis=1)>1)]), 
                             len(wiCAS2[wiCAS2.sum(axis=1)==1])],
                             index = ['CAS+chinese+english', 'between', 'only CAS'], columns=['count'])
count_wiCAS.loc['Total'] = count_wiCAS.sum()


# no CAS
woCAS2 = woCAS[['ChiScore', 'EngScore', 'CASscore']]
count_woCAS = pd.DataFrame([ len(woCAS2[woCAS2.sum(axis=1)==2]), 
                             len(woCAS2[(woCAS2['ChiScore']==1) | (woCAS2['EngScore']==1)]) ,
                             len(woCAS2[woCAS2.sum(axis=1)==0])],
                             index = ['chinese+english', 'chinese or english', 'unmatched'], columns=['count'])
count_woCAS.loc['partial Chi and Eng'] = len(woCAS2) - count_woCAS.sum().values[0]
