# Set Filter Criteria
criteria= {'Direc':None,
         'Div':None,
         'Branches':None,
         'AssetValCond':['Fair','Poor'],
         'AssetVal':[None, None],
         'RepCost':[None,None]}



from AnalyzeData.FilterData import filter_data
files = filter_data(criteria)

from AnalyzeData.WriteOutput import WriteOutput
WriteOutput(files)

temp=1