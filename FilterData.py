def filter_data(criteria):
# FilterData.py allows the user to filter data based on defined criteria.
# Criteria Considered:
#
#   1 - List of Divisions (e.g., LM)
#   2 - Lsit of Branches (e.g., LMS)
#   3 - Value of Total Assets (between x and y)
#   4 - Total Replacement Cost (between x and y)

    # Import Modules
    import glob
    import os

    # Set Paths
    home = os.getcwd()
    data_path = r'C:\Users\bhearley\Box\Lab Infrastructure Data Collection\Final'

    os.chdir(data_path)
    files_all = glob.glob('*.txt')

    FilesOut = {}

    for q in range(len(files_all)):
        # Read the Text File
        with open(os.path.join(data_path,files_all[q])) as f:
            lines = f.readlines()

        # Get the Branch Name
        key = 'Branch:'
        for i in range(len(lines)):
            if key in lines[i]:
                val  = lines[i][len(key)+1:len(lines[i])-1]

        if len(val) == 3:
            Direc = val[0]
            Div = val[0:2]
            Branch = val[0:3]

        # Set Flag
        crit_flag = 1

        # Evaluate Division Criteria
        if criteria['Div'] != None:
            if Div not in criteria['Div']:
                crit_flag = 0

        # Evaluate Division Criteria
        if criteria['Branches'] != None:
            if Branch not in criteria['Branches']:
                crit_flag = 0

        # Get Total Asset Cost and Filter

        key = 'Number of Assets:'
        for i in range(len(lines)):
            if key in lines[i]:
                val  = lines[i][len(key)+1:len(lines[i])-1]
                line_num = i
        num_assets  = int(val)
        
        data = ''
        for k in range(line_num+2,line_num+2+num_assets):
            data = data + lines[k]
        data= data.split('\n')
        data_all = []
        for k in range(num_assets):
            data_line = data[k]
            data_line = data_line.split('\t')
            data_all.append(data_line)
        tot_asset_cost = 0
        for k in range(num_assets):
            if data_all[k][5] in criteria['AssetValCond']:
                tot_asset_cost = tot_asset_cost + float(data_all[k][6])

        if criteria['AssetVal'][0] != None or criteria['AssetVal'][0] != None:
            if criteria['AssetVal'][0] != None and tot_asset_cost < criteria['AssetVal'][0]:
                crit_flag = 0
            if criteria['AssetVal'][1] != None and tot_asset_cost > criteria['AssetVal'][1]:
                crit_flag = 0
        
        # -- Estimated Cost to Replace Entire Laboratory/Capability ($):
        key = 'Estimated Cost to Replace Entire Laboratory/Capability ($):'
        for i in range(len(lines)):
            if key in lines[i]:
                val  = lines[i][len(key)+1:len(lines[i])-1]
        if criteria['RepCost'][0] != None or criteria['RepCost'][0] != None:
            if criteria['RepCost'][0] != None and tot_asset_cost < criteria['RepCost'][0]:
                crit_flag = 0
            if criteria['RepCost'][1] != None and tot_asset_cost > criteria['RepCost'][1]:
                crit_flag = 0


        if crit_flag == 1:
            div_keys = list(FilesOut.keys())
            if Div not in div_keys:
                FilesOut[Div] = {}

            branch_keys = list(FilesOut[Div].keys())
            if Branch not in branch_keys:
                FilesOut[Div][Branch] = []

            FilesOut[Div][Branch].append(files_all[q])

    return FilesOut




