import pandas as pd
import numpy as np
import os


def excelProcessing(filePath: str, costPath: str, parameterPath: str, outPath: str):
    """Takes Excel sheets from Sourcefile, and Processing it and output the completed file without styles
        - all path has to be full path
        - filePath: Path of input files
        - costPath: Path of Cost file
        - parameterPath: Path of Parameters file
        - outPath: Path of output file
    """
    assert ".xlsx" in filePath, f"Input FilePath given is not Excel"
    assert ".xlsx" in costPath, f"Costs FilePath given is not Excel"
    assert ".xlsx" in parameterPath, f"Parameters FilePath given is not Excel"
    assert ".xlsx" in outPath, f"Output FilePath given is not Excel"

    # Import all constants from existing files
    costFiles = pd.read_excel(costPath, sheet_name=['Sheet1', 'Sheet4'])
    parameters = pd.read_excel(parameterPath, sheet_name=None)
    bao = parameters['Sheet1']
    guige = parameters['Sheet2']
    shipping = parameters['Sheet3']

    # Main file manipulation
    project = pd.read_excel(filePath, sheet_name=None)

    for key, subproject in project.items():
        # Check num of clients
        clients = len(set(subproject['需求单位']))

        # Identify relevant Cost File
        if '低压电力电缆' in key:
            cost = costFiles['Sheet4']
        else:
            cost = costFiles['Sheet1']

        rowcount = subproject.shape[0]
        subproject = subproject.assign(
            **{'含税单价': np.zeros(rowcount), 'CB': np.zeros(rowcount), '比例': np.zeros(rowcount)})

        for index, row in subproject.iterrows():

            # Processing step 1, rename Baos and Search relevant costs
            if len(row['包名称']) == 2:
                subproject.loc[index, '包名称'] = row['包名称'].replace("包", "包0")
            subproject.loc[index, 'CB'] = cost.loc[cost['产品名称']
                                                   == row['物资名称'], '每千米万元'].values[0]

            # Processing Step 2, Calculate Percentage based on multiple information, type, delivery method, number of clients
            bili = 0
            try:
                bili += guige.loc[(guige['项目编号'] == row['分标编号']) & (guige['项目单位'] ==
                                                                    row['项目单位']) & (guige['物资名称'] == row['物资名称']), '合计调整'].values[0]
            except:
                return "Missing 物资 in Cost File."

            if '地面' in str(row['交货地点']) or '地面' in str(row['交货方式']):
                bili += 0.01

            try:
                if clients > 1:
                    bili += shipping[shipping['需求单位'] ==
                                     row['需求单位']]['运费调整'].values[0]
            except:
                return "Missing 需求单位 in 运费调整"

            bili += bao[(bao['项目编号'] == row['分标编号']) & (bao['项目单位'] == row['项目单位'])
                        & (bao['分标编号'] == key) & (bao['包号'] == subproject.loc[index, '包名称'])]['积'].values[0]

            subproject.loc[index,
                           '含税单价'] = subproject.loc[index, 'CB'] / (1 - bili)

        subproject['CB总价'] = subproject['CB'] * subproject['数量']
        subproject['未含税单价'] = np.round(subproject['含税单价'] / 1.13, 6)
        subproject['含税总价'] = subproject['未含税单价'] * 1.13 * subproject['数量']
        subproject['比例'] = 1 - subproject['CB']/subproject['含税单价']

        new_cols = ['分标编号', '包名称', '分包编号', '项目单位', '需求单位', '项目名称', '工程电压等级', '物资名称', '物资描述',
                    '单位', '数量', '未含税单价', '含税单价', '含税总价', 'CB', 'CB总价', '比例', '首批交货日期',
                    '最后一批交货日期', '交货地点', '备注', '技术规范编码', '网省采购申请号', '总部采购申请号 ', '物料编码', '扩展描述', '扩展编码']
        subproject = subproject.reindex(columns=new_cols)
        subproject.sort_values(by=['包名称', '需求单位', '物资名称', '数量'], inplace=True)

        if os.path.isfile(outPath):
            with pd.ExcelWriter(outPath, mode="a", engine='openpyxl', if_sheet_exists="new") as writer:
                subproject.to_excel(writer, sheet_name=key, index=False)
        else:
            subproject.to_excel(outPath, sheet_name=key, index=False)
