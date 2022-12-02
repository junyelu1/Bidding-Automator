import pandas as pd
import numpy as np
import os
import seaborn as sns
import matplotlib.pyplot as plt
from components import dataframeToImg


def bidExcelProcessing(filePath: str, costPath: str, parameterPath: str, outPath: str) -> dict:
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
            subproject.to_excel(outPath, engine="openpyxl",
                                sheet_name=key, index=False)


def plotBidPercentages(filePath):

    assert ".xlsx" in filePath, f"Input FilePath given is not Excel"

    bidFiles = pd.read_excel(filePath, sheet_name=None)

    _, axs = plt.subplots(nrows=len(bidFiles), figsize=(6, 4))
    i = 0

    # Plot Percengtages on the graph to show distribution
    for key, bidFile in bidFiles.items():
        # Remove unnecessary data, Calculate sum again because of inability to receive data
        bidFile = bidFile[bidFile['包名称'].str.match('^((?!Total).)*$')]
        plotFile = bidFile.groupby(['包名称']).sum(
            numeric_only=True).reset_index()
        plotFile['比例'] = 1 - plotFile['CB总价']/plotFile['含税总价']

        if len(bidFiles) == 1:
            sns.scatterplot(data=plotFile, x=np.arange(
                plotFile.shape[0]), y='比例', ax=axs)
            axs.set_title(key)
            axs.set_ylim(min(plotFile['比例'])-0.01, max(plotFile['比例'])+0.01)

        else:
            # Plot with Seaborn
            sns.scatterplot(data=plotFile, x=np.arange(
                plotFile.shape[0]), y='比例', ax=axs[i])
            axs[i].set_title(key)
            axs[i].set_ylim(min(plotFile['比例'])-0.01, max(plotFile['比例'])+0.01)

            i += 1


def exportTempBid(filePath):

    assert ".xlsx" in filePath, f"Input FilePath given is not Excel"

    bidFiles = pd.read_excel(filePath, sheet_name=None)

    # Return Number of Baos in the file
    height = max([sum(bidFile['包名称'].str.contains('Total')) -
                 1 for bidFile in bidFiles.values()])

    # Dynamically adjust figure height and width based on content
    _, axs = plt.subplots(ncols=len(bidFiles),
                          figsize=(10, max(height // 2.5, 4.5)))
    i = 0

    for key, bidFile in bidFiles.items():
        bidFile = bidFile[bidFile['包名称'].str.match('^((?!Total).)*$')]
        bidOwner = bidFile.loc[0, '项目单位']
        bidFile = bidFile.groupby(['包名称']).sum(numeric_only=True)
        bidFile = bidFile['含税总价'].apply('{:.6f}'.format)
        bidFile = bidFile.reset_index()

        if len(bidFiles) == 1:
            dataframeToImg(axs, bidFile, bidOwner + '\n' + key)

        else:
            dataframeToImg(axs[i], bidFile, bidOwner + '\n' + key)
            i += 1


def migration(filePath, outPath):
    '''
    Mapping completed table onto worksheets ready for submission
    ---
    input: confirmed Excelsheets
    output: Excelsheets ready for submssion
    ---
    '''

    # Input and process all files, locate relevant sheets by matching number
    try:
        inputFile = pd.read_excel(
            filePath, sheet_name=None, converters={'网省采购申请号': int})
        table = pd.read_excel(
            outPath, skiprows=0, usecols='A:S', converters={'网省采购申请行号': int})
        table = table.drop('未含税单价(万)', axis=1)
    except:
        return "File format might have changed, check source files."

    if len(table.loc[0, '分包名称']) == 2:
        table.loc[:, '分包名称'] = table.loc[:, '分包名称'].apply(
            lambda x: x.replace('包', '包0') if len(x) == 2 else x)
    for _, sheet in inputFile.items():
        if sheet.loc[0, '网省采购申请号'] in list(table['网省采购申请行号']):
            inputSheet = sheet
            break

    # Migrate data
    temp = inputSheet.loc[:, ['包名称', '网省采购申请号', '物资名称', '未含税单价']]
    temp = temp.rename(columns={
        '包名称': '分包名称', '网省采购申请号': '网省采购申请行号', '物资名称': '物料名称', '未含税单价': '未含税单价(万)'})
    try:
        table = pd.merge(table, temp, on=['分包名称', '物料名称', '网省采购申请行号'])
    except:
        return f"{outPath.split('/')[-1].split('.')[0]} 存在数据问题"

    # Followup with computations
    table['税率（%）'] = 13
    table['含税单价(万)'] = table['未含税单价(万)'] * (1 + table['税率（%）']/100)
    table['未含税合价(万)'] = table['未含税单价(万)'] * table['数量']
    table['含税合价(万)'] = table['含税单价(万)'] * table['数量']

    # Check if two totals is the same
    if not np.isclose(sheet.tail(1)['含税总价'].values[0], table.sum()['含税合价(万)'], rtol=0, atol=1e-10):
        return f"{outPath.split('/')[-1].split('.')[0]} 与表格总价差异较大"

    new_cols = ['分包编号', '分包名称', '轮次', '附件上传状态', '物料编码', '物料名称', '技术规范书ID', '网省采购申请行号',
                '项目单位', '单位', '扩展描述', '包限价(万)', '行限价(万)', '数量', '未含税单价(万)', '税率（%）', '含税单价(万)',
                '未含税合价(万)', '含税合价(万)']
    finalTable = table.reindex(columns=new_cols)

    # Safeguard against price ceilings
    if any(table['行限价(万)'].astype(str).str.contains('\d', regex=True)):
        if any(table['行限价(万)'] < table['含税合价(万)']):
            return f"{outPath.split('/')[-1].split('.')[0]} 超过行限价"

    with pd.ExcelWriter(outPath, mode='a', engine='openpyxl', if_sheet_exists='new') as writer:
        finalTable.to_excel(writer, sheet_name='报价方式-单价',
                            index=False, startrow=1, startcol=1)


def splitFileCal(filePath: str, outPath: str, *submitted_file):
    '''
    Divide Excels downloaded from SGCC ECP2.0 and split them into separate sheets based on bao
    Calculate price scores and other metrics for each individual baos
    ---
    download_file: Directory of file downloaded from ECP2.0
    outpath: Directory of the output file
    submitted_file: Directory of the original pricing file
    --
    return: metrics
    '''

    if submitted_file:
        pricing = pd.read_excel(submitted_file[0], sheet_name=None)

    # Clean up Bao Names and Bid amounts
    tallyFile = pd.read_excel(filePath)
    bid_tally = cleanFile(tallyFile)
    baoNames = list(set(bid_tally['分包名称']))
    baoNames.sort()

    # Needs to create a try/except statement for error handling
    averageMethod = int(input('若本次投标为区间平均价法, 输入1, 次低价平均法输入0'))
    if averageMethod:
        cvalue = float(input('请输入该标段C值, 小数形式'))
    n1value = float(input('请输入该标段n1值'))
    n2value = float(input('请输入该标段n2值'))

    directory = []

    for baoName in baoNames:

        # Use self or number 5 as anchor
        bao = bid_tally[bid_tally['分包名称'] == baoName].sort_values(by='投标价格')
        bao = bao.reset_index(drop=True)
        try:
            anchor = bao[bao['投标人名称'] == '浙江高盛输变电设备股份有限公司']['投标价格'].values[0]
        except:
            print("未参加该包投标, 以第五名为基准")
            anchor = bao.loc[5, '投标价格']
        bao['开标备注'] = bao['投标价格']/anchor - 1
        bao.drop(index=bao[np.abs(bao['开标备注']) > 5].index, inplace=True)
        bao.reset_index(drop=True, inplace=True)

        # If pricing file was provided, can calculate metrics
        if submitted_file:
            pricingbao = pricing[bao.loc[0, '分标名称']]
            percent = pricingbao[pricingbao['包名称'] ==
                                 baoName + " Total"]["比例"].values[0]
        else:
            percent = 0

        # Calculate bidding scores given how many bidders
        # This could become an function outside the parameters if required
        bidnum = bao.shape[0]
        if bidnum > 30:
            bideval = bao.iloc[3:-4, :]
        elif bidnum > 20:
            bideval = bao.iloc[2:-3, :]
        elif bidnum > 10:
            bideval = bao.iloc[1:-2, :]
        elif bidnum > 5:
            bideval = bao.iloc[1:-1, :]
        else:
            bideval = bao
        c1 = np.average(bideval['投标价格'])
        bidupper = c1 * 1.1
        bidlower = c1 * 0.85
        # Determine lower and upper bound
        lower = bideval.iloc[0]['投标价格'] >= bidlower
        upper = bideval.iloc[-1]['投标价格'] <= bidupper
        if lower and upper:
            c2 = c1
        else:
            c2 = np.average(bideval[(bideval['投标价格'] > bidlower) & (
                bideval['投标价格'] < bidupper)]['投标价格'])

        # Differentiate between two methods
        if averageMethod:
            average = c2 * (1-cvalue)
        else:
            if lower:
                average = (bideval.iloc[0]['投标价格'] + c2)/2
            else:
                average = (
                    min(bideval[(bideval['投标价格'] > bidlower)]['投标价格']) + c2)/2
            cvalue = 'N/A'

        bao['得分'] = np.where(bao['投标价格'] <= average, 100 - n2value * abs(100 * (bao['投标价格'] / average - 1)),
                             100 - n1value * abs(100 * (bao['投标价格'] / average - 1)))

        directory.append(
            (c1, lower, upper, c2, cvalue, average, percent, anchor))
        # Print all numbers to Excel
        if os.path.isfile(outPath):
            with pd.ExcelWriter(outPath, mode="a", engine='openpyxl') as writer:
                bao.to_excel(writer, index_label='No.',
                             sheet_name='Sheet' + baoName[1:])
        else:
            bao.to_excel(outPath, engine="openpyxl",
                         sheet_name='Sheet' + baoName[1:])

    return directory


def cleanFile(dataframe):
    """Clean up different kinds of ECP2.0 files"""
    # Deal with different versions of the file
    if 'No.' in dataframe.columns:
        index = 'No.'
    else:
        index = '序号'

    if pd.isna(dataframe.iloc[0, 0]):
        dataframe = dataframe.drop(columns="Unnamed: 0")
        dataframe = dataframe.dropna()

    if not '投标价格' in dataframe.columns:
        dataframe = dataframe.rename({'投标价格（万元）': '投标价格'}, axis=1)
    dataframe['分包名称'] = dataframe['分包名称'].apply(
        lambda x: x.replace('包', '包0') if len(x) == 2 else x)

    # Remove non numerical entries
    dataframe = dataframe[dataframe['投标价格'].apply(lambda x: str(x).isascii())]
    dataframe['投标价格'] = dataframe['投标价格'].astype(str)
    dataframe['投标价格'] = dataframe['投标价格'].str.replace(',', '')
    dataframe['投标价格'] = pd.to_numeric(dataframe['投标价格'], downcast='float')
    dataframe = dataframe.drop(labels=index, axis=1)

    return dataframe
