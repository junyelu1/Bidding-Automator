import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pandas.plotting import table


def bidResultReporting(outPath: str, bidder: str):
    '''
    Gather information from the bidding report into strings
    While printing, gather information about statistics of performance
    ---
    outpath: Directory of the output file
    '''

    assert ".xlsx" in outPath, f"Input FilePath given is not Excel"

    file = pd.read_excel(outPath, sheet_name=None)

    try:
        projectName = file[list(file.keys())[0]].loc[0, '项目单位']
    except:
        projectName = file[list(file.keys())[0]].loc[0, '采购项目名称']
    subprojectName = file[list(file.keys())[0]].loc[0, '分标名称']
    report = {}

    # Provided additional protection for bidders not participating in the subproject
    for _, sheet in file.items():
        try:
            report[sheet.loc[0, '分包名称']] = round(sheet[sheet['投标人名称'] == bidder]['投标价格'].values[0], 2), round(sheet[sheet['投标人名称'] == bidder]['得分'].values[0], 2), sheet[sheet['投标人名称']
                                                                                                                                                                         == bidder].index[0]+1, sheet['投标人名称'].count(), sorted(sheet['得分'], reverse=True).index(sheet[sheet['投标人名称'] == bidder]['得分'].values[0])+1
        except:
            continue

    if not report:
        return f"{bidder} 没有参加 {projectName} 项目投标."

    _, ax = plt.subplots(figsize=(10, max(len(report) // 2.5, 4)))
    result = pd.DataFrame.from_dict(report, orient='index', columns=[
                                    '金额', '得分', '排名', '厂家数', '名次'])
    tab = table(ax, result, loc='center')
    ax.set_frame_on(False)
    ax.xaxis.set_visible(False)
    ax.yaxis.set_visible(False)
    tab.set_fontsize(12)
    tab.auto_set_column_width(0)
    tab.scale(1, 1.5)
    ax.set_title("采购单位: " + projectName + '\n' + "分标名称: " + subprojectName + '\n'+"投标人名称: " + bidder + '\n' + "平均得分: " +
                 str(np.average([x[1] for x in list(report.values())])) + '\n' + "平均名次: " + str(np.average([x[-1] for x in list(report.values())])))
