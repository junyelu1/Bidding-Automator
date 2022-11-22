import pandas as pd
import numpy as np


def filterDownloadExcel(downloadPath: str, savePath='~/Desktop/相关清单/'):
    '''
    Filter Excel files containing multiple products downloaded from ECP 2.0
    '''

    # Input Type check
    assert ".xlsx" in downloadPath, f"Downloaded FilePath given is not Excel"

    try:
        # Import Files into Pandas DataFrame
        downloadFile = pd.read_excel(downloadPath, sheet_name=None)

        # Future Expansions possible into other products
        relatedProd = ['架空绝缘导线', '集束绝缘导线', '架空线']
        prodList = [*downloadFile]

        fileName = ""
        for key in prodList:
            if not any(prod in key for prod in relatedProd):
                file = downloadFile.pop(key)
                if not fileName:
                    # Extract fileName from Owner
                    fileName = list(file['项目单位'])[0].split(
                        '国网')[1].split('电力')[0]

        if not downloadFile:
            return 'No related subprojects present in this file.'

    except:
        return "Erros occured during extraction Process."

    # Writing Found Sheets into new sheets
    try:
        with pd.ExcelWriter(savePath + fileName + '相关清单.xlsx') as writer:
            for key in downloadFile:
                downloadFile[key].to_excel(writer, sheet_name=key, index=False)
        return f"Subprojects successully filtered. {len(downloadFile)} related subprojects found."

    except:
        return "Errors occured during writing process."
