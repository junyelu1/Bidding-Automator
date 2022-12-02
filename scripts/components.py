from pandas.plotting import table
import numpy as np


def dataframeToImg(ax, dataframe, title):
    tab = table(ax, dataframe, loc='center')
    ax.set_frame_on(False)
    ax.xaxis.set_visible(False)
    ax.yaxis.set_visible(False)
    tab.set_fontsize(12)
    tab.auto_set_column_width(list(np.arange(dataframe.shape[1])))
    tab.scale(1, 1.5)
    ax.set_title(title)
