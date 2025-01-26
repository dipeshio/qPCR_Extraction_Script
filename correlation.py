import pandas
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from tkinter import filedialog
from tkinter import *
import textwrap

def correlation():
    root = Tk()
    root.withdraw()
    data = pandas.read_excel(filedialog.askopenfilename(), "HuffSM")
    plt.figure(figsize=(9,7))
    m_matrix = data.corr(method="spearman").round(3)

    ## creating a mask
    m_mask = np.triu(np.ones_like(m_matrix,dtype=bool))

    ax = sns.heatmap(m_matrix, cmap="Blues", vmin=0, vmax=1, annot=True, fmt="0.2f", square=True,
                mask= m_mask)
    ax.set_title('Correlation Heatmap, Huff and SM', fontdict={'fontsize': 20}, pad=12)
    plt.yticks(wrap=True)

    ## text wrapping
    labels = [textwrap.fill(label.get_text(), 12) for label in ax.get_xticklabels()]
    ax.set_xticklabels(labels)

    plt.savefig('heatmapHuffSM.png', dpi=600, bbox_inches='tight')
    plt.show()


correlation()