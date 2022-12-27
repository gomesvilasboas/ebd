#!/usr/bin/env python
# coding: utf-8

import tabula
import pandas as pd
import xlsxwriter
from matplotlib import pyplot as plt
import json
import matplotlib
import numpy as np


def WriteExcel(df, total, unique, countGA, week):
    writer = pd.ExcelWriter(r'/Users/fabriciovilasboas/source/repos/EBD/xlsx/EstatisticaEBD_' + str(week) + '.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name="Total de contribuições", index=False, header=True)
    total.to_excel(writer, sheet_name="Contagem total GA")
    unique.to_excel(writer, sheet_name="Participantes únicos", index=False)
    countGA.to_excel(writer, sheet_name="Contagem únicos GA")
    writer.save()


def EBDStatistics(week):
    fileName = r'/Users/fabriciovilasboas/source/repos/EBD/raw/ebd_' + str(week) + '.pdf'
    dfRaw = tabula.read_pdf(fileName, pages='all')
    dfList = list()
    for page in range(0,len(dfRaw)):
        df = pd.DataFrame()
        if page == 0:
            df = dfRaw[page][['Unnamed: 0', 'Unnamed: 1']]
            df.columns = df.iloc[0]
            df = df[1:]
            df.where(df['Membro'] != 'Membro', inplace=True)
        else:
            df = dfRaw[1][['Membro', 'G.A']]
        dfList.append(df)
    df = pd.concat(dfList)
    df.where(df['G.A'] != 'GA-02- Dalton Douglas da Silva', inplace=True)
    df.dropna(inplace=True)
    df.reset_index(drop=True, inplace=True)

    total = df.groupby('G.A').count()

    unique = df.drop_duplicates().sort_values(by=['G.A', 'Membro'])

    countGA = unique.groupby('G.A').count()

    return df, total, unique, countGA


def PlotCountGA(week,countGA):
    title = "Semana " + str(week)
    countGA.plot(title=title, kind='pie', y='Membro', autopct='%1.1f%%', startangle=90, shadow=False, labels=countGA.index, legend = False, fontsize=12, figsize=(10,10))
    plt.savefig('/Users/fabriciovilasboas/source/repos/EBD/fig/week_' + str(week) + '.png')

def GetListDetails(feature):
    details = list()
    for week in range(0,len(timeline)):
        details.append(timeline[week][feature]['Membro'].count())
    return details


def perGroupPlot():

    features = {'Total': 0, 'Unique': 2}

    labels = range(1,len(timeline)+1)
    total = GetListDetails(features['Total'])
    unique = GetListDetails(features['Unique'])

    x = np.arange(len(labels))  # the label locations
    width = 0.8  # the width of the bars

    fig, ax = plt.subplots()
    rects1 = ax.bar(x - width/2, total, width, label='Total')
    rects2 = ax.bar(x, unique, width, label='Únicos')

    # Add some text for labels, title and custom x-axis tick labels, etc.
    ax.set_ylabel('Contagem')
    ax.set_xlabel('Semana')
    ax.set_title('Série temporal - Contribuições da EBD')
    ax.set_xticks(x)
    ax.tick_params(axis='both', which='major', labelsize=12)
    ax.set_xticklabels(labels)
    ax.legend()
    ax.bar_label(rects1, padding=3)
    ax.bar_label(rects2, padding=3)
    fig.set_size_inches(15,5)
    fig.tight_layout()
    plt.savefig(r'/Users/fabriciovilasboas/source/repos/EBD/fig/TimelineEBD.png')

def calculateEfficiency():
    allParticipants = pd.concat([timeline[i][2] for i in range(0,24)])
    count = allParticipants.groupby('Membro').count()
    efficiency = count/24
    efficiency.sort_values('G.A').to_excel(r'/Users/fabriciovilasboas/source/repos/EBD/xlsx/Efficiency.xlsx')

if __name__ == '__main__':
    timeline = list()
    for week in range(1,25):
        df, total, unique, countGA = EBDStatistics(week)
        #WriteExcel(df, total, unique, countGA, week)
        timeline.append((df, total, unique, countGA, week))
        #PlotCountGA(week, countGA)