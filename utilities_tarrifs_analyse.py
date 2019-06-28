
import pandas as pd
import matplotlib.pyplot as plt
import xlrd, openpyxl
import numpy as np
from pandas import ExcelWriter
import glob
import datetime
import seaborn as sns; sns.set()
from pandas.plotting import register_matplotlib_converters
# from sklearn.preprocessing import PolynomialFeatures
# from sklearn.linear_model import LinearRegression
# from sklearn.pipeline import make_pipeline
register_matplotlib_converters()


# def PolynomialRegression(degree=2, **kwargs):
# 	return make_pipeline(PolynomialFeatures(degree),
# 							LinearRegression(**kwargs))


# variables
region_mask = ''
agency_mask = 'Марушкинское'
tariff_items = ['Холодное водоснабжение', 'Горячее водоснабжение', 'Водоотведение', 'Отопление']
hot_water_option = 'для систем ГВС с полотенцесушителями'

print('Tariff Items: {}'.format(tariff_items))
print()

files = glob.glob("datasets/*.xlsx")
frames = []
for f in files:
    df = pd.read_excel(
        f
        # parse_dates=True,
        # date_parser=dateparse,
    )
    frames.append(df)

src = pd.concat(frames, axis=0, sort=False)
src = src[['StartDate', 'Region', 'TariffItem', 'Agency', 'TariffValue', 
		'ConsumptionTarget'
		]]


src.set_index('StartDate', inplace=True)
src.index = pd.to_datetime(src.index)
src.sort_index(ascending=True, inplace=True)

# fetching list of providers
if agency_mask != '':
    src = src[src['Agency'].str.contains(agency_mask)]
    print('Providers List: {}'.format(src['Agency'].unique()))
    print()

writer = ExcelWriter('reports/report.xlsx')

for tariff_item_mask in tariff_items:
    df = src[src['TariffItem'].str.contains(tariff_item_mask)]

    if tariff_item_mask == 'Горячее водоснабжение':
        df = df[df['ConsumptionTarget'].str.contains(hot_water_option)] ### only for hot water


    tariff_change = df.drop_duplicates(subset=['TariffValue'])['TariffValue']
    change_date = df.drop_duplicates(subset=['TariffValue']).index

    df = df.drop_duplicates(subset=['TariffValue'])
    df['Percent'] = (df['TariffValue'].div(df['TariffValue'].shift(periods=1, fill_value=np.nan)) - 1) * 100

    print(df.head())
    print(df.columns)
    print(df.shape)

    
    df[['TariffItem', 'TariffValue', 'Percent']].to_excel(writer, sheet_name=tariff_item_mask)
    

    # plotting part
    fig, axes = plt.subplots(figsize=(8, 8))
    plt.scatter(df.index, df['TariffValue'])
    axes.grid(axis='y', linestyle='--', alpha=0.5, marker=0)
    axes.grid(axis='x', linestyle='--', alpha=0.5, marker=0)

    for i, j in zip(tariff_change.index, tariff_change.values):
        axes.annotate(str(j), xy=(i, j + 0.2))
        plt.xticks(change_date, change_date.strftime('%m.%d.%y'), rotation='vertical')

    plt.xlabel('Date')
    plt.ylabel('Tariff Value')
    plt.title(tariff_item_mask)
    plt.tight_layout()
    plt.savefig('reports/{}.png'.format(tariff_item_mask))

writer.save()
writer.close()
