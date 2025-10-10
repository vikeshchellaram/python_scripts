import pandas as pd
import numpy as np
import os
import pwinput
import configparser
from io import StringIO
import logging
import seaborn as sns
import matplotlib as plt
from fredapi import Fred
from prophet import Prophet
import requests
from bs4 import BeautifulSoup
import cx_Oracle
import warnings



def store_password(msg='Password: '):
    password = pwinput.pwinput(msg)
    return password



def read_credentials_from_ini(ini_path, section):
    if os.path.isfile(ini_path):
        config = configparser.ConfigParser()
        try:
            print('Reading %s credentials from %s' %(section, ini_path))
            config.read(ini_path)
            username = config[section]['username']
            password = config[section]['password']
            print('%s creadentials retrieved succefully!' %section)
        except KeyError:
            username = None
            password = None
            print('Make sure a section [%s] exists in the ini file.' %section)
    else:
        folder, _ = os.path.split(ini_path)
        print('Make sure an ini file is located at %s' %folder)
        username = None
        password = None
        
    return username, password



def read_file(file,name_of_sheet=0):
    try:
        if isinstance(file,str) ==True:
            file_in_memory = file
        else:
            file_in_memory = StringIO (file.getvalue().decode('utf-8')) 
           
        df = pd.read_csv(file_in_memory,sep=',',thousands=',')
            
    except AttributeError or UnicodeDecodeError:
        df = pd.read_excel(file,sheet_name=name_of_sheet)
        
    return df



def log(logfile, log_file_name):
    logging.basicConfig(level=logging.INFO, filename=os.path.join(logfile,log_file_name), filemode='a',
                        format='[%(filename)s] - %(levelname)s [%(asctime)s]: %(message)s',
                        datefmt = '%d/%m/%y %H:%M')
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.DEBUG)
 
    return logger



class text_analysis(object):
    
    def __init__ (self, text):
        # pre-processing
        text = text.replace('.','').replace('!','').replace('?','').replace(',','')
        text = text.lower()
        self.fmtText = text
        
    def freq_all(self):        
        w_list = self.fmtText.split(' ')
        frequency = {}
        for w in set(w_list):
            frequency[w] = w_list.count(w)
        return frequency
    
    def freq_of_word(self,w):
        freq_dict = self.freq_all()
        if w in freq_dict:
            return freq_dict[w]
        else:
            return 0



def boxplot(x, y, data):
    sns.boxplot(x, y, data)
    
    
    
def scatterplot(df, x, y):
    x = df[x]
    y = df[y]
    plt.pyplot.scatter(x,y)

    plt.pyplot.title(f"{x.name} vs {y.name}")
    plt.pyplot.xlabel(x.name)
    plt.pyplot.ylabel(y.name)
    
   
    
def heatmap(df, cmap="RdBu"):
    plt.pyplot.pcolor(df, cmap=cmap)
    plt.pyplot.colorbar()
    plt.pyplot.show()



def binning(x, df):
    bins = np.linspace(min(df[x]), max(df[x]),4)
    group_names = ["Low", "Medium", "High"]
    df[f"{x}-binned"] = pd.cut(df[x],bins,labels=group_names,include_lowest=True)
    plt.pyplot.hist(df[f"{x}-binned"], bins = 3)
    plt.pyplot.xlabel(x)
    plt.pyplot.ylabel("count")
    plt.pyplot.title(f"{x.capitalize()} per bin")



def fred_data(user_input):
    fred = Fred(api_key='')
    series = fred.get_series(user_input)
    series = series.reset_index()
    series.rename(columns={'index': 'Date', 0: 'Price/Rate'}, inplace=True)
    
    series.plot(x="Date", y="Price/Rate")
    
    print (series.tail())
    return series



def forecasting(data):
    data.rename(columns={'Date': 'ds', 'Price/Rate': 'y'}, inplace=True)
    m = Prophet()
    m.fit(data)
    
    future = m.make_future_dataframe(periods=365)
    future.tail()
    
    forecast = m.predict(future)
    forecast[['ds', 'yhat', 'yhat_lower', 'yhat_upper']].tail()
    m.plot(forecast)
    m.plot_components(forecast)
    
    df2 = forecast[['ds', 'yhat', 'yhat_lower', 'yhat_upper']]
    df2 = df2.merge(data, on='ds', how='left')



def oracle_connection(query, file):

    id = pd.read_excel(os.path.join(os.getcwd(), file), header=None)
    connection = cx_Oracle.connect(user=id.iloc[0,0], password=id.iloc[1,0], dsn='')
    
    with connection:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df = pd.read_sql(sql=query, con=connection)
    return df





