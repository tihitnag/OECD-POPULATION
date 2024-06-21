import jpype
jpype.startJVM()
from asposecells.api import Workbook
import pandas as pd
import os
import numpy as np
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import datetime
import geopandas as gpd
import random
import requests
from fpdf import FPDF
import plotly as pl
import plotly.graph_objects as go
import plotly.io as pio
import folium
from folium.plugins import HeatMap
from statsmodels.tsa.arima.model import ARIMA
from statsmodels.tsa.api import VAR
from sklearn.ensemble import RandomForestRegressor
from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_squared_error

if not os.path.exists('data/raw'):
  os.makedirs('data/raw')

if not os.path.exists('data/converted'):
  os.makedirs('data/converted')

if not os.path.exists('resources/generated-charts'):
  os.makedirs('resources/generated-charts')

file_list = os.listdir('data/raw')
for file in file_list[1:]:
    workbook = Workbook(f'data/raw/{file}')
    workbook.save(f'data/converted/{file.split(".")[0]}.xls')

jpype.shutdownJVM()