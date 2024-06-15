#!/usr/bin/env python
# coding: utf-8

import geopy, os, time, pandas as pd, pickle, folium
from geopy.geocoders import Nominatim
from folium.plugins import HeatMap
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tqdm import tqdm

# In[3]:


Tk().withdraw()


#Asks user to select an Excel file that contains their sales data


file_path = askopenfilename(filetypes=[("Excel File","*.xlsx *.xls")], title = "Select Excel File")


#Prints to terminal if file selection was successful.


if not file_path:
    print("No file selected.")
    raise FileNotFoundError("No file was selected.")
else:
    print(f"Files selected: {file_path}")


#Prints to terminal if their data loaded successfully


try:
    data = pd.read_excel(file_path)
    print("Data loaded successfully")
except Exception as e:
    print(f"Data failed to load: {e}.")
    raise


# In[7]:


data.columns = ["MONTH","CITY_STATE", "AMOUNT"]


# In[8]:


try:
    data_grouped = data.groupby("CITY_STATE")["AMOUNT"].sum().reset_index()
    print("Data aggregated successfully.")
except Exception as e:
    print(f"Failed to group data: {e}.")
    raise

period = str(input("What time period does the data represent? "))

# In[9]:


geolocator = Nominatim(user_agent="RMMisTheBest661")


# In[10]:


geocode_cache = {}


# In[11]:


cache_file = 'geocode_cache.pkl'


# In[12]:


if os.path.exists(cache_file):
    with open(cache_file, "rb") as f :
        geocode_cache = pickle.load(f)


# In[13]:


def get_coordinates(location, pbar=None):
    if location in geocode_cache:
        return geocode_cache[location]
    try:
        loc = geolocator.geocode(location)
        if loc:
            geocode_cache[location] = (loc.latitude, loc.longitude)
            if pbar:
                pbar.update(1)
            return loc.latitude, loc.longitude
        else: 
            return None
    except Exception as e:
        print(f"Geocode failed for {location}: {e}.")
        if pbar:
            pbar.update(1)
        return None

#Adds a progress bar

with tqdm(total=len(data_grouped), desc = "Geocoding") as pbar:
    data_grouped["Coordinates"] = data_grouped["CITY_STATE"].apply(lambda x: get_coordinates(x,pbar))

# In[14]:


data_grouped["Coordinates"] = data_grouped["CITY_STATE"].apply(get_coordinates)


# In[15]:


data_grouped = data_grouped.dropna()


# In[16]:


data_grouped[["Latitude", "Longitude"]] = pd.DataFrame(data_grouped['Coordinates'].tolist(), index=data_grouped.index)


# In[17]:


data_grouped.drop(columns = ['Coordinates'], inplace = True)


# In[18]:


heatmap = folium.Map(location = [data_grouped["Latitude"].mean(), data_grouped["Longitude"].mean()],zoom_start = 6)


# In[19]:


heat_data = [[row["Latitude"], row["Longitude"], row["AMOUNT"]] for index, row in data_grouped.iterrows()]


# In[20]:


HeatMap(heat_data).add_to(heatmap)


# Initialize region nodes that display sale amount, location, and desc.:


for index, row in data_grouped.iterrows():
    amount_formatted = f"${row['AMOUNT']:,.0f}"
    folium.CircleMarker(
        location=[row["Latitude"], row["Longitude"]],
        radius=7,
        popup=f'Location: {row["CITY_STATE"]} <br> Amount: {amount_formatted} ({period})',
        color='blue',
        fill=True,
        fill_color='blue'
    ).add_to(heatmap)


# Saves the heat map in .html, usually in the /user/ folder:

new_file_name = str(input("Input the heat map's file name: ")) + ".html"

heatmap.save(new_file_name) 

print("Interactive heat map file saved onto your computer.")

