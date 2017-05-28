# SuperTueDataCleaner
## Background
The Moreland Council conducted bicycle surveys between 7am and 9am on a Tuesday in early March between 2006 and 2017. 
The methods used for these surveys were similar to the 'Super Tuesday' survey conducted by bicycle network https://www.bicyclenetwork.com.au/general/for-government-and-business/3060/

Data from these serveys and analysis of the results was recorded the excel file 'Traffic Count - Bicycle Count - Bike count - Morning Peak 7am to 9am - Weekday - Super ~ 2017.xlsx'

The script super_tue_cleaner.py performs a scrape of the excel document to create 'clean' tabular data.

The jupyter notebooks 'Super Tuesday Single Site Data Analysis Tool.ipynb' and 'Super Tuesday Multiple Site Data Analysis Tool.ipynb' aggregate the 'clean' observation .csv files located in ./script_output/count_observations/ develop trend information about each count site. Optionally these scripts an analyse a subset of the count (eg only femal bike riders) by modifying configuration details.

Location information for each site including GIS co-ordinates are stored in ./script_output/count_locations/
The Coordinate Reference System used in the orginal excel spreadsheet (and count_location_details.csv is GDA 94 MGA Zone 55 http://spatialreference.org/ref/epsg/gda94-mga-zone-55/

## Config
super_tue_cleaner.py script written in python 2.7.13 and uses xlrd (available from https://pypi.python.org/pypi/xlrd)
.ipynb files uses pandas


# TODO
* Set up a *.github.io page to display and discuss results
* Convert location point information display count site points on a WD84 (Web Macador - used by Google Maps and Open Street Maps) base map 
* Configure the *.github.io page to support leaflet  - display points on a map window
* Create pop up boxes that depict 'change over time' graphs




