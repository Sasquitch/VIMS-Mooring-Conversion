#Place this file in the directory of the excel file


import xlrd
import pandas as pd
from netCDF4 import Dataset
import numpy as np
import time
import os
from os.path import isfile, join

number_of_headers = 2 # the number of rows in the header of the file

today = time.time()
today = ("File created on: ", time.strftime("%Y-%m-%d %H:%M:%S",time.gmtime(today)))
dir_path = os.path.dirname(os.path.realpath(__file__))
onlyfiles = [f for f in os.listdir(dir_path) if isfile(join(dir_path, f))]

for file in onlyfiles:
	if file.endswith('xls'): #Not all excel files end in xls
		ncFileName = file[:-5] + ('.nc')
		nc = Dataset(ncFileName, 'w', format = 'NETCDF4')
		print("Working on file: ", file)
		excel_file_name = str(file)

wb = xlrd.open_workbook(excel_file_name)
directions_sheet = wb.sheets()[0]
velocity_sheet = wb.sheets()[1]

nc.createDimension('time1',directions_sheet.nrows-number_of_headers) 

c = directions_sheet.col_values(6)
c = list(filter(('').__ne__, c)) #remove blank entries

nc.createDimension('time2', len(c)-number_of_headers+1) #+1 assumes there is a blank space in the header, if no blank space between header and data, remove +1

time_var1 = nc.createVariable("Date and Time 1", np.double, ('time1')) #flesh out information
time_var1.standard_name = "time"
time_var1.calendar = "julian"
time_var1.axis = "T"

lat_var1 = nc.createVariable("Latitude 1", np.double, ('time1'))
lat_var1.standard_name = "latitude"
lat_var1.units = "degrees_north"
lat_var1.axis = "Y"
lon_var1 = nc.createVariable("Longitude 1", np.double, ('time1'))
lon_var1.standard_name = "longitude"
lon_var1.units = "degrees_east"
lon_var1.axis = "X"
direction_raw_var1 = nc.createVariable("Direction Raw 1", np.double, ('time1'))
direction_raw_var1.standard_name = "direction_of_sea_water_velocity"
direction_corrected_var1 = nc.createVariable("Direction Corrected 1", np.double, ('time1'))
direction_corrected_var1.standard_name = "direction_of_sea_water_velocity"
velocity_var1 = nc.createVariable("Velocity 1", np.double, ('time1'))
velocity_var1.standard_name = "sea_water_speed"
velocity_var1.units = "m s-1"

time_var2 = nc.createVariable("Date and Time 2", np.double, ('time2'))
time_var2.standard_name = "time"
time_var2.calendar = "julian"
time_var2.axis = "T"
lat_var2 = nc.createVariable("Latitude 2", np.double, ('time2'))
lat_var2.standard_name = "latitude"
lat_var2.units = "degrees_north"
lat_var2.axis = "Y"
lon_var2 = nc.createVariable("Longitude 2", np.double, ('time2'))
lon_var2.standard_name = "longitude"
lon_var2.units = "degrees_east"
lon_var2.axis = "X"
direction_raw_var2 = nc.createVariable("Direction Raw 2", np.double, ('time2'))
direction_raw_var2.standard_name = "direction_of_sea_water_velocity"
direction_corrected_var2 = nc.createVariable("Direction Corrected 2", np.double, ('time2'))
direction_corrected_var2.standard_name = "direction_of_sea_water_velocity"
velocity_var2 = nc.createVariable("Velocity 2", np.double, ('time2'))
velocity_var2.standard_name = "sea_water_speed"
velocity_var2.units = "m s-1"


#GLOBAL VARIABLES
time_var1[:] = list(filter(('').__ne__, directions_sheet.col_values(0)[number_of_headers:])) #Get data and remove headers
lat_var1[:] = list(filter(('').__ne__, directions_sheet.col_values(1)[number_of_headers:]))
lon_var1[:] = list(filter(('').__ne__, directions_sheet.col_values(2)[number_of_headers:]))
direction_raw_var1[:] = list(filter(('').__ne__, directions_sheet.col_values(3)[number_of_headers:]))
direction_corrected_var1[:] = list(filter(('').__ne__, directions_sheet.col_values(4)[number_of_headers:]))
vel_data1 = list(filter(('').__ne__, velocity_sheet.col_values(8)[number_of_headers:]))
vel_data1 = [int(x) / 100 for x in vel_data1]
velocity_var1[:] = vel_data1

time_var2[:] = list(filter(('').__ne__, directions_sheet.col_values(6)[number_of_headers:]))
lat_var2[:] = list(filter(('').__ne__, directions_sheet.col_values(7)[number_of_headers:]))
lon_var2[:] = list(filter(('').__ne__, directions_sheet.col_values(8)[number_of_headers:]))
direction_raw_var2[:] = list(filter(('').__ne__, directions_sheet.col_values(9)[number_of_headers:]))
direction_corrected_var2[:] = list(filter(('').__ne__, directions_sheet.col_values(10)[number_of_headers:]))
vel_data2 = list(filter(('').__ne__, velocity_sheet.col_values(3)[number_of_headers:]))
vel_data2 = [int(x) / 100 for x in vel_data2]
velocity_var2[:] = vel_data2

#global attributes
nc.ncei_template_version = "NCEI_TimeSeries_Orthogonal"
nc.featureType = "TimeSeries"
nc.title = 'Data collected from N.B. Palmer in Ross Sea from 2001-04-01 to 2006-04-01'
nc.standard_name = ("ISO 19115-2 Geographic Information - Metadata - Part 2: Extensions for Imagery and Gridded Data")
nc.Conventions = "CF-1.6"
nc.id = file
nc.notes = "Data Type: FLUORESCENCE (measured); Units: unitless; Observation Type: in situ; Sampling Instrument: wetlabs fluorometer; Sampling and Analyzing Method: 5 minute averages, 2005-6 data; Data Quality Information"
nc.naming_authority = "edu.vims"
nc.history = today
nc.date_created = today
nc.date_modified = today
nc.creator_name = "Dr. Walker smith"
nc.creator_email = "wos@vims.edu"
nc.creator_url = "http://www.vims.edu/"
nc.institution = "Virginia Institute of Marine Science"
nc.funding = "Related Funding Agency: National Science Foundation (NSF)"
#nc.project = fileNameRaw
nc.publisher_name = "US National Centers for Environmental Information"
nc.publisher_email = "ncei.info@noaa.gov"
nc.publisher_url = "https://www.ncei.noaa.gov"
nc.summary = "'Smith, W.O. Jr., V. Asper, S. Tozzi, X. Liu and S.E. Stammerjohn. 2011a. Surface layer variability in the Ross Sea, Antarctica as assessed by in situ fluorescence measurements. Prog. Oceanogr.  88: 28-45 (doi: 10.1016/j.pocean.2010.08.002).Smith, W.O. Jr., A.R. Shields, J. Dreyer, J.A. Peloquin and V. Asper. 2011b. Interannual variability in vertical export in the Ross Sea: magnitude, composition, and environmental correlates. Deep-Sea Res. I 58: 147-159.'nc.contributor_name = 'Suggested Author List: Walker Smith, Vernon Asper'"
nc.sea_name = 'Southern Ross Sea'
nc.close()