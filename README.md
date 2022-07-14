# s_america_national_dam_inventory_project
Data scraping South American national dam datasets for SWOT project.

Objectives:
- pull out data from reports into excel format
- scrape relevant data from raw excel into preferred format
    for Peru:
        - dam_UID/dam_source/reg_string/northing/easting
            - reg string including:
            'code, east, north, dam name, admin authority, local, river, district, province'

    for Bolivia
        - dam_UID/dam_source/reg_string/lat/long
            - reg string including:
            'name, id, type, area, height, coronoa length, capacity, corona dimension, use, municipality, watershed, river'
...
- for Peru, convert Northing/Easting to Lat/Long
    

Useful links:

article containing Bolivia and Peru national registries:
https://essd.copernicus.org/articles/13/213/2021/essd-13-213-2021.html

openypyxl documentation:
https://openpyxl.readthedocs.io

xlsxwriter documentation:
https://xlsxwriter.readthedocs.io/
