# Lake Sediment Analysis

## Project Description

My project is an interactive application for analyzing data regarding the content of chemical elements in the sediments of various lakes in Poland. The application uses data contained in an Excel file, which includes information about the names of the lakes, pH values, and the content of various chemical elements.

![image](https://github.com/roberts911/LakeSedimentAnalysis/assets/85109223/6d37d5fd-4d98-4901-8061-78c4e40bb1f2)


## Functionalities

The application enables:

1. **Lake selection** - users can choose one of the available lakes from a dropdown list.
2. **Element selection** - users can choose which elements they want to include in their analysis.
3. **Displaying a pie chart** - after selecting a lake and elements, users can generate a pie chart that shows the content of each of the chosen elements in the sediments of the selected lake.
4. **Displaying a map** - after selecting a lake, users can generate a map that shows the location of the chosen lake.

## Requirements

To use this application, the following libraries must be installed:

- `tkinter`
- `openpyxl`
- `matplotlib`
- `folium`
- `geopy`
- `fake_useragent`

Command to install libraries:

    pip install tkinter openpyxl matplotlib folium geopy fake_useragent

## Instructions

1. Run the main.py file.
2. Select a lake from the dropdown list.

![image](https://github.com/roberts911/LakeSedimentAnalysis/assets/85109223/a2f20e19-244f-47a9-8d52-4cf5185b9cec)

3. Select the elements you want to include in your analysis.

![image](https://github.com/roberts911/LakeSedimentAnalysis/assets/85109223/b4f85e28-dd51-44fe-b59f-a932ab3d9157)

4. Click the "Display Chart" button to generate a pie chart.

![image](https://github.com/roberts911/LakeSedimentAnalysis/assets/85109223/53575020-7b93-47d1-b09d-8502d161bfcc)

5. Click the "Display Map" button to generate a map with the location of the chosen lake.

![image](https://github.com/roberts911/LakeSedimentAnalysis/assets/85109223/2fc0d394-0e33-4e3a-8da4-0d711e46b70b)

## Authors

This project was created by Robert Siurek.
