# Excel VBA - read from WEB API

Read from Web API using Excel VBA Macro

## Description

Excel macro, written in VBA reads from [ARSO Weather API](http://meteo.arso.gov.si/uploads/probase/www/fproduct/text/sl/fcast_SLOVENIA_latest.xml), queries the returned XML using xpath and copies the data into Sheet1. Error handling is included.

## Requirements

1. Excel with VBA enabled
2. Enabled developer tab in Excel (options -> customize Ribbon, check Developer)

##Features

- Connects to [ARSO Weather API](http://meteo.arso.gov.si/uploads/probase/www/fproduct/text/sl/fcast_SLOVENIA_latest.xml).
- Queries the returned XML using xpath to get the correct information. 
- Creates an Excel shape (arrow) according to the temperature change (arrow up/arrow down/circle).
- Writes the result in Sheet1.
- In case of errors message box is shown.
