# companyProfile

this little GUI based off of Gooey allows you to take an excel file with a list of company names and stock codes and retrieve vital financial and basic information according to your needs and taste.

please make sure to have the first row have the columns 'Name' and 'Code' as the first two column names. Otherwise it won't work.

Required packages(install using pip): pandas, yfinance, gooey, os, json, time, datetime, pyinstaller 

to create an executable stand-alone application from this code, run

pyinstaller -F --windowed build.spec

this product is not related to any company or organization in any way. 

personal development project of Taekyu Kim. 