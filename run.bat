set pythonpath=C:\Python27_Excel_PDF\Python27_Excel_PDF
set programPath=C:\conversionLib\program
set path=%pythonpath%;%path%;%programPath%

python -u run.py

python scripts\createHTMLreport.py Results\Report.xlsx Results\Results.html
python scripts\createJUnitReport.py QRHooks QRHooks Results\Report.xlsx Results\Results.xml
pause