@pushd %~dp0


md "C:\Users\%username%\Desktop\Temp_AMP_PYTHON_CLI"

copy "\\tx1cifs\tx1data\Austin Share\CLI ScaleUp Grant\Year 2\CLI Scheduled Jobs (Do Not Rename)\Calculate Shipping Discrepancies.py" "C:\Users\%username%\Desktop\Temp_AMP_PYTHON_CLI"
copy "\\tx1cifs\tx1data\Austin Share\CLI ScaleUp Grant\Year 2\CLI Scheduled Jobs (Do Not Rename)\SRC Testing Tracking - Final.py" "C:\Users\%username%\Desktop\Temp_AMP_PYTHON_CLI"
copy "\\tx1cifs\tx1data\Austin Share\CLI ScaleUp Grant\Year 2\CLI Scheduled Jobs (Do Not Rename)\Teacher Survey Summary.py" "C:\Users\%username%\Desktop\Temp_AMP_PYTHON_CLI"


set pypath=NULL
if "%username%"=="jmeakin" set pypath=C:\Users\jmeakin\AppData\Local\Programs\Python\Python35-32
if "%pypath%"=="NULL" set /p pypath=Where is Python Installed (e.g. mine is  "C:\Users\jmeakin\AppData\Local\Programs\Python\Python35-32")? Copy Here Then Hit Enter:

"%pypath%\python.exe"  "C:\Users\%username%\Desktop\Temp_AMP_PYTHON_CLI\Calculate Shipping Discrepancies.py"
"%pypath%\python.exe"  "C:\Users\%username%\Desktop\Temp_AMP_PYTHON_CLI\SRC Testing Tracking - Final.py"
"%pypath%\python.exe"  "C:\Users\%username%\Desktop\Temp_AMP_PYTHON_CLI\Teacher Survey Summary.py"

del "C:\Users\%username%\Desktop\Temp_AMP_PYTHON_CLI\Calculate Shipping Discrepancies.py"
del "C:\Users\%username%\Desktop\Temp_AMP_PYTHON_CLI\SRC Testing Tracking - Final.py"
del "C:\Users\%username%\Desktop\Temp_AMP_PYTHON_CLI\Teacher Survey Summary.py"

rd "C:\Users\%username%\Desktop\Temp_AMP_PYTHON_CLI"

@popd