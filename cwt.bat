del /Q /S "C:\Windows\Temp\*" >> C:\CWT.log 2>&1
(for /D /R "C:\Windows\Temp\" %%d in (*) do rmdir /Q /S "%%d" ) >> CWT.log 2>&1
