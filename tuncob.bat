set TUNCOB_PATH=C:\temp\tuncob
set TUNCOB_TARGET=stepuser@stephost
set TUNCOB_IDKEY=C:\Users\pcuser\.ssh\id_rsa
set OPENSVC=3389

for /f %a in (%TUNCOB_PATH%\tunport.txt) do (
set STEPPORT=%%a
)

ssh -N -R %STEPPORT%:localhost:%OPENSVC% %TUNCOB_TARGET% -i %%TUNCOB_IDKEY%
