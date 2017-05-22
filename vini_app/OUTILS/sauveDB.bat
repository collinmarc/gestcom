
SET J=%date:~-10,2%
SET A=%date:~-4%
SET M=%date:~-7,2%
SET H=%time:~0,2%
SET MN=%time:~3,2%
SET S=%time:~-5,2%

IF "%time:~0,1%"==" " SET H=0%time:~1,1%

@IF EXIST E:\ (GOTO vinicom)
SET REPERTOIRE=C:\SAUVDB
@GOTO end

:vinicom
SET REPERTOIRE=E:\SAUVDB

:end

IF NOT exist "%REPERTOIRE%" md "%REPERTOIRE%"

SET FICHIER=%REPERTOIRE%\vnc5_%1%.bak


rem cd C:\Program Files\Microsoft SQL Server\90\Tools\Binn

sqlcmd -S localhost\SQLEXPRESS -Q "BACKUP DATABASE vnc5 TO DISK = N'%FICHIER%' WITH INIT, NAME = N'Sauvegarde automatique du %J%_%M%_%A%_A_%H%_%MN%_%S%', STATS = 1" >> %REPERTOIRE%/Sauvdb.log