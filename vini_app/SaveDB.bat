SET J=%date:~-10,2%
SET A=%date:~-4%
SET M=%date:~-7,2%
SET H=%time:~0,2%
SET MN=%time:~3,2%
SET S=%time:~-5,2%

SET REPERTOIRE=%3
SET FICHIER=%REPERTOIRE%\vnc4_%1%2.bak
REM SET FICHIER=%REPERTOIRE%\vnc4_%J%_%M%_%A%_A_%H%_%MN%_%S%.bak
IF NOT exist "%REPERTOIRE%" md "%REPERTOIRE%" 

cd C:\Program Files\Microsoft SQL Server\100\Tools\Binn

sqlcmd -S ANTEC\SQLEXPRESS -Q "BACKUP DATABASE vnc4 TO DISK = N'%FICHIER%' WITH INIT, NAME = N'Sauvegarde automatique de la base de données', STATS = 1"