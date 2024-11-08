@echo off

 
echo ===========================
echo    Sistema de respaldo
echo 	v1.0
echo	git branch XxxxXX
echo ===========================

echo ATENCION: Asegurese de tener conexion a internet...
Pause

cd /d "%~dp0"           :: Changes directory to the location of the batch file

git add mydata.csv      :: Stage the specific file you want to push
git commit -m "Your commit message here"  :: Commit with a message
git push origin main     :: Push to the main branch

echo mydata.csv has been pushed successfully.
pause                    :: Pause to keep the window open and see results

