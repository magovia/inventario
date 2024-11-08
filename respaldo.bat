@echo off

 
echo ===========================
echo    Sistema de respaldo
echo 	v1.0
echo	git branch distleimi
echo ===========================

echo ATENCION: Asegurese de tener conexion a internet...
Pause

git add inv_backEnd.accdb      :: Stage the specific file you want to push
git commit -m "Base de datos respaldada exitosamente"  :: Commit with a message
git push distleimi     :: Push to the main branch

echo Base de datos ha sido respaldada exitosamente.
pause                    :: Pause to keep the window open and see results

