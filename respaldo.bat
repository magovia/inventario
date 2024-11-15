@echo off

 
echo ===========================
echo    Sistema de respaldo
echo 	v1.0
echo	git branch distleimi
echo ===========================

echo ATENCION: Asegurese de tener conexion a internet...
Pause

git add inv_backEnd.accdb 
git commit -m "Base de datos respaldada exitosamente"
git push

echo ==============================================
echo Base de datos ha sido respaldada exitosamente.
echo ==============================================
pause 

