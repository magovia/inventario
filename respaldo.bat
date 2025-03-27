@echo off

 
echo ===========================
echo    Sistema de respaldo
echo 	v1.0
echo	git branch Main
echo ===========================

echo ATENCION: Asegurese de tener conexion a internet...
Pause
echo
git status
echo
git add .
echo ==============================================
git commit -m "Base de datos respaldada exitosamente"
echo ==============================================
git push
echo
echo ==============================================
echo Base de datos ha sido respaldada exitosamente.
echo ==============================================
pause 

