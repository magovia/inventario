@echo off

 
echo ===========================
echo    Sistema de respaldo
echo 	v1.0
echo	git branch Main
echo ===========================

echo ATENCION: Asegurese de tener conexion a internet...
Pause

git checkout main
git pull origin main

echo ==============================================
echo Base de datos ha sido Actualizada exitosamente.
echo ==============================================
pause 

