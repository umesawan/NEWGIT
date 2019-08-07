@echo off
pushd %~dp0
powershell.exe -ExecutionPolicy Bypass -windowstyle hidden -Command "& { & '.\Backup and restore.ps1'}" 
