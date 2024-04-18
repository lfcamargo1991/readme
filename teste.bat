@echo off
mkdir c:\temp\testeaqui
echo Seu texto aqui >> c:\temp\testeaqui\arquivo.txt
wmic product get name | findstr -i office > C:\temp\office.txt

