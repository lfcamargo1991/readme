@echo off
mkdir c:\temp\testeaqui
echo Seu texto aqui >> c:\temp\testeaqui\arquivo.txt
wmic product where "name like 'Microsoft Office%'" get name, version > C:\temp\office.txt

