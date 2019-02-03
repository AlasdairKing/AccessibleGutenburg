@echo off
path=%path%;C:\Program Files\Microsoft Visual Studio\VB98\;C:\Program Files\7-Zip;C:\Program Files\Microsoft SDKs\Windows\v6.0\Bin\;C:\Program Files\SoftwarePassport;C:\Program Files\WinZip Self-Extractor\;C:\Program Files\Crimson Editor

echo ********************************
echo Doing code store 
echo ********************************
7z a -tzip gutenberg.zip @filetypestozip.txt

echo ********************************
echo Gutenberg
echo ********************************
echo Compiling to D:\Installers\Alasdair\files\Program Files\WebbIE
vb6 /m "AccessibleGutenberg.vbp"  /outdir "D:\Installers\Alasdair\files\Program Files\WebbIE"
echo.

echo Copying to this folder and Powerwraps
copy "D:\Installers\Alasdair\files\Program Files\WebbIE\AccessibleGutenberg.exe" "C:\Documents and Settings\Alasdair King\My Documents\accessible\AccessibleGutenberg"
copy "D:\Installers\Alasdair\files\Program Files\WebbIE\AccessibleGutenberg.exe" "D:\Installers\Powerwraps\WebbIE"

pause