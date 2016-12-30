set /p invoiceFile=invoiceFile:
set /p xslFile=xslFile:
echo %invoiceFile% %xslFile%
.\april.exe CAT5 %invoiceFile% %xslFile%
set /p aaa=finish!