:: do-invoice.bat - 7-Jul-2020 - www.michaelmarmor.com
:: Batch process a folder of OPML outlines for invoice processing
:: This lives in C:\Users\marmo\Invoice-tool-bin\do-invoice.bat
pushd C:\Users\marmo\Invoice-tool-bin\msm-list-to-invoice
for %%F in (C:\Users\marmo\Invoice-tool-bin\OPML-drop-folder\*.opml) do (
   msm-list-to-invoice.exe "%%~dpnxF"
)
popd
