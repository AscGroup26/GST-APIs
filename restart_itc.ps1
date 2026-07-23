# Restart Combined R1 + ITC app (port 8502)
Get-NetTCPConnection -LocalPort 8502 -ErrorAction SilentlyContinue |
    Select-Object -ExpandProperty OwningProcess -Unique |
    ForEach-Object { Stop-Process -Id $_ -Force -ErrorAction SilentlyContinue }
Start-Sleep -Seconds 2
Set-Location $PSScriptRoot
python -m streamlit run gstr1_dashboard.py --server.port 8502
