cd C:\MyProjects\ShellRechargePdfToExcel
& C:\MyProjects\ShellRechargePdfToExcel\.venv\Scripts\Activate.ps1
python shell_recharge_extractor.py "C:\MyLocation\shellevcharging"
move "C:\MyLocation\shellevcharging\*.pdf" "C:\MyLocation\shellevcharging\archive"

