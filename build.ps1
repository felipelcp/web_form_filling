$exclude = @("venv", "WebFormFilling.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "WebFormFilling.zip" -Force