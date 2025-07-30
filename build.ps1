$exclude = @("venv", "botPython.zip", "#material", "json", "downloads", "download_relatorios", ".logs")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "botPython.zip" -Force