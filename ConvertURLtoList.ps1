[CmdletBinding()]
param ([string]$dir = "c:\users\rick\")
$links = New-Object System.Collections.ArrayList
Get-ChildItem $dir -filter *.url |
ForEach-Object {
    $content = $_
    #$wshShell2 = New-Object -ComObject "Wscript.Shell"
    $fileName = $content.BaseName
    $extension = $content.Extension
    $fileContent = Get-Content -Path $dir$fileName$extension
    $url = ""
    foreach ($item in $fileContent){
        if ($item.StartsWith("URL=")){
            $url = $item.Substring(4)
            $snippet = @'
            <li><a href="{0}">{1}</a></li>
'@ -f $url, $fileName
            $links.Add($snippet)
        }
 
    }
    
    $html = @"
    <html>
    <body> 
        <ul> Video Links from Class `r`n
"@

    foreach ($link in $links)
    {
        $html += "`r`n $link"
    }
    $html += "`r`n</ul>`r`n</body>`r`n</html>"
    $html > $dir\links.html
}