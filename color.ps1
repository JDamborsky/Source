function Get-ConsoleBufferAsHtml
{
  $html = [Text.StringBuilder]''
  $null = $html.Append("<pre style='MARGIN: 0in 10pt 0in;
      line-height:normal';
      font-family:Consolas;
  font-size:10pt; >")
  $bufferWidth = $host.UI.RawUI.BufferSize.Width
  $bufferHeight = $host.UI.RawUI.CursorPosition.Y

  $rec = [Management.Automation.Host.Rectangle]::new(
    0,0,($bufferWidth - 1),$bufferHeight
  )
  $buffer = $host.ui.rawui.GetBufferContents($rec)

  for($i = 0; $i -lt $bufferHeight; $i++)
  {
    $span = [Text.StringBuilder]''
    $foreColor = $buffer[$i, 0].Foregroundcolor
    $backColor = $buffer[$i, 0].Backgroundcolor
    for($j = 0; $j -lt $bufferWidth; $j++)
    {
      $cell = $buffer[$i,$j]
      if (($cell.ForegroundColor -ne $foreColor) -or ($cell.BackgroundColor -ne $backColor))
      {
        $null = $html.Append(
"<span style='color:$foreColor;background:$backColor'>$($span)</span>"
        )
        $span = [Text.StringBuilder]''
        $foreColor = $cell.Foregroundcolor
        $backColor = $cell.Backgroundcolor
      }
      $null = $span.Append([Web.HttpUtility]::HtmlEncode($cell.Character))

    }
    $null = $html.Append(
"<span style='color:$foreColor;background:$backColor'>$($span)</span><br/>"
    )
  }

  $null = $html.Append("</pre>")
  $html.ToString()
}

Add-Type -AssemblyName System.Web

Get-ConsoleBufferAsHtml | Set-Content $env:temp\test.html  

Invoke-Item $env:temp\test.html  