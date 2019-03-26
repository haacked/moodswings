$WHITE = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::White)
$BLACK = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Black)
$RED = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Red)

$properties = $dte.Properties("FontsAndColors", "TextEditor")
$fontsAndColorsItems = $properties.Item("FontsAndColorsItems").Object
$plainTextColors = $fontsAndColorsItems.Item("Plain Text")

$moods = @{}

$moods.Default = {
	$properties.Item("FontSize").Value = 10
	$plainTextColors.Background = $WHITE
	$plainTextColors.Foreground = $BLACK
}
$moods.Presentation = {
	$properties.Item("FontSize").Value = 15
	$plainTextColors.Background = $WHITE
	$plainTextColors.Foreground = $BLACK
}; 
$moods.Coding = {
	$properties.Item("FontSize").Value = 10
	$plainTextColors.Background = $BLACK
	$plainTextColors.Foreground = $WHITE
}
$moods.Rick = {
	$rickScript = Join-Path $PSScriptRoot rick.ps1
	Start-Process powershell -ArgumentList "-noprofile -noexit -File $rickScript"
}
$moods.Wild = {
	$properties.Item("FontSize").Value = 20
	$plainTextColors.Background = $RED
	$plainTextColors.Foreground = $WHITE
}


function Set-Mood($Mood) {
	&$moods[$Mood]
}

Register-TabExpansion 'Set-Mood' @{
	'Mood' = { $moods.Keys }
}

Export-ModuleMember Set-Mood