Param(
  [switch]$light,
  [switch]$dark
)

$theme_literal_path = 'HKCU:Software\Microsoft\Windows\CurrentVersion\Themes\Personalize'

$app_theme = Get-ItemPropertyValue -LiteralPath $theme_literal_path -Name AppsUseLightTheme
$sys_theme = Get-ItemPropertyValue -LiteralPath $theme_literal_path -Name SystemUsesLightTheme

$theme = 1 - $app_theme

If ($light) {
  $theme = 1
} ElseIf ($dark) {
  $theme = 0
}

Set-ItemProperty -LiteralPath $theme_literal_path -Name AppsUseLightTheme -Value $theme
Set-ItemProperty -LiteralPath $theme_literal_path -Name SystemUsesLightTheme -Value $theme

# restart explorer.exe
Stop-Process -name explorer -force
