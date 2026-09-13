# 1. Шлях до підписаного XLL
$xllPath = "$env:APPDATA\Microsoft\AddIns\ZsuTools-AddIn64-packed.xll"

if (-not (Test-Path $xllPath)) {
    Write-Error "Файл $xllPath не знайдено!"
    return
}

# 2. Створення правила Publisher для сертифіката
$signerRule = New-CIPolicyRule -DriverFilePath $xllPath -Level Publisher

# 3. Шляхи
$xmlPath = "C:\ZsuTools_SAC_Exemption.xml"
$cipPath = "C:\ZsuTools_SAC_Exemption.cip"

# 4. Генерація базової Supplemental політики (прапорець -MultiplePolicyFormat робить її сумісною)
New-CIPolicy -Rules $signerRule -FilePath $xmlPath -UserPEs -MultiplePolicyFormat

# 5. Встановлення параметрів UMCI
Set-RuleOption -FilePath $xmlPath -Option 0   # Enabled:UMCI
Set-RuleOption -FilePath $xmlPath -Option 18  # Disabled:Runtime FilePath Protection

# 6. Генеруємо ID та прикликуємо до базової політики Smart App Control
$PolicyID = [guid]::NewGuid().ToString("B").ToUpper()
Set-CIPolicyIdInfo -FilePath $xmlPath -PolicyID $PolicyID -PolicyName "ZsuTools_SAC_Override" -SupplementsBasePolicyID "{0283ac0f-fff1-49ae-ada1-8a933130cad6}"

# 7. Компіляція бінарника
ConvertFrom-CIPolicy $xmlPath $cipPath

# 8. Копіювання в активний системний каталог WDAC
$targetGuid = $PolicyID.Trim("{}")
$sysCiPath = "$env:windir\System32\CodeIntegrity\CiPolicies\Active\{$targetGuid}.cip"
Copy-Item -Path $cipPath -Destination $sysCiPath -Force

# 9. Реєстрація через citool
citool.exe -i $cipPath
