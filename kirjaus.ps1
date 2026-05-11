try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop
} catch {
    Write-Host "Kriittinen virhe: Kirjastojen lataus epäonnistui. $($_.Exception.Message)" -ForegroundColor Red
    pause
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

$dataFolder = Join-Path ([Environment]::GetFolderPath("MyDocuments")) "PlatinumDashboard"
if (-not (Test-Path $dataFolder)) { 
    New-Item -ItemType Directory -Path $dataFolder | Out-Null 
}

$configPath = Join-Path $dataFolder "config_safe.json"
$historyPath = Join-Path $dataFolder "historia_safe.json"

$script:googleScriptUrl = "https://script.google.com/macros/s/AKfycbwhRSWotX4sJv4u8834Bq6RmZT74QjmdqDsMQSbER7LPTctwYOA7cbs9-sQempfpT4p/exec"

Write-Host "--- DEBUG: Käynnistetään Dashboard v5.0 ---" -ForegroundColor Cyan

function Save-All {
    $script:asetukset | ConvertTo-Json | Out-File $configPath -Encoding utf8
    $script:historia | ConvertTo-Json | Out-File $historyPath -Encoding utf8
}

if (Test-Path $configPath) { 
    $tempJson = Get-Content $configPath | ConvertFrom-Json
    $script:asetukset = @{}
    foreach($prop in $tempJson.psobject.Properties) { $script:asetukset[$prop.Name] = $prop.Value }
} else { $script:asetukset = @{} }

$defaults = @{
    tuntipalkka = 15.0; tuotebonus = 150.0; palvelubonus = 100.0; 
    teema = "Tumma"; valuutta = "€"; tuhatErotin = " ";
    ikkunaLeveys = 950; ikkunaKorkeus = 950;
    otsikkoVari = "#0078D4"; apiKey = ""
}
foreach ($key in $defaults.Keys) {
    if (-not $script:asetukset.ContainsKey($key)) { $script:asetukset[$key] = $defaults[$key] }
}

if ($script:asetukset["apiKey"] -eq "") {
    $prompt = "Anna API-avain (salasana) pilviyhteyttä varten:"
    $inputKey = [Microsoft.VisualBasic.Interaction]::InputBox($prompt, "Kirjaudu Pilveen", "")
    if ($inputKey -eq "") { exit }
    $script:asetukset["apiKey"] = $inputKey
    Save-All
}

if (Test-Path $historyPath) {
    $tempHist = Get-Content $historyPath | ConvertFrom-Json
    $script:historia = if ($null -eq $tempHist) { @() } elseif ($tempHist -is [Array]) { $tempHist } else { @($tempHist) }
} else { $script:historia = @() }

function Get-SafeNum ([object]$val) {
    if ($null -eq $val -or $val.ToString().Trim() -eq "") { return 0.0 }
    $clean = $val.ToString().Replace(" ", "").Replace(",", ".")
    $out = 0.0
    if ([double]::TryParse($clean, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$out)) { return $out }
    return 0.0
}

function Format-Number ($num) {
    $n = Get-SafeNum $num
    $culture = [System.Globalization.CultureInfo]::CreateSpecificCulture("fi-FI")
    return $n.ToString("N2", $culture)
}

function Invoke-Cloud ($data) {
    Write-Host "--- DEBUG: Lähetetään tietoja pilveen... ---" -ForegroundColor Yellow
    try {
        $payload = @{}
        foreach($prop in $data.psobject.Properties) { 
            if ($prop.Name -eq "s" -or $prop.Name -eq "t") {
                $cleanNum = Get-SafeNum $prop.Value
                $payload[$prop.Name] = $cleanNum.ToString([System.Globalization.CultureInfo]::InvariantCulture)
            } else {
                $payload[$prop.Name] = $prop.Value 
            }
        }
        $payload["auth"] = $script:asetukset["apiKey"]
        $json = $payload | ConvertTo-Json
        
        Write-Host "DEBUG Payload: $json" -ForegroundColor Gray
        $response = Invoke-RestMethod -Uri $script:googleScriptUrl -Method Post -Body $json -ContentType "application/json"
        Write-Host "DEBUG Pilven vastaus: $response" -ForegroundColor Green
        return $response
    } catch { 
        Write-Host "DEBUG VIRHE LÄHETYKSESSÄ: $($_.Exception.Message)" -ForegroundColor Red
        return "Yhteysvirhe" 
    }
}

$f = New-Object Windows.Forms.Form
$f.Text = "FiveM Platinum Dashboard v4.9"
$f.Size = New-Object Drawing.Size([int]$script:asetukset["ikkunaLeveys"], [int]$script:asetukset["ikkunaKorkeus"])
$f.StartPosition = "CenterScreen"
$f.TopMost = $true

$bgHtml = if($script:asetukset["teema"] -eq "Tumma") { "#121212" } else { "#FFFFFF" }
$f.BackColor = [Drawing.ColorTranslator]::FromHtml($bgHtml)
$textC = if($script:asetukset["teema"] -eq "Tumma") { [Drawing.Color]::White } else { [Drawing.Color]::Black }


$gbCalc = New-Object Windows.Forms.GroupBox; $gbCalc.Text = "Palkkalaskuri aikavälille"; $gbCalc.Location = "30,20"; $gbCalc.Size = "400,100"; $gbCalc.ForeColor = $textC
$f.Controls.Add($gbCalc)

$tAlku = New-Object Windows.Forms.TextBox; $tAlku.Text = (Get-Date).AddMonths(-1).ToString("dd.MM.yyyy"); $tAlku.Location = "20,40"; $tAlku.Width = 90; $gbCalc.Controls.Add($tAlku)
$tLoppu = New-Object Windows.Forms.TextBox; $tLoppu.Text = (Get-Date).ToString("dd.MM.yyyy"); $tLoppu.Location = "120,40"; $tLoppu.Width = 90; $gbCalc.Controls.Add($tLoppu)
$btnLaske = New-Object Windows.Forms.Button; $btnLaske.Text = "LASKE"; $btnLaske.Location = "230,38"; $btnLaske.BackColor = [Drawing.Color]::SeaGreen; $btnLaske.ForeColor = [Drawing.Color]::White; $gbCalc.Controls.Add($btnLaske)
$lblRes = New-Object Windows.Forms.Label; $lblRes.Text = "Tulos: 0 €"; $lblRes.Location = "20,75"; $lblRes.AutoSize = $true; $gbCalc.Controls.Add($lblRes)

$btnLaske.Add_Click({
    try {
        $d1 = [datetime]::Parse($tAlku.Text.Trim()); $d2 = [datetime]::Parse($tLoppu.Text.Trim())
        $total = 0.0
        $script:historia | ForEach-Object { 
            $p = [datetime]::Parse($_.pvm)
            if ($p -ge $d1 -and $p -le $d2) { $total += Get-SafeNum $_.s }
        }
        $lblRes.Text = "Palkka yhteensä: " + (Format-Number $total) + " " + $script:asetukset["valuutta"]
    } catch { $lblRes.Text = "Virhe päivämäärissä!" }
})


$btnSet = New-Object Windows.Forms.Button; $btnSet.Text = "ASETUKSET"; $btnSet.Location = "780,30"; $btnSet.Size = "130,35"; $btnSet.BackColor = [Drawing.Color]::FromArgb(51,51,51); $btnSet.ForeColor = [Drawing.Color]::White; $f.Controls.Add($btnSet)
$btnClear = New-Object Windows.Forms.Button; $btnClear.Text = "TYHJENNÄ PILVI"; $btnClear.Location = "780,75"; $btnClear.Size = "130,35"; $btnClear.BackColor = [Drawing.Color]::FromArgb(68,34,34); $btnClear.ForeColor = [Drawing.Color]::White; $f.Controls.Add($btnClear)


$gbNew = New-Object Windows.Forms.GroupBox; $gbNew.Text = "LISÄÄ UUSI TYÖVUORO"; $gbNew.Location = "30,140"; $gbNew.Size = "880,240"; $gbNew.ForeColor = $textC
$f.Controls.Add($gbNew)

$inLabels = @("Päivämäärä:", "Aloitus (HH:MM):", "Lopetus (HH:MM):", "Työkalusarjat (kpl):", "Tuunaukset (kpl):")
$txts = @()
for($i=0;$i -lt 5;$i++) {
    $l = New-Object Windows.Forms.Label; $l.Text = $inLabels[$i]; $l.Location = "20,$(40+($i*35))"; $l.AutoSize = $true; $gbNew.Controls.Add($l)
    $t = New-Object Windows.Forms.TextBox; $t.Location = "180,$(38+($i*35))"; $t.Width = 100; $gbNew.Controls.Add($t); $txts += $t
}
$txts[0].Text = (Get-Date).ToString("dd.MM.yyyy")
$txts[1].Text = "08:00"; $txts[2].Text = "16:00"; $txts[3].Text = "0"; $txts[4].Text = "0"

$btnSave = New-Object Windows.Forms.Button
$btnSave.Text = "TALLENNA JA LÄHETÄ"; $btnSave.Location = "350,40"; $btnSave.Size = "480,170"
$btnSave.BackColor = [Drawing.ColorTranslator]::FromHtml($script:asetukset["otsikkoVari"])
$btnSave.ForeColor = [Drawing.Color]::White; $btnSave.Font = New-Object Drawing.Font("Segoe UI", 14, [Drawing.FontStyle]::Bold)
$gbNew.Controls.Add($btnSave)


$lv = New-Object Windows.Forms.ListView; $lv.View = "Details"; $lv.Location = "30,400"; $lv.Size = "880,450"; $lv.FullRowSelect = $true
$lv.BackColor = [Drawing.ColorTranslator]::FromHtml("#222222"); $lv.ForeColor = [Drawing.Color]::White
$lv.Columns.Add("Pvm", 120) | Out-Null; $lv.Columns.Add("Aloitus", 100) | Out-Null; $lv.Columns.Add("Lopetus", 100) | Out-Null; $lv.Columns.Add("Tunnit", 100) | Out-Null; $lv.Columns.Add("Summa", 150) | Out-Null
$f.Controls.Add($lv)

function Refresh-List {
    $lv.Items.Clear()
    foreach($h in $script:historia) {
        $item = New-Object Windows.Forms.ListViewItem([string]$h.pvm)
        [void]$item.SubItems.Add([string]$h.alku); [void]$item.SubItems.Add([string]$h.loppu); [void]$item.SubItems.Add("$([string]$h.t) h"); [void]$item.SubItems.Add((Format-Number $h.s) + " " + $script:asetukset["valuutta"])
        [void]$lv.Items.Add($item)
    }
}


$btnSave.Add_Click({
    Write-Host "--- DEBUG: Tallennus aloitettu ---" -ForegroundColor Blue
    try {
        $pvmTxt = $txts[0].Text.Trim()
        $alkuTxt = $txts[1].Text.Trim()
        $loppuTxt = $txts[2].Text.Trim()
        $tuotteet = $txts[3].Text.Trim()
        $palvelut = $txts[4].Text.Trim()


        $t1 = [datetime]::ParseExact($alkuTxt, "HH:mm", $null)
        $t2 = [datetime]::ParseExact($loppuTxt, "HH:mm", $null)
        if($t2 -le $t1){ $t2 = $t2.AddDays(1) }
        
        $kesto = ($t2 - $t1).TotalHours
        $brutto = ($kesto * (Get-SafeNum $script:asetukset["tuntipalkka"])) + 
                  ((Get-SafeNum $tuotteet) * (Get-SafeNum $script:asetukset["tuotebonus"])) + 
                  ((Get-SafeNum $palvelut) * (Get-SafeNum $script:asetukset["palvelubonus"]))
        
        $uusi = [PSCustomObject]@{ 
            pvm = $pvmTxt; alku = $alkuTxt; loppu = $loppuTxt; 
            t = [math]::Round($kesto,2); k = [int](Get-SafeNum $tuotteet); 
            p = [int](Get-SafeNum $palvelut); s = [math]::Round($brutto,2) 
        }
        

        if ($null -eq $script:historia) { $script:historia = @() }
        $tempList = [System.Collections.Generic.List[object]]::new()
        foreach ($item in $script:historia) { [void]$tempList.Add($item) }
        [void]$tempList.Add($uusi)
        $script:historia = $tempList.ToArray()

        Save-All
        Refresh-List
        
        $status = Invoke-Cloud $uusi
        [Windows.Forms.MessageBox]::Show("Tila: $status")
    } catch { 
        Write-Host "--- DEBUG VIRHE ---" -ForegroundColor Red
        Write-Host "Virheen syy: $($_.Exception.Message)" -ForegroundColor Red
        [Windows.Forms.MessageBox]::Show("Virhe syötteissä! Katso konsoli.") 
    }
})


$btnSet.Add_Click({
    $f.TopMost = $false
    $sF = New-Object Windows.Forms.Form; $sF.Text = "Asetukset"; $sF.Size = "450,600"; $sF.StartPosition = "CenterParent"
    $flow = New-Object Windows.Forms.FlowLayoutPanel; $flow.Dock = "Fill"; $flow.AutoScroll = $true; $sF.Controls.Add($flow)
    $settingsList = @(@("Tuntipalkka (€)", "tuntipalkka"),@("Tuotebonus (€)", "tuotebonus"),@("Palvelubonus (€)", "palvelubonus"),@("API-avain (Pilvi)", "apiKey"),@("Teema (Tumma/Vaalea)", "teema"),@("Valuuttamerkki", "valuutta"),@("IkkunaLeveys", "ikkunaLeveys"),@("IkkunaKorkeus", "ikkunaKorkeus"),@("Otsikkoväri (Hex)", "otsikkoVari"))
    foreach($item in $settingsList) {
        $pan = New-Object Windows.Forms.Panel; $pan.Size = "400,40"
        $lab = New-Object Windows.Forms.Label; $lab.Text = $item[0]; $lab.Location = "10,10"; $lab.Width = 150; $pan.Controls.Add($lab)
        $tex = New-Object Windows.Forms.TextBox; $tex.Text = $script:asetukset[$item[1]]; $tex.Location = "180,8"; $tex.Name = $item[1]; $pan.Controls.Add($tex)
        $flow.Controls.Add($pan)
    }
    $bSave = New-Object Windows.Forms.Button; $bSave.Text = "TALLENNA JA PÄIVITÄ"; $bSave.Dock = "Bottom"; $bSave.Height = 45
    $bSave.Add_Click({
        foreach($c in $flow.Controls) { foreach($cc in $c.Controls) { if($cc -is [Windows.Forms.TextBox]){ $script:asetukset[$cc.Name] = $cc.Text } } }
        Save-All; $sF.Close(); $f.Close()
    })
    $sF.Controls.Add($bSave); $sF.Add_FormClosing({ $f.TopMost = $true }); $sF.ShowDialog()
})

$btnClear.Add_Click({
    if([Windows.Forms.MessageBox]::Show("Tyhjennetäänkö pilvi?", "Vahvistus", "YesNo") -eq "Yes") {
        $res = Invoke-Cloud @{ action = "clear" }
        [Windows.Forms.MessageBox]::Show("$res")
    }
})

Refresh-List
$f.ShowDialog()
