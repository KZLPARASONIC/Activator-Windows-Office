# Writer by: Kerem Kaya
Write-Host "Destek icin: https://keremkaya.netlify.app" -ForegroundColor Magenta
Write-Host "----------------------------------------------" -ForegroundColor Green

function Activate-Windows {
    Write-Host "Windows aktivasyonu basladi..."

    # EditionID'yi al
    $edition = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").EditionID

    # Sürüm-key eşlemesi
    $keys = @{
        "Core"                  = "TX9XD-98N7V-6WMQ6-BX7FG-H8Q99"
        "CoreN"                 = "3KHY7-WNT83-DGQKR-F7HPR-844BM"
        "CoreSingleLanguage"    = "7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH"
        "CoreCountrySpecific"   = "PVMJN-6DFY6-9CCP6-7BKTT-D3WVR"
        "Professional"          = "W269N-WFGWX-YVC9B-4J6C9-T83GX"
        "ProfessionalN"         = "MH37W-N47XK-V7XM9-C7227-GCQG9"
        "Education"             = "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2"
        "EducationN"            = "2WH4N-8QGBV-H22JP-CT43Q-MDWWJ"
        "Enterprise"            = "NPPR9-FWDCX-D2C8J-H872K-2YT43"
        "EnterpriseN"           = "DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4"
    }

    if ($keys.ContainsKey($edition)) {
        $winKey = $keys[$edition]
        slmgr /ipk $winKey
        slmgr /skms kms8.msguides.com
        slmgr /ato
        Write-Host "Windows aktivasyonu tamamlandi." -ForegroundColor Green
    } else {
        Write-Host "Sürüm tanınmadı: $edition" -ForegroundColor Red
    }
}

function Deactivate-Windows {
    Write-Host "Windows lisansi kaldiriliyor..." -ForegroundColor Blue
    slmgr /upk
    slmgr /cpky
    Write-Host "Windows lisansi kaldirildi." -ForegroundColor Yellow
}

function Activate-Office {
    Write-Host "Office aktivasyonu basladi..."
    $officeKey = "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"
    $officePaths = @(
        "$env:ProgramFiles\Microsoft Office\Office16",
        "$env:ProgramFiles(x86)\Microsoft Office\Office16",
        "$env:ProgramFiles\Microsoft Office\Office15",
        "$env:ProgramFiles(x86)\Microsoft Office\Office15"
    )

    $activated = $false
    foreach ($path in $officePaths) {
        if (Test-Path $path) {
            Write-Host "Office yolu bulundu: $path"
            Set-Location $path
            cscript ospp.vbs /inpkey:$officeKey | Out-Null
            cscript ospp.vbs /sethst:kms8.msguides.com | Out-Null
            cscript ospp.vbs /act | Out-Null
            Write-Host "Office aktivasyonu tamamlandi." -ForegroundColor Green
            $activated = $true
            break
        }
    }
    if (-not $activated) {
        Write-Host "Office yuklu degil veya yol bulunamadi." -ForegroundColor Red
    }
}

function Deactivate-Office {
    Write-Host "Office lisansi kaldiriliyor..." -ForegroundColor DarkRed
    $officePaths = @(
        "$env:ProgramFiles\Microsoft Office\Office16",
        "$env:ProgramFiles(x86)\Microsoft Office\Office16",
        "$env:ProgramFiles\Microsoft Office\Office15",
        "$env:ProgramFiles(x86)\Microsoft Office\Office15"
    )

    $deactivated = $false
    foreach ($path in $officePaths) {
        if (Test-Path $path) {
            Set-Location $path
            cscript ospp.vbs /unpkey:XXXXX | Out-Null
            Write-Host "Office lisansi kaldirildi." -ForegroundColor Yellow
            $deactivated = $true
            break
        }
    }
    if (-not $deactivated) {
        Write-Host "Office bulunamadi." -ForegroundColor Red
    }
}

do {
    Write-Host "Windows ve Office Aktivasyon Scripti - Kerem Kaya" -ForegroundColor Magenta
    Write-Host "Destek icin: https://keremkaya.netlify.app" -ForegroundColor Magenta
    Write-Host "----------------------------------------------" -ForegroundColor Green
    Write-Host "Ne yapmak istiyorsun?" -ForegroundColor Green
    Write-Host "1) " -ForegroundColor Green -NoNewline; Write-Host "Sadece Windows Aktivasyon" -ForegroundColor Blue
    Write-Host "2) " -ForegroundColor Green -NoNewline; Write-Host "Sadece Office Aktivasyon" -ForegroundColor DarkRed
    Write-Host "3) " -ForegroundColor Green -NoNewline
    Write-Host "Windows" -ForegroundColor Blue -NoNewline
    Write-Host " ve " -NoNewline
    Write-Host "Office " -ForegroundColor DarkRed -NoNewline
    Write-Host "(ikisi birden aktivasyon)" -ForegroundColor Green
    Write-Host "4) " -ForegroundColor Green -NoNewline; Write-Host "Sadece Windows Lisans Kaldir" -ForegroundColor Blue
    Write-Host "5) " -ForegroundColor Green -NoNewline; Write-Host "Sadece Office Lisans Kaldir" -ForegroundColor DarkRed
    Write-Host "6) " -ForegroundColor Green -NoNewline
    Write-Host "Windows" -ForegroundColor Blue -NoNewline
    Write-Host " ve " -NoNewline
    Write-Host "Office " -ForegroundColor DarkRed -NoNewline
    Write-Host "(ikisi birden kaldir)" -ForegroundColor Green
    Write-Host "0) " -ForegroundColor Green -NoNewline; Write-Host "Cikis" -ForegroundColor Red

    Write-Host ""
    $secim = Read-Host "Seciminizi girin (0-6)"

    switch ($secim) {
        '1' { Activate-Windows }
        '2' { Activate-Office }
        '3' {
            Activate-Windows
            Activate-Office
        }
        '4' { Deactivate-Windows }
        '5' { Deactivate-Office }
        '6' {
            Deactivate-Windows
            Deactivate-Office
        }
        '0' { Write-Host "Cikis yapiliyor..." -ForegroundColor Yellow }
        default { Write-Warning "Gecersiz secim! Tekrar deneyin." }
    }
} while ($secim -ne '0')
