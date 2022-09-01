#! pwsh

Clear-Host

$dlc = (Get-Date).AddDays(-15)
$personneslibreacces = "https://service.annuaire.sante.fr/annuaire-sante-webservices/V300/services/extraction/PS_LibreAcces"
$derarcrpps = "directfromasip.zip"
$reparcrpps = "directfromasip"

$filtresmetiers = "Médecine Générale$"
$filtresdepartement = "^(66|11)"

$finess = "660000084"


Write-Host "Vérification de l´archive de l´ASIP"
if (!(Test-Path $derarcrpps)) {
    Write-Host "❌ aucune archive trouvée, téléchargement déclenché"
    Invoke-WebRequest -Uri $personneslibreacces -OutFile $derarcrpps

} elseif ($dlc -gt (Get-ChildItem $derarcrpps).LastWriteTime) {
    Remove-Item  $derarcrpps
    Write-Host "❌ archive présente mais ancienne, téléchargement déclenché"
    Invoke-WebRequest -Uri $personneslibreacces -OutFile $derarcrpps
} else {
    Write-Host "✅ archive présente et à jour"
}

if (!(Test-Path $reparcrpps)) {
    Expand-Archive -Path $derarcrpps -DestinationPath $reparcrpps
} else {
    Write-Host "les listes ont déjà été décompressées dans « $($reparcrpps) »"
    Get-ChildItem $reparcrpps
}

Write-Host "⚙ chargement des données des professionnels de santé du département"


$lespersonnesf = Get-ChildItem -Path "$($reparcrpps)/PS_LibreAcces_Personne_activite_*"
$lespersonnest = Import-Csv $lespersonnesf -Encoding utf8 -Delimiter "|" | Where-Object {$_."Code postal (coord. structure)" -match $filtresdepartement -and $_."Libellé savoir-faire" -match $filtresmetiers}

$auchp = $lespersonnest | Where-Object {$_."Numéro FINESS site"-eq $finess}  | Select-Object "Identifiant PP", "Code civilité d'exercice","Libellé civilité d'exercice","Code civilité","Libellé civilité","Nom d'exercice","Prénom d'exercice","Code profession","Libellé profession"
$nonchp = $lespersonnest | Where-Object {$_."Numéro FINESS site" -ne $finess} | Select-Object "Identifiant PP", "Code civilité d'exercice","Libellé civilité d'exercice","Code civilité","Libellé civilité","Nom d'exercice","Prénom d'exercice","Code profession","Libellé profession","Libellé mode exercice","Adresse e-mail (coord. structure)","Téléphone (coord. structure)","Code postal (coord. structure)"

Write-Host "⚗ exports des tableaux statistiques"

$lespersonnest | Group-Object "Libellé profession" | Select-Object @{l="profession";e={$_.Name}}, @{l="population";e={$_.Count}} | Export-Csv -Encoding utf8 -Delimiter ";" -Path poptot.csv
$auchp | Group-Object "Libellé profession" | Select-Object @{l="profession";e={$_.Name}},"Code profession", @{l="population";e={$_.Count}} | Export-Csv -Encoding utf8 -Delimiter ";" -Path popchp.csv

Write-Host "⚗ exports des listes des professionnels"

$nonchp | Where-Object { $_."Code postal (coord. structure)" -match "^11"} | Export-Csv -Path "laliste11.csv" -Delimiter ";" -Encoding utf8
$nonchp | Where-Object { $_."Code postal (coord. structure)" -match "^66"} | Export-Csv -Path "laliste66.csv" -Delimiter ";" -Encoding utf8
$auchp | Export-Csv -Path "lalisteCHP.csv" -Delimiter ";" -Encoding utf8