#!pwsh

$Datemaj = Get-Date -UFormat "%d %B %Y"

$filtreclasse = "^(user|computer|group|organizationalUnit|printQueue|contact)$"
$filtre = "^(\d{6}|zz\d{4})"
$reprapp = "C:\Users\181879\OneDrive - GHT Aude Pyrénées\Carnet CHP\CHP communication\"
$fichierrapport = "$($reprapp)l'A.D. du CHP en 2021.md"

$structure = Import-Csv "refpolesuf.csv" -Delimiter ";" -Encoding utf8

$touslesobjets = Get-ADObject -Filter *
write-host "extraction du $($Datemaj) des $(($touslesobjets|Measure-Object).Count) objets de l'ad"

$mesclasses = $touslesobjets | Where-Object {$_.objectclass -match $filtreclasse} |  Group-Object -Property objectclass | Sort-Object count -Descending | Select-Object @{l="classe d'objet";e={$_.name}}, @{l="nombre";e={$_.count}}

$mesclasses | Export-Csv "compo.csv" -Encoding utf8 -Delimiter ","

$dgos = Import-Csv -Path .\codesdgos.csv -Encoding utf8 -Delimiter ';' -Header 'code', 'libelle'

write-host "🧮 création des tableaux :"
write-host "⇒ des utilisateurs"

$touslesutilisateurs = Get-ADUser -Filter *
$utilphysique = $touslesutilisateurs | Where-Object {$_.SamAccountName -match $filtre}
$utilanonyme = $touslesutilisateurs | Where-Object {$_.SamAccountName -notmatch $filtre}
$physiquesactifs = $utilphysique | Where-Object {$_.enabled -eq $true}
$physiquesinactifs = $utilphysique | Where-Object {$_.enabled -eq $false}
$anonymesactifs = $utilanonyme | Where-Object {$_.enabled -eq $true}
$anonymesinactifs = $utilanonyme | Where-Object {$_.enabled -eq $false}

$tableauutils = @"
,actifs,inactifs,total
**nominatifs**,$($physiquesactifs.count),$($physiquesinactifs.count),$($utilphysique.count)
**anonymes**,$($anonymesactifs.count),$($anonymesinactifs.count),$($utilanonyme.count)
**total**,$($physiquesactifs.count + $anonymesactifs.count),$($physiquesinactifs.count + $anonymesinactifs.count),$($touslesutilisateurs.count)
"@

$tableauutils | Out-File -Encoding utf8 -Force -Path "$($reprapp)utils.csv"

write-host "⇒ des informations de contact"

$contphys = $physiquesactifs | Get-ADUser -Properties homephone, mobilephone, officephone, mail, emailAddress
$contanos = $anonymesactifs | Get-ADUser -Properties homephone, mobilephone, officephone , mail, emailAddress

$fixephys = $contphys | Where-Object {$null -ne $_.officephone}
$mobphys = $contphys | Where-Object {$null -ne $_.mobilephone}
$homephys = $contphys | Where-Object {$null -ne $_.homephone}
$mailphys = $contphys | Where-Object {$null -ne $_.mail -and $null -ne $_.emailAddress}

$fixeanos = $contanos | Where-Object {$null -ne $_.officephone}
$mobanos = $contanos | Where-Object {$null -ne $_.mobilephone}
$homeanos = $contanos | Where-Object {$null -ne $_.homephone}
$mailanos = $contanos | Where-Object {$null -ne $_.mail -and $null -ne $_.emailAddress}


$tableaucontacts = @"
type de compte,avec tél. fixe,avec tél. mobile,avec tél. domicile, avec mail
nominatif,$(($fixephys|Measure-Object).Count),$(($mobphys|Measure-Object).Count),$(($homephys|Measure-Object).Count),$(($mailphys|Measure-Object).Count)
anonyme,$(($fixeanos|Measure-Object).Count),$(($mobanos|Measure-Object).Count),$(($homeanos|Measure-Object).Count),$(($mailanos|Measure-Object).Count)
"@

$tableaucontacts | Out-File -Encoding utf8 -Force -Path "$($reprapp)contacts.csv"
write-host "⇒ des localisations"

$locaphys = $physiquesactifs | Get-ADUser -Properties office, streetAddress, postalCode, City
$locaanos = $anonymesactifs | Get-ADUser -Properties office, streetAddress, postalCode, City

$bureauphys = $locaphys | Where-Object {$null -ne $_.office}
$cpphys = $locaphys | Where-Object {$null -ne $_.postalCode}
$ruephys = $locaphys | Where-Object {$null -ne $_.streetAddress}
$villephys = $locaphys | Where-Object {$null -ne $_.City}

$bureauanos = $locaanos | Where-Object {$null -ne $_.office}
$cpanos = $locaanos | Where-Object {$null -ne $_.postalCode}
$rueanos = $locaanos | Where-Object {$null -ne $_.streetAddress}
$villeanos = $locaanos | Where-Object {$null -ne $_.City}

$tableaulocal = @"
type de compte,bureau,ville,rue,code postal
nominatif,$(($bureauphys | Measure-Object).Count),$(($villephys | Measure-Object).Count),$(($ruephys | Measure-Object).Count),$(($cpphys | Measure-Object).Count)
anonyme,$(($bureauanos | Measure-Object).Count),$(($villeanos | Measure-Object).Count),$(($rueanos | Measure-Object).Count),$(($cpanos | Measure-Object).Count)
"@

$tableaulocal | Out-File -Encoding utf8 -Force -Path "$($reprapp)local.csv"

write-host "⇒ de l'organisation"

$orgaphys = $physiquesactifs | Get-ADUser -Properties department, departmentNumber, organization, company, title, personalTitle, DistinguishedName
$orgaanos = $anonymesactifs | Get-ADUser -Properties department, departmentNumber, organization, company, title, personalTitle, DistinguishedName

$depphys = $orgaphys | Where-Object {$null -ne $_.department}
$orgphys = $orgaphys | Where-Object {$null -ne $_.organization}
$socphys = $orgaphys | Where-Object {$null -ne $_.company}
$titphys = $orgaphys | Where-Object {$null -ne $_.title}
$foncphys = $orgaphys | Where-Object {$null -ne $_.personalTitle}

$foncphys | Add-Member -Name "Métier DGOS" -MemberType Noteproperty -Value "" -Force

$depanos = $orgaanos | Where-Object {$null -ne $_.department}
$organos = $orgaanos | Where-Object {$null -ne $_.organization}
$socanos = $orgaanos | Where-Object {$null -ne $_.company}
$titanos = $orgaanos | Where-Object {$null -ne $_.title}
$foncanos = $orgaanos | Where-Object {$null -ne $_.personalTitle}
$foncanos | Add-Member -Name "Métier DGOS" -MemberType Noteproperty -Value "" -Force

foreach ($m in $dgos) {
    foreach ($ua in $foncanos) {
        $ua.personalTitle = $ua.personalTitle.Trim()
        if ($ua.personalTitle -eq $m.code) {
            $ua."Métier DGOS" = $m.libelle
        }
    }
    foreach ($up in $foncphys) {
        $up.personalTitle = $up.personalTitle.Trim()
        if ($up.personalTitle -eq $m.code) {
            $up."Métier DGOS" = $m.libelle
        }
    }
}

$tableauorga = @"
type de compte,service, organisation,société,titre,fonction
nominatif,$(($depphys | Measure-Object).Count),$(($orgphys | Measure-Object).Count),$(($socphys | Measure-Object).Count),$(($titphys | Measure-Object).Count),$(($foncphys | Measure-Object).Count)
anonyme,$(($depanos | Measure-Object).Count),$(($organos | Measure-Object).Count),$(($socanos | Measure-Object).Count),$(($titanos | Measure-Object).Count),$(($foncanos | Measure-Object).Count)
"@

$tableauorga | Out-File -Encoding utf8 -Force -Path "$($reprapp)orga.csv"

$socphys + $socanos | Group-Object company | sort-object count -descending | Select-Object @{l="société";e={$_.name}}, @{l="nb. de comptes";e={$_.count}} | Export-Csv "$($reprapp)sociétés.csv" -Encoding utf8 -Delimiter ","
$titphys + $titanos | Group-Object title | sort-object count -descending | Select-Object @{l="grade";e={$_.name}}, @{l="nb. de comptes";e={$_.count}} | Export-Csv "$($reprapp)annexes title.csv" -Encoding utf8 -Delimiter ","
$foncphys + $foncanos | Group-Object "Métier DGOS", personalTitle | sort-object count -descending | Select-Object @{l="Nom métier";e={($_.name -split ",")[0]}}, @{l="Code métier";e={($_.name -split ",")[1]}} , @{l="nb. de comptes";e={$_.count}} | Export-Csv "$($reprapp)annexes personalTitle.csv" -Encoding utf8 -Delimiter ","
$depphys + $depanos | Group-Object department, departmentNumber | sort-object {($_.name -split ",")[1]} | Select-Object @{l="service";e={($_.name -split ",")[0]}},@{l="code";e={($_.name -split ",")[1] -replace "\{|\}",""}}, @{l="nb. de comptes";e={$_.count}} | Export-Csv "$($reprapp)annexes department.csv" -Encoding utf8 -Delimiter ","

write-host "⚙ recherche des incohérences"

$incohérences = $orgaphys + $organos | Where-Object {$_.distinguishedname -notmatch $_.departmentNumber -or $_.department -notmatch $_.departmentNumber}
$incohérences | Select-Object department, @{l="departmentNumber";e={$_.departmentNumber}[0]}, DistinguishedName | Sort-Object departmentNumber | Export-Csv "$($reprapp)annexes comptes incohérents.csv" -Encoding utf8 -Delimiter ","

$nbinco = ($incohérences | Measure-Object).Count


write-host "✏  mise-à-jour du fichier rapport"

$yaml = @("---","Objets: $(($touslesobjets|Measure-Object).Count)","Datemaj: $($Datemaj)","Incoh: $($nbinco)","lignesficom: $(($structure|Measure-Object).Count)","---")

$lerapport = Get-Content -Path $fichierrapport -Encoding utf8

$lerapporttemp = $lerapport[6..($lerapport.count - 1)]

$lerapport = $yaml
$lerapport += $lerapporttemp

Clear-Content $fichierrapport

$lerapport | ForEach-Object {add-Content -Path $fichierrapport -value $_ -Encoding utf8}
