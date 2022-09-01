# Souvenirs hospitalier

Ce *repo* est une compilation d’outils créés pour me faciliter la vie lors de mon passage au centre hospitalier de Perpignan

## audit_info_comptes_ad.ps1

Script Powershell extrayant les données d’annuaire de l’*Active Directory*. Il requiert le module [PowerShell Active Directory](https://docs.microsoft.com/en-us/powershell/module/activedirectory/?view=windowsserver2022-ps).

### variables

- $filtreclasse : pour définir les classes d’objet à incorporer
- $filtre : pour définir le motif correspondant aux identifiants des utilisateur
- $reprapp : chemin absolu vers le répertoire d’enregistrement des exports
- $fichierrapport : nom de mon fichier MarkDown de rapport principal

### fichiers nécessaires

- refpolesuf.csv : fichier contenant les poles et unités fonctionelles
- codesdgos.csv : fichier compilant les codes des métiers de la DGOS

### fichiers créés

- compo.csv
- utils.csv
- contacts.csv
- local.csv
- orga.csv
- sociétés.csv
- annexes title.csv
- annexes personalTitle.csv
- annexes department.csv
- annexes comptes incohérents.csv
