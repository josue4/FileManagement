# FileManagement
PowerShell script to manage document input and output.


# AppFile
Gestion des entrées et sorties des archives.
Application WPF en PowerShell Xaml

## Compilation de l'application

Pour compiler l'application, vous pouvez utiliser l'outil [PS2EXE](https://github.com/MScholtes/PS2EXE). Voici les étapes à suivre :

1. Installez le module `ps2exe` avec la commande suivante :
   
PS C:> Install-Module ps2exe


2. Utilisez la commande `Invoke-ps2exe` ou `ps2exe` pour compiler votre script PowerShell en un fichier exécutable. Par exemple, pour compiler un script nommé `source.ps1` en un fichier exécutable nommé `target.exe`, utilisez l'une des commandes suivantes :

PS C:> Invoke-ps2exe .\source.ps1 .\target.exe

ou encore avec :

PS C:> ps2exe .\source.ps1 .\target.exe


## Sécurité du mot de passe

Il est important de ne jamais stocker de mots de passe en clair dans votre script compilé. En effet, il est possible de décompiler un script compilé avec l'option `-extract`. Voici un exemple de commande permettant de décompiler le script stocké dans le fichier `Output.exe` :

Output.exe -extract:C:\Output.ps1
