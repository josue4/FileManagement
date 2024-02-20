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


# English :

#AppFile
Management of archive entries and exits.
WPF application in PowerShell Xaml

## Compiling the application

To compile the application, you can use the [PS2EXE](https://github.com/MScholtes/PS2EXE) tool. Here are the steps to follow:

1. Install the `ps2exe` module with the following command:
   
PS C:> Install-Module ps2exe


2. Use the `Invoke-ps2exe` or `ps2exe` command to compile your PowerShell script into an executable file. For example, to compile a script named `source.ps1` into an executable file named `target.exe`, use one of the following commands:

PS C:> Invoke-ps2exe .\source.ps1 .\target.exe

or with:

PS C:>ps2exe.\source.ps1.\target.exe


## Password security

It is important to never store plaintext passwords in your compiled script. Indeed, it is possible to decompile a script compiled with the `-extract` option. Here is an example command to decompile the script stored in the `Output.exe` file:

Output.exe -extract:C:\Output.ps1
