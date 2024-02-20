

Set-Location -Path "Se positionner dans le répertoire de la solution"
Get-Location

Add-Type -AssemblyName PresentationFramework

$xamlFile ="Chemin de l'interface principale" # .\MainWindow.xaml


$inputXAML = Get-Content -Path $xamlFile -Raw
$inputXAML=$inputXAML -replace 'mc:Ignorable="d"','' -replace "x:N","N" -replace '^<Win.*','<Window'
[XML]$XAML=$inputXAML

$reader = New-Object System.Xml.XmlNodeReader $XAML
try{
    $psform=[Windows.Markup.XamlReader]::Load($reader)
}catch{
    Write-Host $_.Exception
    throw
}

$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    try{
        Set-Variable -Name "var_$($_.Name)" -Value $psform.FindName($_.Name) -ErrorAction Stop
    }catch{
        throw
    }
} 

Get-Variable var_*


$ListData= Import-Csv -Path "Chemin du CSV" -Delimiter ',' 

# A chaque changement de valeur dans le combobox, la fonction Get-Liste est appelÃ©e
$var_ddlHeader.Add_SelectionChanged({Get-Liste})

#La fonction Get-Liste permet de remplir le combobox avec les valeurs de la colonne sÃ©lectionnÃ©e
function Get-Liste {
    
    $var_ddlfiltre.Items.Clear()
    $entete = $var_ddlHeader.SelectedItem
    if ($entete -eq $null) {
        $entete ="date"
    }

    $ListData | ForEach-Object {$var_ddlfiltre.Items.Add($_.$entete)}
    

    # Permet de mettre un scroll dans le combobox filtrer
    $var_ddlfiltre.IsTextSearchEnabled = $True
    $var_ddlfiltre.IsEditable = $True
    $var_ddlfiltre.MaxDropDownHeight = "100"
}



# $MatriculeDomaine = whoami
# $index = $MatriculeDomaine.IndexOf('\')
# $Matricule = $MatriculeDomaine.Substring($index  + 1)
# write-host $Matricule

function Get-Autorisation {
    
    # fonction de recherche de L'EDS de l'utlisateur qui ouvre le formulaire dans la BDD
    # Write-Host
    # Write-Host " fonction SQL" 

    # # http://la-capsule.fr/articles/francais/utiliser-powershell-pour-se-connecter-a-sql-server-et-lire-les-donnees-d-une-table/
    # [string]$DBServer = ""
    # [string]$DBName = ""
    # [string]$SQLServerLogin = ""
    # [string]$SQLServerLoginPW = ""
      
    # $sqlConnection = New-Object System.Data.SqlClient.SqlConnection
    # $sqlConnection.ConnectionString = "Server=$DBServer; Database=$DBName; User Id=$SQLServerLogin; Password=$SQLServerLoginPW;"
    # $sqlConnection.Open()
    
    # #### VÃ©rifier que la connexion fonctionne avant d'aller plus loin
    # if ($sqlConnection.State -ne [Data.ConnectionState]::Open) {
        
    #     write-host "Impossible d'ouvrir la connexion."
        
    #     Exit
    # }
    
    # # SELECT CodeXXX from [XXXX] where ident = FXXXXXXXX
    # try{  
    #     $resultsDataTable = New-Object System.Data.DataTable
    #     $Command = New-Object System.Data.SQLClient.SQLCommand 
    #     $Command.Connection = $sqlConnection
    #     $Command.CommandText = "SELECT , FROM []"
    #     $Reader = $Command.ExecuteReader() 
    #     $resultsDataTable.Load($Reader) 
         
    #     foreach ($Row in $resultsDataTable.Rows){ 
    #         if ($($Row[0]) -eq $Matricule) {
    #         $EDS = $($Row[1])
    #         write-host "Mon EDS $EDS"
    #         }
    #     }
         
    #     $Reader.Close()
    #     $Reader.Dispose()
    #     $resultsDataTable.Dispose()
    #     }catch {  
    #         write-host " Erreur instruction SQL : " $Error[0]     
    #     }
        $EDS = 24VR24
        # Permet de verifier que d afficher les boutons pour le service ura et ats
        if (($EDS -eq 24VR24 ) -or ($EDS -eq 24F44)) {
            
            
            # Permet d'activer et d'afficher le bouton d'entree
            $var_btn_update.IsEnabled="True"
            $var_btn_update.Visibility="Visible"

            # Permet d'activer et d'afficher le bouton de sortie
            $var_btn_sortir.IsEnabled="True"
            $var_btn_sortir.Visibility="Visible"

            # Permet d'activer et d'afficher le bouton nouvelle
            $var_btn_nouvelle.IsEnabled="True"
            $var_btn_nouvelle.Visibility="Visible"
            $Columns = @(
                'Dossier'
                'id'
                'status'
                'Date'
                'Nom_client'
                'Date_Creation'
                'Localisation'
                'Type_Contrat'
                'Demandeur'
            )

        }else{
            $Columns = @(
                'Dossier'
                'id'
                'status'
                'date'
                'Nom_client'
                'Type_Contrat'
            )}       
    return $Columns   
}
Get-Autorisation $Columns | ForEach-Object{$var_ddlHeader.Items.Add($_)}







function GetDetails{

    
    $header = $var_ddlHeader.SelectedItem
    $Search =$var_ddlfiltre.selectedItem
    write-host
    write-host "valeur de search '$Search'"
    write-host
    $ListData | Where-Object $header -eq $Search | Select-Object -ExpandProperty "status"
    $var_lblrecherche.Content = "$header :"
    $var_lblresultid.Content = $Search
    $var_lblstatus.Content = $ListData | Where-Object $header -eq $Search | Select-Object -ExpandProperty "status"

    switch ($var_lblstatus.Content) {

        "Disponible" {$var_lblstatus.foreground = 'green'}
        "Indisponible" {$var_lblstatus.foreground = 'orange'}
        "lost" {$var_lblstatus.foreground = 'red'}
        default {$var_lblstatus.foreground = 'black'}
    }
    write-host $var_lblstatus.Text = $ListData | Select-Object status | Where-Object {$_.profession -eq $Search}

}

    
$var_ddlfiltre.Add_SelectionChanged({GetDetails})

$Columns = Get-Autorisation
$Services = $ListData | Select-Object $Columns
$ServiceDataTable = New-Object System.Data.DataTable
[void]$ServiceDataTable.Columns.AddRange($Columns)

foreach($Service in $Services){
    $Entry = @()
    foreach($column in $Columns){
        $Entry += $Service.$column
    }
    [void]$ServiceDataTable.Rows.Add($Entry)
}

$var_dg_services.ItemsSource=$ServiceDataTable.DefaultView
$var_dg_services.IsReadOnly=$true
$var_dg_services.GridLinesVisibility="None"


function Get-Datatable{
    write-host
    write-host "fonction Get-Datatable"
    $ListDataNew = Import-Csv -Path "Chemin du CSV" -Delimiter ','
    $ServicesNew = $ListDataNew | Select-Object $Columns

    $ServiceDataTableNew = New-Object System.Data.DataTable
    [void]$ServiceDataTableNew.Columns.AddRange($Columns)

    foreach ($Service in $ServicesNew) {
        $Entry = @()
        foreach ($column in $Columns) {
            $Entry += $Service.$column
        }
        [void]$ServiceDataTableNew.Rows.Add($Entry)
    }

    $var_dg_services.ItemsSource = $ServiceDataTableNew.DefaultView
    return $ServiceDataTableNew
}




# A chaque changement de selection dans le datagrid, on affiche les informations dans les labels
$var_dg_services.Add_SelectionChanged({
$var_lblresultid.content= $var_dg_services.selectedItem.date #Display the selected item in the label
$var_lblStatus.content= $var_dg_services.selectedItem.Status

    switch ($var_lblstatus.Content) {

        "Disponible" {$var_lblstatus.foreground = 'green'}
        "Indisponible" {$var_lblstatus.foreground = 'orange'}
        "lost" {$var_lblstatus.foreground = 'red'}
        default {$var_lblstatus.foreground = 'black'}
    }

})


# Bouton pour une recherche
$var_btn_search.Add_Click({
    # Get-SearhElse
    write-host " "
    $entete = $var_ddlHeader.SelectedItem
    write-host 
    $entete = "$entete"
    write-host "Valeur de entre '$entete'"
    $filter = "$entete LIKE '$($var_ddlfiltre.Text)%'"
    
    $ServiceDataTable.DefaultView.RowFilter = $filter
    
})





# Bouton pour une nouvelle entrÃ©e
# ############################################
# ############################################
# ############################################
$var_btn_update.Add_Click({
    #   Bouton pour une Entrer modification de la colonne statut : disponible/indisponible
    Write-Host "Bouton Entree (Entrer/Statut)"
    
    $xamlFile =".\entree.xaml"
    $inputXAML = Get-Content -Path $xamlFile -Raw
    $inputXAML=$inputXAML -replace 'mc:Ignorable="d"','' -replace "x:N","N" -replace '^<Win.*','<Window'
    [XML]$XAML=$inputXAML

    $reader = New-Object System.Xml.XmlNodeReader $XAML
    try{
        $psformEntrer=[Windows.Markup.XamlReader]::Load($reader)
    }catch{
        Write-Host $_.Exception
        throw
    }

    $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
        try{
            Set-Variable -Name "var_$($_.Name)" -Value $psformEntrer.FindName($_.Name) -ErrorAction Stop
        }catch{
            throw
        }
    }

    # Initialisation de de $Servicename avec la valeur de la selection dans le combobox
    $Servicename = $var_ddlfiltre.SelectedItem # $Servicename est le champ de l'element rechercher
    $var_entree_lblIdentification.content = $Servicename
    $header = $var_ddlHeader.SelectedItem

    # Verifier si la variable est vide
    if($Servicename -eq $null){
        write-host "aucune Identifacation selectionner"
        $Servicename = $var_dg_services.selectedItem.id
        $header ="id"
        $var_entree_lblIdentification.content = $Servicename
    }
    

    
    $entree_lblstatut = $ListData | Where-Object $header -eq $Servicename | Select-Object -ExpandProperty "status"
     
    $var_entree_lblcreation.Content = $ListData | Where-Object $header -eq $Servicename | Select-Object -ExpandProperty "Date_Creation"
    $var_entree_lblIdentification.Content = $ListData | Where-Object $header -eq $Servicename | Select-Object -ExpandProperty "id"
    $var_entree_lblclient.Content = $ListData | Where-Object $header -eq $Servicename | Select-Object -ExpandProperty "Nom_client"
    $var_entree_lbluser.Content = $ListData | Where-Object $header -eq $Servicename | Select-Object -ExpandProperty "Demandeur"
    $var_entree_lbldossier.Content = $ListData | Where-Object $header -eq $Servicename | Select-Object -ExpandProperty "Dossier"
    $var_entree_lblsortfinal.Content = $ListData | Where-Object $header -eq $Servicename | Select-Object -ExpandProperty "Date_de_sort_final"
    $var_entree_lblLocalisation.Content = $ListData | Where-Object $header -eq $Servicename | Select-Object -ExpandProperty "Localisation"
    $var_entree_lblsocial.Content= $ListData | Where-Object $header -eq $Servicename | Select-Object -ExpandProperty "Type_Contrat"
    $var_entree_lblDatederetour.Content = Get-Date -UFormat %d/%m/%Y
    
    if($entree_lblstatut  -eq "Disponible"){

        $var_entree_lblstatut.foreground = 'red'
        $var_entree_lblstatut.fontWeight = "Bold"
        $var_entree_lblstatut.content =" Deja  Disponible"
    }
    
    
    
    

    function sauvegarder{
        write-host "de service name $Servicename"
        $var_entree_lblIdentification.Content= $Servicename
        write-host
        write-host "Lancement de la fonction sauvegarder"
        $Date_retour = Get-Date -UFormat %d/%m/%Y # avec separateur -  %d-%m-%Y annÃ©e en sur 2 digite  %d-%m-%y
        $Demandeur = $var_sortie_textboxFullName.Text
        $Statut = "Disponible"
         
        $columnIdentification = "id"
        $searchtest = $var_entree_lblIdentification.Content
        $columnNom_Demandeur = "Demandeur"
        $columnDate_retour = "Date_de_retour"
        
        $columnStatus ="status"

        $ListData | ForEach-Object {
            if ($_."$columnIdentification" -eq $searchtest) {
                $_."$columnNom_Demandeur" = $Demandeur
                $_."$columnDate_retour" = $Date_retour
                $_."$columnStatus" = $Statut
            }
        }
        $ListData | Export-Csv -Path "Chemin du CSV" -Delimiter ',' -NoTypeInformation
        Start-Sleep -Seconds 1

    }

   
    $var_btn_apply.Add_Click({
        write-host "sauvegarder"
        sauvegarder
    })

    $var_btn_save.Add_Click({
        write-host "fermer"
        Get-Datatable
        $psformEntrer.close()
        
        
        # fermer : faire fonction actualise excel nouveaux import
    })

    $valeur = "Disponible"
    $psformEntrer.ShowDialog() 
})



# Bouton pour une nouvelle sortie
# ############################################
# ############################################
# ############################################
$var_btn_sortir.Add_Click({
    #   Bouton pour renseigner la sortie et enregister qui a recuperer le document
    Write-Host "Bouton sortie"
    $xamlFile =".\sortie.xaml"
    
    $inputXAML = Get-Content -Path $xamlFile -Raw
    $inputXAML=$inputXAML -replace 'mc:Ignorable="d"','' -replace "x:N","N" -replace '^<Win.*','<Window'
    [XML]$XAML=$inputXAML

    $reader = New-Object System.Xml.XmlNodeReader $XAML
    try{
        $psformSortie=[Windows.Markup.XamlReader]::Load($reader)
    }catch{
        Write-Host $_.Exception
        throw
    }

    $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
        try{
            Set-Variable -Name "var_$($_.Name)" -Value $psformSortie.FindName($_.Name) -ErrorAction Stop
        }catch{
            throw
        }
    }
    Get-Variable var_*
    
    # Initialisation de de $Servicename avec la valeur de la selection dans le combobox
    $Servicename = $var_ddlfiltre.SelectedItem # $Servicename est le champ de l'element rechercher
    $var_sortie_lblIdentification.content = $Servicename
    $header = $var_ddlHeader.SelectedItem
    # $Search =$var_ddlfiltre.selectedItem

    # Verifier si la variable est vide
    if($Servicename -eq $null){
        write-host "aucune Identifacation selectionner"
        $Servicename = $var_dg_services.selectedItem.id
        $header ="id"
        $var_sortie_lblIdentification.content = $Servicename
    }

    
    

    $sortie_lblstatut = $ListData | Where-Object $header -eq $Servicename | Select-Object -ExpandProperty "status"  
    $var_sortie_lblcreation.Content = $ListData | Where-Object $header -eq $Servicename | Select-Object -ExpandProperty "Date_Creation"
    $var_sortie_lblIdentification.Content = $ListData | Where-Object $header -eq $Servicename | Select-Object -ExpandProperty "id"
    $var_sortie_lblclient.Content = $ListData | Where-Object $header -eq $Servicename | Select-Object -ExpandProperty "Nom_client"
    $Demandeur = $var_sortie_tbxuser.Text
    $var_sortie_lbldossier.Content = $ListData | Where-Object $header -eq $Servicename | Select-Object -ExpandProperty "Dossier" 
    $var_sortie_lblsortfinal.Content = $ListData | Where-Object $header -eq $Servicename | Select-Object -ExpandProperty "Date_de_sort_final" 
    $var_sortie_lblLocalisation.Content = $ListData | Where-Object $header -eq $Servicename | Select-Object -ExpandProperty "Localisation"
    $var_sortie_lblsocial.Content= $ListData | Where-Object $header -eq $Servicename | Select-Object -ExpandProperty "Type_Contrat"
    $var_sortie_lblDatedesortie.Content = Get-Date -UFormat %d/%m/%Y
    
    write-host
    write-host "Valeur de statut dans sortie "$sortie_lblstatut

    if($sortie_lblstatut -eq "Indisponible"){

        $var_sortie_lblstatut.foreground = 'red'
        $var_sortie_lblstatut.fontWeight = "Bold"
        $var_sortie_lblstatut.content = "Deja  Indisponible"
    }
    

    function sauvegarder{
        write-host "de service name $Servicename"
        $var_sortie_lblIdentification.Content= $Servicename
        write-host
        write-host "Lancement de la fonction sauvegarder"
        $Date_sortie = Get-Date -UFormat %d/%m/%Y # avec separateur -  %d-%m-%Y annÃ©e en sur 2 digite  %d-%m-%y
        $Demandeur = $var_sortie_tbxuser.Text
        $Statut = "Indisponible"
            
        $columnIdentification = "id"
        $searchtest = $var_sortie_lblIdentification.Content
        $columnNom_Demandeur = "Demandeur"
        $columnDate_sortie = "Date_de_sortie"
        
        $columnStatus ="status"

        $ListData | ForEach-Object {
            if ($_."$columnIdentification" -eq $searchtest) {
                $_."$columnNom_Demandeur" = $Demandeur
                $_."$columnDate_sortie" = $Date_sortie
                $_."$columnStatus" = $Statut
            }
        }
        $ListData | Export-Csv -Path "Chemin du CSV" -Delimiter ',' -NoTypeInformation
        Start-Sleep -Seconds 1

    }

    $var_btn_apply.Add_Click({
        write-host "Bouton sauvegarder"
        sauvegarder # appel de la fonction sauvegarder
    })


    # Bouton pour fermer la fenetre
    $var_btn_save.Add_Click({
        write-host "fermer"
        Get-Datatable
        $psformSortie.close()   
    })

    $psformsortie.ShowDialog() 

})



# Bouton pour renseigner de nouvelle archive dans l'excel
# Pemet de renseigner une nouvelle archive dans l'excel ou une nouvelle synchronisation
#######################################################################################################################
$var_btn_nouvelle.Add_Click({
    #   Bouton pour renseigner de nouvelle archive dans l'excel
    Write-Host "Bouton nouvelle"

    $xamlFile =".\Nouvelle.xaml"
    $inputXAML = Get-Content -Path $xamlFile -Raw
    $inputXAML=$inputXAML -replace 'mc:Ignorable="d"','' -replace "x:N","N" -replace '^<Win.*','<Window'
    [XML]$XAML=$inputXAML

    $reader = New-Object System.Xml.XmlNodeReader $XAML
    try{
        $psformNouveau=[Windows.Markup.XamlReader]::Load($reader)
    }catch{
        Write-Host $_.Exception
        throw
    }

    $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
        try{
            Set-Variable -Name "var_$($_.Name)" -Value $psformNouveau.FindName($_.Name) -ErrorAction Stop
        }catch{
            throw
        }
    }
    
    $TypeContrat = @('Credit Bail'
        'Caution bancaire')

    $TypeContrat | ForEach-Object {$var_nouvelle_ddlsocial.Items.Add($_)}

    $ListStatut = @('Disponible'
        'Indisponible')
        
    $ListStatut | ForEach-Object {$var_nouvelle_ddlStatut.Items.Add($_)}
    #Ajout du 19/12 par Sylvio - Conversion de l'excel en CSV
    # Function ExcelToCsv ($File) {
    #     $myDir = "C:\Users\W504037\Documents\script"
    #     $excelFile = "$myDir\" + $File + ".xlsx"
    #     $objExcel = New-Object -ComObject Excel.Application
    #     $objExcel.Visible = $false
    #     $wb = $objExcel.Workbooks.Open($excelFile)
    #     foreach ($ws in $wb.Worksheets) {
    #         $ws.SaveAs("$myDir\" + $File + ".csv", 6)
    #     }
    #     $objExcel.Quit()
    # }
    # $FileName = "test"
    # ExcelToCsv -File $FileName

    # function Get-Fusion{
    #     # La fonction Get-Fusion fusionne deux fichiers CSV en un seul fichier CSV
    #     # Lisez les fichiers CSV en utilisant Import-Csv
    #     $file1 = Import-Csv -Path 'C:\file1.csv'
    #     $file2 = Import-Csv -Path 'C:\file2.csv'

    #     # Fusionnez les deux fichiers en utilisant la cmdlet + (concatÃ©nation)
    #     $fusion = $file1 + $file2

    #     # Ã‰crivez le rÃ©sultat de la fusion dans un nouveau fichier CSV
    #     $fusion | Export-Csv -Path 'C:\merged.csv' -NoTypeInformation

    # }


    function Get-Add_Archive {
        # La fonction Get-Add_Archive ajoute une nouvelle ligne dans le fichier CSV
        write-host
        write-host "Lancement de la fonction Get-Add_Archive"
        write-host
        write-host $var_nouvelle_txtb_creation.Text
        $NouvelleLinge = New-Object psobject -Property @{
            # Dictionnaire pour ajouter une nouvelle ligne dans le fichier CSV
            "Date_Creation" = if ($var_nouvelle_txtb_creation.Text) { $var_nouvelle_txtb_creation.Text } else { "N/A" }
            "id" = if ($var_nouvelle_txtb_Id.Text) { $var_nouvelle_txtb_Id.Text } else { "N/A" }
            "Nom_client" = if ($var_nouvelle_txtb_NomClient.Text) { $var_nouvelle_txtb_NomClient.Text } else { "N/A" }
            "status" = if ($var_nouvelle_ddlStatut.SelectedItem) { $var_nouvelle_ddlStatut.SelectedItem } else { "N/A" }
            "Demandeur" = if ($var_nouvelle_txtb_Utilisateur.Text) { $var_nouvelle_txtb_Utilisateur.Text } else { "N/A" }
            "Dossier" = if ($var_nouvelle_txtb_dossier.Text) { $var_nouvelle_txtb_dossier.Text } else { "N/A" }
            "Date_de_sort_final" = if ($var_nouvelle_txtb_DateDestruction.Text) { $var_nouvelle_txtb_DateDestruction.Text } else { "N/A" }
            "Localisation" = if ($var_nouvelle_txtb_Localisation.Text) { $var_nouvelle_txtb_Localisation.Text } else { "N/A" }
            "Type_Contrat" = if ($var_nouvelle_ddlsocial.SelectedItem) { $var_nouvelle_ddlsocial.SelectedItem } else { "N/A" }
            "Date_de_sortie" = if ($var_nouvelle_txtb_DateSortie.Text) { $var_nouvelle_txtb_DateSortie.Text } else { "N/A" }
            "date" = Get-Date -UFormat %d/%m/%Y
            
        }
        # Ajoute la nouvelle ligne dans le fichier CSV
        
        $data = Import-Csv -Path "Chemin du CSV" -Delimiter ',' 
        $data += $NouvelleLinge
        $data | Export-Csv -Path "Chemin du CSV"  -NoTypeInformation
        write-host
        write-host "Fin de la fonction Get-Add_Archive"
        
    }



    $var_btn_apply.Add_Click({
        write-host "sauvegarder"
        # Bouton pour enregistrer les deux fichiers CSV
        Get-Add_Archive
        
    })

    $var_btn_save.Add_Click({
        write-host "fermer"
        Get-Datatable
        $psformNouveau.close()
        # fermer : faire fonction actualise excel nouveaux import
    })
        
    $psformNouveau.ShowDialog() 
})

$psform.ShowDialog()
