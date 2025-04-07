Sub ProtectSheets()
    Dim ws As Worksheet
    
    ' Protéger toutes les feuilles du classeur
    For Each ws In ThisWorkbook.Sheets
        ws.Protect Password:="motdepasse" ' Remplacez par le mot de passe souhaité
    Next ws
    
    ' Déprotéger une feuille spécifique pour modification
    ThisWorkbook.Sheets("Sheet1").Unprotect Password:="motdepasse"
    
    MsgBox "Les feuilles ont été protégées !"
End Sub
