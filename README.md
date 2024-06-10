# Contents
[1-Project Overview](#1-project-overview)

[2-Making of the program](#2-making-of-the-program)

[3-Description of the Program](#3-description-of-the-program)

[4-Raw Code](#4-raw-code)

# 1-Project Overview
This project has for objective to allow users to add printers to a list with the consummables, the main objective behind the project is the use of the coding language VBA.

There is a few tests described in [# 3 Description of the Program](#3-description-of-the-program).

The code is available inside the Spreadsheet.

# 2-Making of the program
The program is using simple VBA trough EXCEL, the program is entirely made by hand and might need some adjustements, it's not perfect in any way but it works and is a proof of concept.

Before making the program we need to first design the User Form that will be completed:
Here is the first draw of the forms:

![Modèle de formulaire d'ajout d'imprimante](https://github.com/C-Brq/VosReves/assets/156824818/3b47a3f6-fc1f-4985-ac2b-e6933b76866f)

There is 2 main user forms, the first one includes consummables and their caracteristics such as Reference, Type of Printer, Brand, Price, Color and others...
Here is a quick view of the form:

<img width="218" alt="image" src="https://github.com/C-Brq/VosReves/assets/156824818/98e6e3d8-6ce7-4796-82db-dfe9db7b0d5e">

The second userform is purely about the Printer, it includes caracteristics such as the Brand,Reference , Price , Type of Printer and the references for the colors.
Here is a quick view of the form:

<img width="341" alt="image" src="https://github.com/C-Brq/VosReves/assets/156824818/97d13ccf-4fca-4c96-bcbf-e938332dbc5b">

# 3-Description of the Program
The Program is made to limit the user error to the maximum by preserving the errors from the inputs.

The program checks the type of printer and such parameters like the different references, The Spreadsheet has 3 pages:

The first spreadsheet is the printers

The second spreadsheet is the consummables and such.

And the third spreadsheet is the brands that are accepted and the ID linked to the brand and also includes the brand of the consummables, in addition the type of printer is error checked, by using data : a laser printer can't be linked with a ink consummable.

On the second spreadsheet there is an autoconcatenation for the references, that way the user doesn't have to type the brand of the consummable and just has to type the new reference to add, this also allows no errors linked to similar references while being different brands.

# 4-Raw Code
Here is the raw code used in the .XLS:


Here is the code for the Function to find the first empty line in the sheet "Printers":

<code>
Sub TrouverPremiereLigneVideImprimante()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(Imprimante) ' Remplacez "Nom de la feuille" par le nom de votre feuille
    Dim col As Range
    Set col = ws.Columns("B") ' Remplacez "A" par la lettre de la colonne que vous voulez vérifier
    
    Dim i As Long
    i = 1
    Do While Not IsEmpty(col.Cells(i, 1).Value)
        i = i + 1
    Loop
    MsgBox "La première ligne vide dans la colonne A est la ligne " & i
End Sub
</code>

Here is the code for the Function to find the first empty line in the sheet "Consummable":

<code>Sub TrouverPremiereLigneVideConsommable()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(Consommable) ' Remplacez "Nom de la feuille" par le nom de votre feuille
    Dim col As Range
    Set col = ws.Columns("B") ' Remplacez "A" par la lettre de la colonne que vous voulez vérifier
    
    Dim i As Long
    i = 1
    Do While Not IsEmpty(col.Cells(i, 1).Value)
        i = i + 1
    Loop
    
    MsgBox "La première ligne vide dans la colonne A est la ligne " & i
End Sub</code>

Here is the code for the "UserForm_Example" (consummables):

<code>Private Sub ComboBox1_Change()
End Sub
Private Sub ComboBox2_Change()
End Sub
Private Sub CommandButton1_Click()
    Dim ligne As Long
    Dim colonne As Long
    Dim ws As Worksheet
    colonne = 1
    'feuille de calcul
    Set ws = ThisWorkbook.Worksheets("Consommable")
    PremiereLigne = 12
     Do While ws.Cells(PremiereLigne, 3).Value <> ""
        PremiereLigne = PremiereLigne + 1
    Loop
    
    'Association entre le Choixype et la combobox
    Dim Choixtype As String
    Choixtype = ComboBox2.Value
    ws.Cells(PremiereLigne, colonne + 1).Value = Choixtype
    
    'Association entre le Choixconso et la combobox
    Dim Choixconso As String
    Choixconso = ComboBox5.Value
 
 'Limitateur de Type et consommable
 If Choixtype = "LASER" And Choixconso = "encre" Then
    MsgBox "Erreur le type d'imprimante et le Choix de consommables ne sont pas compatibles.", vbExclamation
    Exit Sub
    End If
    
 If Choixtype = "JET D'ENCRE" And Choixconso = "toner" Then
    MsgBox "Erreur le type d'imprimante et le Choix de consommables ne sont pas compatibles.", vbExclamation
    Exit Sub
    End If
    
If Choixtype = "JET D'ENCRE" And Choixconso = "tambour" Then
    MsgBox "Erreur le type d'imprimante et le Choix de consommables ne sont pas compatibles.", vbExclamation
    Exit Sub
    End If
       
       'Association entre les et la combobox
ws.Cells(PremiereLigne, colonne + 2).Value = Choixconso
Dim Choixmarque As String
Choixmarque = ComboBox3.Value
ws.Cells(PremiereLigne, colonne + 3).Value = Choixmarque

'Concaténation pour la marque et la référence
If InStr(1, Choixmarque, "BROTHER") > 0 Then
    Valeur_Référence2 = "BROTHER"
ElseIf InStr(1, Choixmarque, "CANON") > 0 Then
    Valeur_Référence2 = "CANON"
ElseIf InStr(1, Choixmarque, "EPSON") > 0 Then
    Valeur_Référence2 = "EPSON"
ElseIf InStr(1, Choixmarque, "FUDJI") > 0 Then
    Valeur_Référence2 = "FUDJI"
ElseIf InStr(1, Choixmarque, "HP") > 0 Then
    Valeur_Référence2 = "HP"
ElseIf InStr(1, Choixmarque, "KODAK") > 0 Then
    Valeur_Référence2 = "KODAK"
ElseIf InStr(1, Choixmarque, "SAMSUNG") > 0 Then
    Valeur_Référence2 = "SAMSUNG"
ElseIf InStr(1, Choixmarque, "XEROX") > 0 Then
    Valeur_Référence2 = "XEROX"
ElseIf InStr(1, Choixmarque, "RICOH") > 0 Then
    Valeur_Référence2 = "RICOH"
End If
     
Dim Valeur_Référence As String
    Valeur_Référence = TextBox1.Value
Dim Valeur_RéférenceF As String
    Valeur_RéférenceF = Valeur_Référence2 + " " + Valeur_Référence
    'insertion de la référence
ws.Cells(PremiereLigne, colonne).Value = Valeur_RéférenceF
      
      'Association entre les données et combobox
    Dim Prix As String
    Prix = TextBox4.Value
    ws.Cells(PremiereLigne, colonne + 4).Value = Prix
       
    Dim rendement As String
    rendement = TextBox5.Value
    ws.Cells(PremiereLigne, colonne + 5).Value = rendement
    
    Dim Choixcouleur As String
    Choixcouleur = ComboBox4.Value
    ws.Cells(PremiereLigne, colonne + 6).Value = Choixcouleur
    
    'Fermeture
    Unload Me
End Sub
Private Sub ListBox1_Click()
End Sub
Private Sub Label3_Click()
End Sub
Private Sub TextBox1_Change()
End Sub
Private Sub TextBox2_Change()
End Sub
Private Sub UserForm_Initialize()
    Dim cell As Range
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("DATA")
    
    
    'type imprimante
    ' Loop through the cells
    For Each cell In ws.Range("I4:I" & ws.Cells(ws.Rows.Count, "I").End(xlUp).Row)
        ' Add each cell value
        ComboBox2.AddItem cell.Value
    Next cell
    
    Dim cell2 As Range
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("DATA")
    
    
    'Marque
    ' Loop through the cells
    For Each cell2 In ws.Range("G4:G" & ws.Cells(ws.Rows.Count, "G").End(xlUp).Row)
        ' Add each cell value
        ComboBox3.AddItem cell2.Value
    Next cell2
    
    ComboBox4.AddItem "noir"
    ComboBox4.AddItem "cyan"
    ComboBox4.AddItem "jaune"
    ComboBox4.AddItem "magenta"
    
    Dim cell3 As Range
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("DATA")
    
    'type consommable
    ' Loop through the cells
    For Each cell3 In ws.Range("E4:E" & ws.Cells(ws.Rows.Count, "E").End(xlUp).Row)
        ' Add each cell value
        ComboBox5.AddItem cell3.Value
    Next cell3
    
    
    
    
    UserForm_Exemple.Height = 296.25
    UserForm_Exemple.Width = 227.75
End Sub
Sub lancerUserform()
    UserForm_Exemple.Show
End Sub</code>

Here is the code for the UserForm_Imp :

<code> Private Sub ComboBoxMarqueImp_Change()

End Sub

Private Sub CommandButton1_Click()
    Dim PremiereLigne As Long
    Dim colonne As Long
    Dim ws As Worksheet
    Dim Choixmarque As String
    Dim Référence As String
    Dim Prix As String
    Dim Choixtype As String
    
    
    colonne = 1
    ' feuille de calcul utilisée
    Set ws = ThisWorkbook.Worksheets("Imprimante")
    PremiereLigne = 4
     Do While ws.Cells(PremiereLigne, 3).Value <> ""
        PremiereLigne = PremiereLigne + 1
    Loop
    
    'Association entre les combobox et les données
    Choixmarque = ComboBoxMarqueImp.Value
    ws.Cells(PremiereLigne, colonne).Value = Choixmarque
    
    
    Référence = TextBox1.Value
    ws.Cells(PremiereLigne, colonne + 1).Value = Référence
    
    
    
    
    Prix = TextBox2.Value
    ws.Cells(PremiereLigne, colonne + 2).Value = Prix
    
    Choixtype = ComboBoxTypeImpImp.Value
    ws.Cells(PremiereLigne, colonne + 3).Value = Choixtype
    
       
         Ref_Noir = ComboBoxNoir.Value
    ws.Cells(PremiereLigne, colonne + 4).Value = Ref_Noir
    
        Ref_Cyan = ComboBoxCyan.Value
    ws.Cells(PremiereLigne, colonne + 5).Value = Ref_Cyan
    
        Ref_Jaune = ComboBoxJaune.Value
    ws.Cells(PremiereLigne, colonne + 6).Value = Ref_Jaune
    
        Ref_Magenta = ComboBoxMagenta.Value
    ws.Cells(PremiereLigne, colonne + 7).Value = Ref_Magenta
    
    
    'Fermeture
    Unload Me
End Sub
Private Sub Label1_Click()
End Sub
Private Sub Label3_Click()
End Sub
Private Sub TextBox2_Change()
End Sub
Sub AddItemsWithIncrement()
    Dim comboBox As OLEObject
    Dim i As Integer
    Dim increment As Integer

    ' Set the ComboBox
    Set comboBox = Worksheets("UserForm_Imp").OLEObjects("ComboBoxMarqueImp")
    
    
    'increment value
    increment = 1
    
    ' Add items  with an increment
    For i = 1 To 100 Step increment
        comboBox.Object.AddItem "Item " & i
    Next i
End Sub
Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim cell As Range
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("DATA")
    
    ' Loop through the cells
    For Each cell In ws.Range("B4:B" & ws.Cells(ws.Rows.Count, "B").End(xlUp).Row)
        ComboBoxMarqueImp.AddItem cell.Value
    Next cell

    
    Dim cell2 As Range
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("DATA")
    
    ' Loop through the cells
    For Each cell2 In ws.Range("I4:I" & ws.Cells(ws.Rows.Count, "I").End(xlUp).Row)
        ComboBoxTypeImpImp.AddItem cell2.Value
    Next cell2
    
    
    Dim cell3 As Range
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Consommable")
    
    ' Loop through the cells
    For Each cell3 In ws.Range("A5:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
    If InStr(1, ws.Cells(cell3.Row, "G").Value, "noir", vbTextCompare) > 0 Then
        ComboBoxNoir.AddItem cell3.Value
    End If
    Next cell3


    Dim cell4 As Range
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Consommable")
    
    ' Loop through the cells
    For Each cell4 In ws.Range("A5:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
    If InStr(1, ws.Cells(cell4.Row, "G").Value, "jaune", vbTextCompare) > 0 Then
        ComboBoxJaune.AddItem cell4.Value
    End If
    Next cell4

    Dim cell5 As Range
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Consommable")
    
    ' Loop through the cells
    For Each cell5 In ws.Range("A5:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
    If InStr(1, ws.Cells(cell5.Row, "G").Value, "cyan", vbTextCompare) > 0 Then
        ComboBoxCyan.AddItem cell5.Value
    End If
    Next cell5

    



   Dim cell6 As Range
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Consommable")
    
    ' Loop through the cells
    For Each cell6 In ws.Range("A5:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
    If InStr(1, ws.Cells(cell6.Row, "G").Value, "magenta", vbTextCompare) > 0 Then
        ComboBoxMagenta.AddItem cell6.Value
    End If
    Next cell6
    
    
    UserForm_Exemple.Height = 296.25
    UserForm_Exemple.Width = 227.75
End Sub
Sub lancerUserform()
    UserForm_Imp.Show
End Sub
Private Sub SaveBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
CommandButton1.BackColor = &HFFFFFF
CommandButton1.ForeColor = &H8000000D
End Sub
 
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
CommandButton1.BackColor = &HFFFFFF
CommandButton1.ForeColor = &H8000000D
End Sub</code>
