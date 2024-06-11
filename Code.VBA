
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

Here is the code for the Function to find the first empty line in the sheet "Consummable":


Sub TrouverPremiereLigneVideConsommable()
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

    
  

End Sub

Here is the code for the "UserForm_Example" (consummables):


Private Sub ComboBox1_Change()
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
End Sub

Here is the code for the UserForm_Imp : 
Private Sub ComboBoxMarqueImp_Change()
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
End Sub
