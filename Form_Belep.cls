VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Belep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Enter_Click()
' Felhasználó tárolása TempVars-ba
Dim selectedUser As TempVars

TempVars!selectedUser = Me.userName.Value

' Felhasználónév és vagy jelszó nem maradhat üresen, ennek ellenõrzése
   If IsNull(Me.userName) Or IsNull(Me.pass) Then
      MsgBox "Jelszót és felhasználónevet be kell írnod!", vbOKOnly + vbInformation, "Hiányzó Adatok!"
        Me.userName.SetFocus
        Exit Sub
    End If
    ' User validálása, beléptetése a programba
    If Me.pass.Value = DLookup("[password]", "USER", "[userId]='" & Me.userName.Value & "'") Then
                    
            
            DoCmd.Close acForm, "Belep", acSaveNo
            DoCmd.OpenForm "Main"
            
        
        
    Else
    ' rossz jelszó esetén hibaüzenet
    
      MsgBox "A felhasználónév és/vagy a jelszó nem egyezik!", vbOKOnly + vbExclamation, "Hibás adatok!"
      Me.pass.SetFocus
    End If

   
    

   
End Sub

Private Sub Exit_Click()
DoCmd.Quit

End Sub

Private Sub userName_AfterUpdate()
' minden belépést logol a program a Log-Tablebe ehhez adatok megadása

CurrentDb.Execute "INSERT INTO Log ( [UserName], [timeStamp] )" & _
 "VALUES('" & Me.userName & "','" & DateTime.Now & "');"


End Sub


