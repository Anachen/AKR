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
' Felhaszn�l� t�rol�sa TempVars-ba
Dim selectedUser As TempVars

TempVars!selectedUser = Me.userName.Value

' Felhaszn�l�n�v �s vagy jelsz� nem maradhat �resen, ennek ellen�rz�se
   If IsNull(Me.userName) Or IsNull(Me.pass) Then
      MsgBox "Jelsz�t �s felhaszn�l�nevet be kell �rnod!", vbOKOnly + vbInformation, "Hi�nyz� Adatok!"
        Me.userName.SetFocus
        Exit Sub
    End If
    ' User valid�l�sa, bel�ptet�se a programba
    If Me.pass.Value = DLookup("[password]", "USER", "[userId]='" & Me.userName.Value & "'") Then
                    
            
            DoCmd.Close acForm, "Belep", acSaveNo
            DoCmd.OpenForm "Main"
            
        
        
    Else
    ' rossz jelsz� eset�n hiba�zenet
    
      MsgBox "A felhaszn�l�n�v �s/vagy a jelsz� nem egyezik!", vbOKOnly + vbExclamation, "Hib�s adatok!"
      Me.pass.SetFocus
    End If

   
    

   
End Sub

Private Sub Exit_Click()
DoCmd.Quit

End Sub

Private Sub userName_AfterUpdate()
' minden bel�p�st logol a program a Log-Tablebe ehhez adatok megad�sa

CurrentDb.Execute "INSERT INTO Log ( [UserName], [timeStamp] )" & _
 "VALUES('" & Me.userName & "','" & DateTime.Now & "');"


End Sub


