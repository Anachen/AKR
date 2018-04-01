VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_staffData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
' betöltésnél az aktuális adóazonosítót kitölti, a rögzítõ személyét is megszerzi
If Forms![Main].[hasMain].Value = False Then
Me.Rogzito = TempVars!selectedUser
Me.SZEM_ADOAZON = TempVars!adoazon
Me.SZEM_TAJSZAM = TempVars!tajszam
Me.SZEM_SZULDATUM = TempVars!szuldatum
Me.Rogzito = TempVars!selectedUser
Else
Me.Rogzito = TempVars!selectedUser
End If

End Sub


Private Sub SZEM_TAJSZAM_AfterUpdate()
Dim rs As Recordset
Set rs = CurrentDb.OpenRecordset("Szemely", dbOpenSnapshot)
rs.FindFirst "SZEM_TAJSZAM = """ & Me![SZEM_TAJSZAM] & """"
If Not rs.NoMatch Then
     MsgBox "A Beírt tajszám már létezik, írd be újra!", vbCritical + vbOKOnly, "Duplikált adat"
     Me.SZEM_TAJSZAM.Value = ""
       Cancel = True
        
End If
End Sub

Private Sub Save_Click()
If Len(Me.SZEMV_ALLAMPOLG) > 2 Then
MsgBox "Az adat nem menthetõ, mert nem megfelelõ állampolgársági adatot adtál meg. Válassz a listából!", vbCritical _
+ vbOKOnly, "Érvénytelen állampolgárság adat"
Me.SZEMV_ALLAMPOLG.SetFocus
Exit Sub
ElseIf Me.SZEM_TAJSZAM = "" Or Me.SZEM_CSNEV = "" _
Or IsNull(Me.SZEM_UNEV) = True Or Me.SZEM_SZCSNEV = "" _
Or Me.SZEM_SZULNEV = "" Or Me.SZEM_ACSNEV = "" _
Or Me.SZEM_AUNEV = "" Or Me.SZEM_SZULHELY = "" _
Or Me.SZEM_SZULDATUM = "" Or Me.SZEMV_ALLAMPOLG = "" Then
MsgBox "Nem töltöttél ki minden kötelezõ mezõt! Ellenõrizd a kitöltött adatokat és próbáld újra!", vbCritical _
+ vbOKOnly, "Hiányzó adatok"
Exit Sub
Else

MsgBox "A munkavállaló az alábbi adatokkal lett mentve: " & vbNewLine & _
Me.SZEM_ADOAZON & vbNewLine & Me.SZEM_TAJSZAM & vbNewLine & Me.SZEM_ELONEV & " " & _
Me.SZEM_CSNEV & vbNewLine & Me.SZEM_UNEV & " " & Me.SZEM_UNEV2 & vbNewLine & _
Me.SZEM_SZCSNEV & vbNewLine & Me.SZEM_SZULNEV & " " & Me.SZEM_SZULNEV2 & vbNewLine & Me.SZEM_ACSNEV & vbNewLine & _
Me.SZEM_AUNEV & " " & Me.SZEM_AUNEV2 & vbNewLine & Me.SZEM_SZULHELY & vbNewLine & Me.SZEM_SZULDATUM & vbNewLine & _
Me.SZEMV_ALLAMPOLG, vbInformation + vbOKOnly, "Felrögzített adatok"
Forms![Main].[hasMain].Value = True
Forms![Main].[newWorker].Visible = False
Forms![Main].[Address].Visible = True
Forms![Main].[Address].SetFocus
Forms![Main].[newWorker].Visible = False
Forms![Main].Form.[Szemely].Visible = False
Forms![Main].[currentStaff].Visible = True
Forms![Main].[currentAdo].Visible = True
Forms![Main].[currentTaj].Visible = True
Forms![Main].[currentCsal].Visible = True
Forms![Main].[currentUt1].Visible = True
Forms![Main].[currentUt2].Visible = True
Forms![Main].[currenthely].Visible = True
Forms![Main].[currentIdo].Visible = True
Forms![Main].[Foglalkoz].Visible = True
Forms![Main].[vegzett].Visible = True
Forms![Main].[vezeto].Visible = True
Forms![Main].[nyelv].Visible = True
Forms![Main].[Printing].Visible = True
Forms![Main].[Reset].Visible = True
Forms![Main].[modifyMain].Visible = True
Forms![Main].[currentAdo].Value = Me.SZEM_ADOAZON
Forms![Main].[currentTaj].Value = Me.SZEM_TAJSZAM
Forms![Main].[currentCsal].Value = Me.SZEM_CSNEV
Forms![Main].[currentUt1].Value = Me.SZEM_UNEV
Forms![Main].[currentUt2].Value = Me.SZEM_UNEV2
Forms![Main].[currenthely].Value = Me.SZEM_SZULHELY
Forms![Main].[currentIdo].Value = Me.SZEM_SZULDATUM
TempVars!tajszam = Null
TempVars!szuldatum = Null

DoCmd.Close
End If

End Sub


Private Sub SZEMV_ALLAMPOLG_AfterUpdate()
If Len(Me.SZEMV_ALLAMPOLG) > 2 Then
Me.SZEMV_ALLAMPOLG.SetFocus
ElseIf Me.SZEMV_ALLAMPOLG = "hu" Or Me.SZEMV_ALLAMPOLG = "Hu" Then
Me.SZEMV_ALLAMPOLG = "HU"
End If


End Sub

Private Sub SZEMV_KFBIZT_Click()
If Me.SZEMV_KFBIZT = "I" Then
MsgBox "Külföldön biztosított esetén adatokat kérem megadni!" _
, vbOKOnly + vbInformation, "További adatok szükségesek!"
Me.szarmorsz.Visible = True
Me.SZEMV_KFBIZTORSZ.Visible = True
Me.kezd.Visible = True
Me.veg.Visible = True
Me.SZEMV_KFBIZTKEZD.Visible = True
Me.SZEMV_KFBIZTVEGE.Visible = True
End If


End Sub

Private Sub SZEMV_KFILL_Click()
If Me.SZEMV_KFILL = "I" Then
MsgBox "Külföldi illetõség esetén adatokat kérem megadni!", _
vbOKOnly + vbInformation, "További adatok szükségesek!"
Me.okmanysz.Visible = True
Me.SZEMV_KFOKSZAM.Visible = True
Me.orszag.Visible = True
Me.SZEMV_KFALLAM.Visible = True
End If



End Sub


