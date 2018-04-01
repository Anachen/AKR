VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_AddressData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CIM_CIMTIP_AfterUpdate()
MsgBox "Válaszd ki a lakcímhez tartozó országot és add meg a település nevét!" _
, vbInformation + vbOKOnly, "Ország, Település adatok"
Me.CIM_ORSZKOD.SetFocus

End Sub
'a magyar és külföldi címek bevitele
Private Sub CIM_ORSZKOD_Click()
If Me.CIM_ORSZKOD = "HU" Then
Me.tel.Visible = True
Me.CIM_TELNEV.Visible = True
Else
Me.CIM_KHELYSEG.Visible = True
Me.ktel.Visible = True

End If

End Sub

'irányítószám meghatározása a település nevének megadásával
Private Sub CIM_TELNEV_Click()
Dim irtol As Integer
Dim irig As Integer
Dim finalIr As Integer


Me.CimData = Me.CIM_TELNEV
Me.CIM_TEL_AZON = Me.CimData.Column(0)
Me.CIM_KSH = Me.CimData.Column(2)
irtol = DLookup("[IRSZTOL]", "[TELEPULES]", "[Azon]='" & Me.CIM_TEL_AZON & "'")
irig = DLookup("[IRSZIG]", "[TELEPULES]", "[Azon]='" & Me.CIM_TEL_AZON & "'")
finalIr = (irig - irtol)

If finalIr > 1 Then
MsgBox "Több lehetséges irányítószám.(" & irtol & "-" & irig & " ) Kérem ellenõrizze!", _
vbOKOnly + vbInformation, "Figyelmeztetés"
Me.CIM_IRSZAM = Me.CimData.Column(3)
Me.CIM_IRSZAM.SetFocus

Else
Me.CIM_IRSZAM = Me.CimData.Column(6)
MsgBox "Az adatbázis beállította az irányítószámot (" & Me.CIM_IRSZAM & ")", _
vbOKOnly + vbInformation, "Információ"
Me.CIM_KOZNEV.SetFocus

End If


End Sub

Private Sub Form_Load()
Me.AD_AZON_CIM = TempVars!adoazon
Me.Rogzito = TempVars!selectedUser
Me.CIM_CIMTIP.SetFocus

End Sub


Private Sub Renew_Click()
If IsNull(Me.CIM_CIMTIP) = True Or Me.CIM_ORSZKOD = "" Or Me.CIM_TELNEV = "" Or _
Me.CIM_IRSZAM = "" Or Me.CIM_KOZNEV = "" Or IsNull(Me.CIM_KOZJEL) Or IsNull(Me.CIM_HAZSZAM) _
Or Me.CIM_HAZSZAM = "" Then
MsgBox "Valamilyen kötelezõ adatot nem adtál meg. Ellenõrizd az adatokat és próbáld újra" _
, vbCritical + vbOKOnly, "Hiányzó adat!"
Exit Sub
Else
Forms![Main].[AddrDone].Visible = True
Forms![Main].[addressDone].Visible = True
Forms![Main].[addressDone].Value = True
DoCmd.Close
End If
End Sub

Private Sub Submit_Click()
If IsNull(Me.CIM_CIMTIP) = True Or Me.CIM_ORSZKOD = "" Or Me.CIM_TELNEV = "" Or _
Me.CIM_IRSZAM = "" Or Me.CIM_KOZNEV = "" Or IsNull(Me.CIM_KOZJEL) Or IsNull(Me.CIM_HAZSZAM) _
Or Me.CIM_HAZSZAM = "" Then
MsgBox "Valamilyen kötelezõ adatot nem adtál meg. Ellenõrizd az adatokat és próbáld újra" _
, vbCritical + vbOKOnly, "Hiányzó adat!"
Exit Sub
Else
Forms![Main].[AddrDone].Visible = True
Forms![Main].[addressDone].Visible = True
Forms![Main].[addressDone].Value = True
DoCmd.Close
End If
End Sub


