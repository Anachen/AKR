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
MsgBox "V�laszd ki a lakc�mhez tartoz� orsz�got �s add meg a telep�l�s nev�t!" _
, vbInformation + vbOKOnly, "Orsz�g, Telep�l�s adatok"
Me.CIM_ORSZKOD.SetFocus

End Sub
'a magyar �s k�lf�ldi c�mek bevitele
Private Sub CIM_ORSZKOD_Click()
If Me.CIM_ORSZKOD = "HU" Then
Me.tel.Visible = True
Me.CIM_TELNEV.Visible = True
Else
Me.CIM_KHELYSEG.Visible = True
Me.ktel.Visible = True

End If

End Sub

'ir�ny�t�sz�m meghat�roz�sa a telep�l�s nev�nek megad�s�val
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
MsgBox "T�bb lehets�ges ir�ny�t�sz�m.(" & irtol & "-" & irig & " ) K�rem ellen�rizze!", _
vbOKOnly + vbInformation, "Figyelmeztet�s"
Me.CIM_IRSZAM = Me.CimData.Column(3)
Me.CIM_IRSZAM.SetFocus

Else
Me.CIM_IRSZAM = Me.CimData.Column(6)
MsgBox "Az adatb�zis be�ll�totta az ir�ny�t�sz�mot (" & Me.CIM_IRSZAM & ")", _
vbOKOnly + vbInformation, "Inform�ci�"
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
MsgBox "Valamilyen k�telez� adatot nem adt�l meg. Ellen�rizd az adatokat �s pr�b�ld �jra" _
, vbCritical + vbOKOnly, "Hi�nyz� adat!"
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
MsgBox "Valamilyen k�telez� adatot nem adt�l meg. Ellen�rizd az adatokat �s pr�b�ld �jra" _
, vbCritical + vbOKOnly, "Hi�nyz� adat!"
Exit Sub
Else
Forms![Main].[AddrDone].Visible = True
Forms![Main].[addressDone].Visible = True
Forms![Main].[addressDone].Value = True
DoCmd.Close
End If
End Sub


