VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Vegzettseg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
If Forms![Main].[educationDone].Value = True Then
Me.Rogzito = TempVars!selectedUser
Else
Me.Kep_AD_AZON = TempVars!adoazon
Me.Rogzito = TempVars!selectedUser
End If
End Sub

Private Sub ISK_KVKOD_Click()
If Me.ISK_KVKOD = "31" Or Me.ISK_KVKOD = "54" Or Me.ISK_KVKOD = "61" Then
MsgBox "Egy�b eset�n, k�rem meghat�rozni", _
vbOKOnly + vbExclamation, "Figyelmeztet�s"
Me.ISK_KVEGYEB.Visible = True
Me.egyeb.Visible = True
Me.ISK_KVEGYEB.SetFocus
End If


End Sub

Private Sub Save_Click()
Dim myDate
myDate = DateValue("12/31/3999")
If Me.ISK_KVJELZO = "" Or IsNull(Me.ISK_KVKOD) = True _
         Or IsNull(Me.ISK_INTEZMENY) = True Or IsNull(Me.ISK_SZAK) = True _
         Or IsNull(Me.ISK_MDAT) = True Or IsNull(Me.ISK_OKLEVEL) = True Then
         MsgBox "Valami nem stimmel! K�rlek ellen�rzid, mindent kit�lt�tt�l-e", vbCritical + vbOKOnly, "Hi�nyz� adat"
         Exit Sub
         ElseIf myDate <> DLookup("[Valid]", "[Vegzettseg]", _
         "[Kep_AD_AZON]='" & TempVars!adoazon & "' and [Expired] = False") Then
     MsgBox "Egy v�gzetts�g r�gz�t�se k�telez�!", vbOKOnly + vbCritical, "Figyelmeztet�s"
         Exit Sub
         
         Else
                     
         
         
Forms![Main].[Vegz].Visible = True
Forms![Main].[educationDone].Visible = True
Forms![Main].[educationDone].Value = True
Forms![Main].[Printing].SetFocus
DoCmd.Close acForm, "Vegzettseg", acSaveYes
End If


End Sub
Private Sub Expired_Click()
Me.Valid = Now
Me.Exported.Value = True
Me.Printed.Value = True
MsgBox "A v�gzetts�get �rv�nytelen�tett�k!", _
vbOKOnly + vbInformation, "Inform�ci�"
DoCmd.Close acForm, "Vegzettseg", acSaveYes
Forms![Main].[vegzett].SetFocus
Me.Expired.Visible = False
End Sub



