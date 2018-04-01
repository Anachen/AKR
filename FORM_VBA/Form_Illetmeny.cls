VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Illetmeny"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub change_JC_Click()
Dim adminjog
adminjog = InputBox("Kérem a jelszót!", "Admin joghoz kötött felület")
If strPasswd = "" Or strPasswd = Empty Then
        MsgBox "Jelszó szükséges!", vbInformation, "Hiányzó adat!"
        Exit Sub
    End If

     If strPasswd = "OHinterface2017" Then
    MsgBox "Admin vagy!", _
               vbOKOnly, "Admin"
   MsgBox "Illetmény Típusának módosítása után az adat nem lesz importálható!", _
vbOKOnly + vbCritical, "Figyelmeztetés"
Me.ChangeIll.Visible = True
    Else
      MsgBox "Nincs jogosultságod", _
               vbOKOnly, "Hiányzó jogosultság"
Exit Sub
End If

End Sub

Private Sub ChangeIll_Click()
If Me.ChangeIll <> "11294" Then
MsgBox "Illetmény Típusának módosítása után az adat nem lesz importálható!", _
vbOKOnly + vbCritical, "Figyelmeztetés"

Me.FLT_JOGCIM = Me.ChangeIll.Column(0)
Me.FLT_NEV = Me.ChangeIll.Column(1)

Else
Me.FLT_JOGCIM = Me.ChangeIll.Column(0)
Me.FLT_NEV = Me.ChangeIll.Column(1)
End If

End Sub

Private Sub Expired_Click()
Me.Valid = Now
Me.Exported.Value = True
Me.Printed.Value = True
MsgBox "Az illetmény lezárásra került. Az új felvételéhez kattintson újra az Illetményadatok gombra!", _
vbOKOnly + vbInformation, "Információ"
DoCmd.Close acForm, "Illetmeny", acSaveYes
Forms![Foglalkoztatas].[Illetm].SetFocus
Forms![Foglalkoztatas].[IlletmenyDone].Value = False

End Sub

Private Sub Form_Load()
If Forms![Foglalkoztatas].[IlletmenyDone].Value = True Then
Me.Rogzito = TempVars!selectedUser
Me.FLT_KEZDETE = Forms![Foglalkoztatas].[FG_JOGVKEZDETE].Value
Else
Me.Ill_AD_AZON = TempVars!adoazon
Me.FLT_KEZDETE = Forms![Foglalkoztatas].[FG_JOGVKEZDETE].Value
Me.Rogzito = TempVars!selectedUser
End If

End Sub

Private Sub Save_Click()
If Me.FLT_ERTEKE = 0 Then
MsgBox "Illetményadat megadása kötelezõ!", vbCritical + vbOKOnly, "Hiányzó adat"

Exit Sub
Else

DoCmd.Close acForm, "Illetmeny", acSaveYes
Forms![Foglalkoztatas].[FV_JELMINTA].SetFocus
Forms![Foglalkoztatas].[IlletmenyDone].Value = True

End If

End Sub


