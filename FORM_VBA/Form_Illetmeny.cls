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
adminjog = InputBox("K�rem a jelsz�t!", "Admin joghoz k�t�tt fel�let")
If strPasswd = "" Or strPasswd = Empty Then
        MsgBox "Jelsz� sz�ks�ges!", vbInformation, "Hi�nyz� adat!"
        Exit Sub
    End If

     If strPasswd = "OHinterface2017" Then
    MsgBox "Admin vagy!", _
               vbOKOnly, "Admin"
   MsgBox "Illetm�ny T�pus�nak m�dos�t�sa ut�n az adat nem lesz import�lhat�!", _
vbOKOnly + vbCritical, "Figyelmeztet�s"
Me.ChangeIll.Visible = True
    Else
      MsgBox "Nincs jogosults�god", _
               vbOKOnly, "Hi�nyz� jogosults�g"
Exit Sub
End If

End Sub

Private Sub ChangeIll_Click()
If Me.ChangeIll <> "11294" Then
MsgBox "Illetm�ny T�pus�nak m�dos�t�sa ut�n az adat nem lesz import�lhat�!", _
vbOKOnly + vbCritical, "Figyelmeztet�s"

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
MsgBox "Az illetm�ny lez�r�sra ker�lt. Az �j felv�tel�hez kattintson �jra az Illetm�nyadatok gombra!", _
vbOKOnly + vbInformation, "Inform�ci�"
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
MsgBox "Illetm�nyadat megad�sa k�telez�!", vbCritical + vbOKOnly, "Hi�nyz� adat"

Exit Sub
Else

DoCmd.Close acForm, "Illetmeny", acSaveYes
Forms![Foglalkoztatas].[FV_JELMINTA].SetFocus
Forms![Foglalkoztatas].[IlletmenyDone].Value = True

End If

End Sub


