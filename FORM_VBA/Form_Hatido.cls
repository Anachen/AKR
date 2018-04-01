VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Hatido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub Form_Load()
If Forms![Foglalkoztatas].[HatDone].Value = True Then
Me.FV_FG_HATFEL = Forms![Foglalkoztatas].[Felad_ellat]
Me.Rogzito = TempVars!selectedUser
Else
Me.FV_FG_HATFEL = Forms![Foglalkoztatas].[Felad_ellat]
Me.ADO_HATFOG = TempVars!adoazon
Me.Rogzito = TempVars!selectedUser
End If
End Sub

Private Sub Save_Click()
DoCmd.Close acForm, "Hatido", acSaveYes
Forms![Foglalkoztatas].[FV_UTMOD].SetFocus
Forms![Foglalkoztatas].[HatDone].Value = True

End Sub




