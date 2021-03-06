VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Szemely"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
' Az aktu�lis Ad�azonos�t� egy TempVars-ban t�rolva
Dim adoazon As TempVars
Dim tajazon
Dim vezN
Dim uN
Dim uN2
Dim szulH
Dim szulI
Private Sub manageSzemely()
' gombok l�that�v� t�tele
Forms![Main].[newWorker].Visible = False
Forms![Main].[Address].Visible = True
Forms![Main].[Address].SetFocus
Forms![Main].Form.[Szemely].Visible = False
Forms![Main].[currentStaff].Visible = True
Forms![Main].[vegzett].Visible = True
Forms![Main].[currentAdo].Visible = True
Forms![Main].[currentTaj].Visible = True
Forms![Main].[currentCsal].Visible = True
Forms![Main].[currentUt1].Visible = True
Forms![Main].[currentUt2].Visible = True
Forms![Main].[currenthely].Visible = True
Forms![Main].[currentIdo].Visible = True
Forms![Main].[Foglalkoz].Visible = True
Forms![Main].[vezeto].Visible = True
Forms![Main].[nyelv].Visible = True
Forms![Main].[Reset].Visible = True
Forms![Main].[Printing].Visible = True
Forms![Main].[taj_txt].Visible = True
Forms![Main].[Ado_txt].Visible = True
Forms![Main].[szul_txt].Visible = True
Forms![Main].[Nev_txt].Visible = True
Forms![Main].[modifyMain].Visible = True

' adatok �tad�sa Main formba
Forms![Main].[currentAdo].Value = Me.ado
Forms![Main].[currentTaj].Value = tajazon
Forms![Main].[currentCsal].Value = vezN
Forms![Main].[currentUt1].Value = uN
Forms![Main].[currentUt2].Value = uN2
Forms![Main].[currenthely].Value = szulH
Forms![Main].[currentIdo].Value = szulI
Forms![Main].[isMod].Value = False
Forms![Main].[hasMain].Value = True
' tov�bbi l�tez� adatok ellen�rz�se, annak megfelel�en st�tuszkijelz�s
Dim myDate
myDate = DateValue("12/31/3999")
    If myDate = DLookup("[Valid]", "[Address]", "[Ad_AZON_CIM]='" & TempVars!adoazon & _
    "' And [Export]= False") Then
    Forms![Main].[AddrDone].Visible = True
    Forms![Main].[addressDone].Visible = True
    Forms![Main].[addressDone].Value = True
    Else
    Forms![Main].[AddrDone].Visible = True
    Forms![Main].[addressDone].Visible = True
    Forms![Main].[addressDone].Value = False
    End If
    If myDate = DLookup("[Valid]", "[Foglalkoztatas]", "[Fog_AD_AZON]='" & TempVars!adoazon & _
    "'And [Exported]= False") Then
        Forms![Main].[fogldone].Visible = True
        Forms![Main].[jobDone].Visible = True
        Forms![Main].[jobDone].Value = True
    Else
       Forms![Main].[fogldone].Visible = True
        Forms![Main].[jobDone].Visible = True
        Forms![Main].[jobDone].Value = False
        End If
    If myDate = DLookup("[Valid]", "[Vegzettseg]", "[Kep_AD_AZON]='" & TempVars!adoazon & _
    "' And [Exported]= False") Then
    Forms![Main].[Vegz].Visible = True
    Forms![Main].[educationDone].Visible = True
    Forms![Main].[educationDone].Value = True
    Else
    Forms![Main].[Vegz].Visible = True
    Forms![Main].[educationDone].Visible = True
    Forms![Main].[educationDone].Value = False
    End If
    ' Kiv�laszt�s eset�n adatok t�rl�se, �llapot default-ba rak�sa
Me.ado = Null
Me.vezNev = Null
Me.uNev = Null
Me.uNev2 = Null
Me.szulHely = Null
Me.szulido = Null
Me.taj = Null
Me.chooseTx.Visible = False
Me.vezNev.Visible = False
Me.uNev.Visible = False
Me.uNev2.Visible = False
Me.szulHely.Visible = False
Me.szulido.Visible = False
Me.New.Visible = True
Me.ado.SetFocus
Me.Choose.Visible = False
Me.resetData.Visible = False

End Sub
Private Sub Choose_Click()
Dim fogsor
If Me.Exported <> 0 Then
MsgBox "A szem�lynek r�gz�tett �s export�lt adata van!", vbOKOnly + vbCritical, "Export�lt adatok!"
Me.ado.SetFocus
Exit Sub
ElseIf DLookup("[FOGSOR]", "[Szemely]", "[SZEM_ADOAZON]='" & Me.ado & "'") > 0 _
And DLookup("[JOG_VOLT]", "[Szemely]", "[SZEM_ADOAZON]='" & Me.ado & "'") = True Then
MsgBox "�j jogviszony felv�tele"
Dim ujjog As TempVars
TempVars!ujjog.Value = Me.ado.Value
DoCmd.SetWarnings (WarningsOff)
DoCmd.OpenQuery "OtherJog"
DoCmd.SetWarnings (WarningsOn)
TempVars!ujjog.Value = Null

manageSzemely
Exit Sub
Else
manageSzemely

End If

End Sub

Private Sub Exit_Click()
Forms![Main].[newWorker].Visible = True
Forms![Main].[newWorker].SetFocus
Forms![Main].[Szemely].Visible = False
End Sub

Private Sub Find_Click()
' Ha �res az ad�mez�
If IsNull(Me.ado) Then
MsgBox "Ad�sz�m megad�sa k�telez�", vbOKOnly + vbInformation, "Hi�nyz� Adatok!"
Me.ado.SetFocus
Exit Sub
' Ha tal�l valami �rt�ket az ad�sz�m alapj�n
ElseIf Me.ado.Value = DLookup("[SZEM_ADOAZON]", "Szemely", "[SZEM_ADOAZON]='" & Me.ado.Value & "'") Then
MsgBox "Sikeres tal�lat", vbOKOnly + vbInformation, "Inform�ci�"
      Me.New.Visible = False
      
      TempVars!adoazon = Me.ado.Value
      
'adatok beszerz�se a Szem�ly t�bl�b�l
Me.selected = Me.ado
Me.taj = Me.selected.Column(1)
Me.vezNev = Me.selected.Column(3)
Me.uNev = Me.selected.Column(4)
Me.uNev2 = Me.selected.Column(5)
Me.szulHely = Me.selected.Column(9)
Me.szulido = Me.selected.Column(10)
Me.Exported = Me.selected.Column(28)
'Adatok �tad�sa v�ltoz�knak
tajazon = Me.taj
vezN = Me.vezNev
uN = Me.uNev
uN2 = Me.uNev2
szulH = Me.szulHely
szulI = Me.szulido

'Adatmez�k l�that�
Me.resetData.Visible = True
Me.chooseTx.Visible = True
Me.vezNev.Visible = True
Me.uNev.Visible = True
Me.uNev2.Visible = True
Me.szulHely.Visible = True
Me.Choose.Visible = True



Else
' Ha nem tal�l adatot akkor van lehet�s�g �j r�gz�t�s�re
MsgBox "Az adat nem szerepel,�jk�nt r�gz�thet�", vbOKOnly + vbInformation, "Inform�ci�"

End If




End Sub



Private Sub New_Click()
Dim tajszam As TempVars
Dim szuldatum As TempVars
Dim myDate
Dim dateResult
Dim paros As Integer
Dim paratlan As Integer
Dim eredmeny
Dim lastChar
Dim lastCharOfCount
' �res adatokkal ne lehessen felvinni
If IsNull(Me.ado) Or IsNull(Me.taj) _
Or IsNull(Me.szulido) Then
 MsgBox "Ad�sz�m, sz�let�si id� �s tajsz�m megad�sa k�telez�!", _
 vbOKOnly + vbInformation, "Hi�nyz� Adatok!"
        Me.ado.SetFocus
        Exit Sub
ElseIf Me.ado = DLookup("[SZEM_ADOAZON]", "[Szemely]", "[SZEM_ADOAZON]='" & Me.ado & "'") Then
 MsgBox "A munkav�llal� adatai m�r szerepelnek a rendszerben, adatai csak m�dos�that�ak!", _
 vbOKOnly + vbInformation, "L�tez� munkav�llal�!"
        Me.ado.SetFocus
        Exit Sub
Else
myDate = DateValue("1/1/1867")
dateResult = Int(Me.szulido - myDate)
paratlan = Mid(Me.taj, 1, 1) * 3 + Mid(Me.taj, 3, 1) * 3 + Mid(Me.taj, 5, 1) * 3 + Mid(Me.taj, 7, 1) * 3
paros = Mid(Me.taj, 2, 1) * 7 + Mid(Me.taj, 4, 1) * 7 + Mid(Me.taj, 6, 1) * 7 + Mid(Me.taj, 8, 1) * 7
eredmeny = (paros + paratlan) / 10
lastCharOfCount = Right(eredmeny, 1)
lastChar = Right(Me.taj, 1)
adoData = Int(Mid(Me.ado, 2, 5))
End If
'rossz ad�val ne lehessen felvinni
If adoData <> dateResult Then
    MsgBox "Az ad�azonos�t� �s sz�let�si d�tum k�z�tt elt�r�s van. K�rem ellen�rizni! ", vbCritical _
    + vbOKOnly, "Hib�s adatok."
    Exit Sub
ElseIf lastChar <> lastCharOfCount Then
MsgBox "A taj sz�m az alapszint� ellen�rz�s szerint nem megfelel�, k�rem ellen�rizni!", vbCritical _
+ vbOKOnly, "Hib�s adatok."
Exit Sub
Else
    TempVars!adoazon = Me.ado.Value
    TempVars!tajszam = Me.taj.Value
    TempVars!szuldatum = Me.szulido.Value
     ' Adatok null�z�sa �s megfelel� r�szek l�thatatlann� t�tele
    Me.ado = Null
    Me.vezNev = Null
    Me.uNev = Null
    Me.uNev2 = Null
    Me.szulHely = Null
    Me.szulido = Null
    Me.taj = Null
    Me.chooseTx.Visible = False
    Me.vezNev.Visible = False
    Me.uNev.Visible = False
    Me.uNev2.Visible = False
    Me.szulHely.Visible = False
    Me.New.Visible = True
    Me.Choose.Visible = False
    Me.resetData.Visible = False
    'szem�lyes adatok adatb�zishoz hozz�ad�s
    DoCmd.OpenForm "staffData", , , , acFormAdd
    Forms![staffData].SetFocus
    End If


            
            
End Sub

Private Sub resetData_Click()
' Adatok null�z�sa �s megfelel� r�szek l�thatatlann� t�tele
Me.ado = Null
TempVars!adoazon = Null
Me.vezNev = Null
Me.uNev = Null
Me.uNev2 = Null
Me.szulHely = Null
Me.szulido = Null
Me.taj = Null
Me.chooseTx.Visible = False
Me.vezNev.Visible = False
Me.uNev.Visible = False
Me.uNev2.Visible = False
Me.szulHely.Visible = False
Me.New.Visible = True
Me.Choose.Visible = False
Me.ado.SetFocus
Me.resetData.Visible = False


End Sub

Private Sub taj_AfterUpdate()

End Sub


