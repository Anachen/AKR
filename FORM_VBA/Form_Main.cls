VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
Private Sub About_Click()
'About Form megnyit�sa
DoCmd.OpenForm "About", , , , acFormReadOnly
End Sub

Private Sub Exit_Click()
'Kil�p�s a programb�l
DoCmd.Quit
End Sub

Private Sub ProgramState_Click()
'St�tuszk�d ideiglenes v�ltoz� l�trehoz�sa
 Dim statuscode As TempVar
 'St�tusz meghat�roz�sa 1-4
 Dim scode

    scode = InputBox("K�rem adja meg a st�tuszt", "St�tusz megad�sa")
    If Not IsNumeric(scode) Then
        MsgBox "�rv�nytelen adat, pr�b�ld �jra!", vbCritical, "�rv�nytelen adat!"
    Exit Sub
    ElseIf scode = "" Or scode = Empty Then
        MsgBox "K�d megad�sa sz�ks�ges!", vbInformation, "Hi�nyz� adat!"
        Exit Sub
    ElseIf scode > 4 Then
      MsgBox "Rossz K�d!", vbInformation, "Hi�nyz� adat!"
        Exit Sub
'A st�tusz friss�t� Querryk futtat�sa
Else
    TempVars!statuscode = scode
 DoCmd.SetWarnings WarningsOff
 DoCmd.OpenQuery "AC_Update"
 DoCmd.OpenQuery "AC_Add"
DoCmd.SetWarnings WarningsOn
End If
'A St�tusz meghat�roz�sa
Dim Code As Integer
Dim Activator
'Akt�v st�tusz
If DLookup("[Azon]", "[Active]", "[CurrentlyActive]=" & False) = 1 Then
Activator = DLookup("[Activator]", "[Active]", "[CurrentlyActive]= false and [Azon]=1")
Code = 1
Me.Act_Code = Code
Me.Status = Activator
Me!Status.BackColor = vbGreen
'Sz�neteltetve st�tusz
ElseIf DLookup("[Azon]", "[Active]", "[CurrentlyActive]=" & False) = 2 Then
Activator = DLookup("[Activator]", "[Active]", "[CurrentlyActive]= false and [Azon]=2")
Code = 2
Me.Act_Code = Code
Me.Status = Activator
Me!Status.BackColor = vbYellow
'Z�rolva st�tusz
ElseIf DLookup("[Azon]", "[Active]", "[CurrentlyActive]=" & False) = 3 Then
Activator = DLookup("[Activator]", "[Active]", "[CurrentlyActive]= false and [Azon]=3")
Code = 3
Me.Act_Code = Code
Me.Status = Activator
Me!Status.BackColor = vbRed
'Teszt�zem st�tusz
Else
Code = 4
Activator = DLookup("[Activator]", "[Active]", "[CurrentlyActive]= false and [Azon]=4")
Me.Act_Code = Code
Me.Status = Activator
Me!Status.BackColor = vbWhite
End If

End Sub

    

Private Sub Queries_Click()
'Lek�rdez�sfunkci�k
DoCmd.OpenForm "Querries"

End Sub

Private Sub Help_Click()
'�tmutat� megnyit�si folyamat
Dim openHelp
openHelp = MsgBox("Az �tmutat� megnyit�sa", vbInformation + vbOKCancel, "S�g� megnyit�sa")
If openHelp = vbCancel Then
Exit Sub
Else
Application.FollowHyperlink "........docx"
End If

End Sub

Private Sub Address_Click()
Dim myDate
'lakc�madatok megnyit�sa, ha van, vagy �j lakc�madat felv�tele
myDate = DateValue("12/31/3999")
If myDate = DLookup("[Valid]", "[Address]", "[Ad_AZON_CIM]='" _
& TempVars!adoazon & "' And [Export]= False") Then
MsgBox "A munkav�llal�nak m�r van �rv�nyes lakc�madata, k�rem ellen�rizze helyess�g�t!", _
vbOKOnly + vbCritical, "Figyelmeztet�s"

'ha van megfelel� adat akkor a megfelel� adatokat t�lti be.
DoCmd.OpenForm "AddressData", , "[Valid]='" & myDate & "'", _
"[AD_AZON_CIM] ='" & TempVars!adoazon & "'and [Export] = False", acFormEdit
Forms![AddressData].[Renew].Visible = True
Forms![AddressData].[Submit].Visible = True
Forms![AddressData].[CIM_TELNEV].Visible = True
Forms![AddressData].[tel].Visible = True
Forms![AddressData].[CIM_KHELYSEG].Visible = True
Forms![AddressData].[ktel].Visible = True
Else
DoCmd.OpenForm "AddressData", , , , acFormAdd
Forms![AddressData].[Submit].Visible = True
End If


End Sub


Private Sub Foglalkoz_Click()
Dim myDate
myDate = DateValue("12/31/3999")
'folglakoztat�s hozz�ad�sa vagy a szerkeszt�se

If myDate = DLookup("[Valid]", "[Foglalkoztatas]", "[Fog_AD_AZON]='" & TempVars!adoazon & _
"' And [Exported]= False") Then
MsgBox "A munkav�llal�nak m�r van �rv�nyes Foglalkoztat�si adata! Amennyiben �jat szeretne felvenni, z�rja le a jelenlegit!", _
vbOKOnly + vbInformation, "Figyelmeztet�s"
'adatok bet�lt�se, ha van
DoCmd.OpenForm "Foglalkoztatas", , "[Valid]='" & myDate & "'", _
"[Fog_AD_AZON] ='" & TempVars!adoazon & "'and [MODOSITHATO]= false", acFormEdit
Forms![Foglalkoztatas].[Hat].Visible = True
Forms![Foglalkoztatas].[Folyoszamla].Visible = True
Me.Foglalkoz.Visible = True
Else
DoCmd.OpenForm "Foglalkoztatas", , , , acFormAdd

End If
End Sub

'�ra felhelyez�se
Private Sub Form_Timer()
Clock = Now()
End Sub

'A k�zponti form bet�lt�s�nek met�dusa
Private Sub Form_Load()
Me.TimerInterval = 1000
Dim Code As Integer
Dim Activator
Dim Verz As String

'id�,verzi� �s felhaszn�l�k a jelenlegi st�tusz bet�lt�se
Me.getDate = DateValue(Now())
Me.getTime = Now()
Me.getDay = Weekday(Me.getDate, vbMonday)
Verz = DMax("Verzio", "Version")
Me.Version = Verz
Me.currentUser = TempVars!selectedUser
If DLookup("[Azon]", "[Active]", "[CurrentlyActive]=" & False) = 1 Then
Activator = DLookup("[Activator]", "[Active]", "[CurrentlyActive]= false and [Azon]=1")
Code = 1
Me.Act_Code = Code
Me.Status = Activator
Me!Status.BackColor = vbGreen
ElseIf DLookup("[Azon]", "[Active]", "[CurrentlyActive]=" & False) = 2 Then
Activator = DLookup("[Activator]", "[Active]", "[CurrentlyActive]= false and [Azon]=2")
Code = 2
Me.Act_Code = Code
Me.Status = Activator
Me!Status.BackColor = vbYellow
ElseIf DLookup("[Azon]", "[Active]", "[CurrentlyActive]=" & False) = 3 Then
Activator = DLookup("[Activator]", "[Active]", "[CurrentlyActive]= false and [Azon]=3")
Code = 3
Me.Act_Code = Code
Me.Status = Activator
Me!Status.BackColor = vbRed
Else
Code = 4
Me.Act_Code = Code
Me.Status = "Teszt�zem"
Me!Status.BackColor = vbWhite


End If


End Sub


Private Sub Navigate_Enter()
DoCmd.OpenForm "Szemely"
End Sub

'munkav�llal� hozz�ad�sa
Private Sub newWorker_Click()
Dim currentTime
Dim notAllowTime
Dim state
Dim stateName
Dim dayName
'Friss�ti azaktu�lisid�t, minden�j munkav�llal�felv�tel�n�l
Me.hasMain = False
Me.getTime = Me.Clock
'Megkeresi a legutols� akt�v �llapotot
state = DLookup("[Azon]", "[Active]", "[CurrentlyActive] = False")
stateName = DLookup("[Activator]", "[Active]", "[CurrentlyActive] = False")
'�t�rja a legutols� akt�v �llapotra az �llapotot
Me.Act_Code = state
Me.Status = stateName
'az aktu�lis id� napj�nak meghat�roz�sa, �tad�sa dayName-nek
'az aktu�lis id� �tad�sa id��rt�knek

currentTime = TimeValue(Now())


'A program haszn�lhat�s�g�nak felt�telmegad�sa
If Me.getDay = 5 Then
notAllowTime = TimeValue("11:50") '11:50
Else
notAllowTime = TimeValue("15:50") '15:50
End If
'A program, ha haszn�lhat�s�gi id�n k�v�njukhaszn�lni
If currentTime > notAllowTime And Me.getDay = 5 Then
MsgBox "A Program csak 8:00 - 12:00 k�z�tt haszn�lhat�!", vbInformation + vbOKOnly, "Z�rol�s"
Exit Sub
ElseIf currentTime > notAllowTime And Me.getDay <> 5 Then
MsgBox "A Program csak 8:00 - 16:00 k�z�tt haszn�lhat�!", vbInformation + vbOKOnly, "Z�rol�s"
Exit Sub
Else
End If

'A f� fel�letProgram st�tusz�nak sz�nbe�ll�t�sa
If Me.Act_Code = 1 Then
Me!Status.BackColor = vbGreen
ElseIf Me.Act_Code = 2 Then
Me!Status.BackColor = vbYellow
ElseIf Me.Act_Code = 3 Then
Me!Status.BackColor = vbRed
Else
Me!Status.BackColor = vbWhite
End If

If Me.Act_Code = 4 Then
MsgBox "A program jelenleg Tesztel�s alatt van, a felr�gz�tett adatok nem lesznek alkalmasak import�l�sra", _
vbInformation + vbOKOnly, "Tesztel�s"
Me.Szemely.Visible = True
Me.Szemely.SetFocus
ElseIf Me.Act_Code <> 1 Then
MsgBox "A program jelenleg fejleszt�s alatt van, vagy Export�l�s zajlik, ez�rt nem haszn�lhat�", _
vbCritical + vbOKOnly, "Z�rolt program"
Exit Sub
Else
Me.Szemely.Visible = True
Me.Szemely.SetFocus
End If
End Sub

'felr�gz�tett munkav�llal� nyomtat�si report funkci�ja
Private Sub Printing_Click()
Dim surePrint As Integer
Dim printAdo As TempVars
Dim userName As TempVars
If Me.addressDone = False Or Me.jobDone = False Or Me.educationDone = False Then
MsgBox "Addig nem lehet nyomtatni, am�g minden sz�ks�ges adatot fel nem viszel!", vbCritical + vbOKOnly, "Hi�nyz� adatok"
Exit Sub
End If
surePrint = MsgBox("Val�ban meg szeretn�d nyitni a " & Me.currentAdo & _
" ad�sz�mmal rendelkez� munkav�llal� adatainak nyomtat�si k�p�t?", vbYesNo + vbQuestion, "Nyomtat�s")
If Me.addressDone.Value = False Or Me.jobDone.Value = False Or Me.educationDone.Value = False Then
MsgBox " A munkav�llal�nak nincs valamilyen adata r�gz�tve, a n�lk�l nem nyomtathat�! Ellen�r�zd a felr�gz�tett adatokat!" _
, vbCritical + vbOKOnly, "Hi�nyz� Adatok"
Exit Sub
ElseIf surePrint = vbNo Then
Exit Sub
Else
TempVars!printAdo = Me.currentAdo.Value
TempVars!userName = Me.currentUser.Value
DoCmd.SetWarnings WaringsOff
DoCmd.OpenQuery "Egyesites", , acAdd
DoCmd.OpenReport "Print", acViewReport, "[SZEM_ADOAZON]='" & Me.currentAdo & "'", _
"[SZEM_ADOAZON]='" & Me.currentAdo & "' and [Printed] = False", acDialog
DoCmd.OpenQuery "printDone"
DoCmd.SetWarnings WaringsOn
TempVars!printAdo = Null
TempVars!userName = Null
End If

End Sub


'Az aktu�lisan szerkesztett adatokb�l val� kil�p�s
Private Sub Reset_Click()
Dim answer As Integer
answer = MsgBox("A " & Me.currentAdo & " ad�az�nos�t�val rendelkez� munkav�llal� adatai felold�sa?", _
vbYesNo + vbQuestion, "K�rd�s")
If Me.addressDone.Value = False Or Me.jobDone.Value = False Or Me.educationDone.Value = False Then
MsgBox " A munkav�llal�nak nincs valamilyen adata r�gz�tve, a n�lk�l nem  lehet feloldani! Ellen�r�zd a felr�gz�tett adatokat!" _
, vbCritical + vbOKOnly, "Hi�nyz� Adatok"
Exit Sub
ElseIf answer = vbYes Then
TempVars!adoazon = Null
Me.currentAdo.Visible = False
Me.currentTaj.Visible = False
Me.currentCsal.Visible = False
Me.currentUt1.Visible = False
Me.currentUt2.Visible = False
Me.currenthely.Visible = False
Me.currentIdo.Visible = False
Me.JogKezd.Visible = False
Me.jogkezd_txt.Visible = False
Me.modifyMain.Visible = False
'kiv�lasztott adatokhoz tartoz� sz�veg l�thatalna lesz
Me.taj_txt.Visible = False
Me.Ado_txt.Visible = False
Me.szul_txt.Visible = False
Me.Nev_txt.Visible = False

'modos�t�s reset
Me.isMod.Value = False
Me.modHat = Null
Me.modHat.Visible = False
Me.modHat_txt.Visible = False
'sz�ml�t figyel� reset
Me.SzamlaDone.Value = False
Me.SzamlaDone.Visible = False
Me.szamlaDone_txt.Visible = False


Me.currentAdo.Value = Null
Me.currentTaj.Value = Null
Me.currentCsal.Value = Null
Me.currentUt1.Value = Null
Me.currentUt2.Value = Null
Me.currenthely.Value = Null
Me.currentIdo.Value = Null
Me.JogKezd.Value = Null

Me.newWorker.Visible = True
Me.ChangeStaff.Visible = True


Me.Address.Visible = False
Me.Foglalkoz.Visible = False
Me.vegzett.Visible = False
Me.vezeto.Visible = False
Me.nyelv.Visible = False
Me.Printing.Visible = False
Me.szamla.Visible = False



Me.AddrDone.Visible = False
Me.addressDone.Visible = False
Me.jobDone.Visible = False
Me.fogldone.Visible = False
Me.educationDone.Visible = False
Me.Vegz.Visible = False

Me.educationDone.Value = False
Me.jobDone.Value = False
Me.addressDone.Value = False

Me.newWorker.SetFocus

Me.Reset.Visible = False

Else
Exit Sub
End If

End Sub

'v�gzetts�g adatok hozz�ad�sa
Private Sub vegzett_Click()
Dim myDate
myDate = DateValue("12/31/3999")
If myDate = DLookup("[Valid]", "[Vegzettseg]", "[Kep_AD_AZON]='" & TempVars!adoazon & "' And [Exported]= False") Then
MsgBox "A munkav�llal�nak m�r �rv�nyes V�gzetts�ge! Ellen�rizze a jelenlegi helyess�g�t!", _
vbOKOnly + vbInformation, "Figyelmeztet�s"
DoCmd.OpenForm "Vegzettseg", , "[Valid]='" & myDate & "'", _
"[Kep_AD_AZON] ='" & TempVars!adoazon & "' And [Exported]= False", acFormEdit
Forms![Vegzettseg].[ISK_KVEGYEB].Visible = True
Forms![Vegzettseg].[egyeb].Visible = True

Else

DoCmd.OpenForm "Vegzettseg", , , , acFormAdd

End If


End Sub