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
'About Form megnyitása
DoCmd.OpenForm "About", , , , acFormReadOnly
End Sub

Private Sub Exit_Click()
'Kilépés a programból
DoCmd.Quit
End Sub

Private Sub ProgramState_Click()
'Státuszkód ideiglenes változó létrehozása
 Dim statuscode As TempVar
 'Státusz meghatározása 1-4
 Dim scode

    scode = InputBox("Kérem adja meg a státuszt", "Státusz megadása")
    If Not IsNumeric(scode) Then
        MsgBox "Érvénytelen adat, próbáld újra!", vbCritical, "Érvénytelen adat!"
    Exit Sub
    ElseIf scode = "" Or scode = Empty Then
        MsgBox "Kód megadása szükséges!", vbInformation, "Hiányzó adat!"
        Exit Sub
    ElseIf scode > 4 Then
      MsgBox "Rossz Kód!", vbInformation, "Hiányzó adat!"
        Exit Sub
'A státusz frissítõ Querryk futtatása
Else
    TempVars!statuscode = scode
 DoCmd.SetWarnings WarningsOff
 DoCmd.OpenQuery "AC_Update"
 DoCmd.OpenQuery "AC_Add"
DoCmd.SetWarnings WarningsOn
End If
'A Státusz meghatározása
Dim Code As Integer
Dim Activator
'Aktív státusz
If DLookup("[Azon]", "[Active]", "[CurrentlyActive]=" & False) = 1 Then
Activator = DLookup("[Activator]", "[Active]", "[CurrentlyActive]= false and [Azon]=1")
Code = 1
Me.Act_Code = Code
Me.Status = Activator
Me!Status.BackColor = vbGreen
'Szüneteltetve státusz
ElseIf DLookup("[Azon]", "[Active]", "[CurrentlyActive]=" & False) = 2 Then
Activator = DLookup("[Activator]", "[Active]", "[CurrentlyActive]= false and [Azon]=2")
Code = 2
Me.Act_Code = Code
Me.Status = Activator
Me!Status.BackColor = vbYellow
'Zárolva státusz
ElseIf DLookup("[Azon]", "[Active]", "[CurrentlyActive]=" & False) = 3 Then
Activator = DLookup("[Activator]", "[Active]", "[CurrentlyActive]= false and [Azon]=3")
Code = 3
Me.Act_Code = Code
Me.Status = Activator
Me!Status.BackColor = vbRed
'Tesztüzem státusz
Else
Code = 4
Activator = DLookup("[Activator]", "[Active]", "[CurrentlyActive]= false and [Azon]=4")
Me.Act_Code = Code
Me.Status = Activator
Me!Status.BackColor = vbWhite
End If

End Sub

    

Private Sub Queries_Click()
'Lekérdezésfunkciók
DoCmd.OpenForm "Querries"

End Sub

Private Sub Help_Click()
'Útmutató megnyitási folyamat
Dim openHelp
openHelp = MsgBox("Az Útmutató megnyitása", vbInformation + vbOKCancel, "Súgó megnyitása")
If openHelp = vbCancel Then
Exit Sub
Else
Application.FollowHyperlink "........docx"
End If

End Sub

Private Sub Address_Click()
Dim myDate
'lakcímadatok megnyitása, ha van, vagy új lakcímadat felvétele
myDate = DateValue("12/31/3999")
If myDate = DLookup("[Valid]", "[Address]", "[Ad_AZON_CIM]='" _
& TempVars!adoazon & "' And [Export]= False") Then
MsgBox "A munkavállalónak már van érvényes lakcímadata, kérem ellenõrizze helyességét!", _
vbOKOnly + vbCritical, "Figyelmeztetés"

'ha van megfelelö adat akkor a megfelelö adatokat tölti be.
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
'folglakoztatás hozzáadása vagy a szerkesztése

If myDate = DLookup("[Valid]", "[Foglalkoztatas]", "[Fog_AD_AZON]='" & TempVars!adoazon & _
"' And [Exported]= False") Then
MsgBox "A munkavállalónak már van érvényes Foglalkoztatási adata! Amennyiben újat szeretne felvenni, zárja le a jelenlegit!", _
vbOKOnly + vbInformation, "Figyelmeztetés"
'adatok betöltése, ha van
DoCmd.OpenForm "Foglalkoztatas", , "[Valid]='" & myDate & "'", _
"[Fog_AD_AZON] ='" & TempVars!adoazon & "'and [MODOSITHATO]= false", acFormEdit
Forms![Foglalkoztatas].[Hat].Visible = True
Forms![Foglalkoztatas].[Folyoszamla].Visible = True
Me.Foglalkoz.Visible = True
Else
DoCmd.OpenForm "Foglalkoztatas", , , , acFormAdd

End If
End Sub

'óra felhelyezése
Private Sub Form_Timer()
Clock = Now()
End Sub

'A központi form betöltésének metódusa
Private Sub Form_Load()
Me.TimerInterval = 1000
Dim Code As Integer
Dim Activator
Dim Verz As String

'idö,verzió és felhasználók a jelenlegi státusz betöltése
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
Me.Status = "Tesztüzem"
Me!Status.BackColor = vbWhite


End If


End Sub


Private Sub Navigate_Enter()
DoCmd.OpenForm "Szemely"
End Sub

'munkavállaló hozzáadása
Private Sub newWorker_Click()
Dim currentTime
Dim notAllowTime
Dim state
Dim stateName
Dim dayName
'Frissíti azaktuálisidõt, mindenúj munkavállalófelvételénél
Me.hasMain = False
Me.getTime = Me.Clock
'Megkeresi a legutolsó aktív állapotot
state = DLookup("[Azon]", "[Active]", "[CurrentlyActive] = False")
stateName = DLookup("[Activator]", "[Active]", "[CurrentlyActive] = False")
'Átírja a legutolsó aktív állapotra az állapotot
Me.Act_Code = state
Me.Status = stateName
'az aktuális idõ napjának meghatározása, átadása dayName-nek
'az aktuális idõ átadása idõértéknek

currentTime = TimeValue(Now())


'A program használhatóságának feltételmegadása
If Me.getDay = 5 Then
notAllowTime = TimeValue("11:50") '11:50
Else
notAllowTime = TimeValue("15:50") '15:50
End If
'A program, ha használhatósági idõn kívánjukhasználni
If currentTime > notAllowTime And Me.getDay = 5 Then
MsgBox "A Program csak 8:00 - 12:00 között használható!", vbInformation + vbOKOnly, "Zárolás"
Exit Sub
ElseIf currentTime > notAllowTime And Me.getDay <> 5 Then
MsgBox "A Program csak 8:00 - 16:00 között használható!", vbInformation + vbOKOnly, "Zárolás"
Exit Sub
Else
End If

'A fõ felületProgram státuszának színbeállítása
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
MsgBox "A program jelenleg Tesztelés alatt van, a felrögzített adatok nem lesznek alkalmasak importálásra", _
vbInformation + vbOKOnly, "Tesztelés"
Me.Szemely.Visible = True
Me.Szemely.SetFocus
ElseIf Me.Act_Code <> 1 Then
MsgBox "A program jelenleg fejlesztés alatt van, vagy Exportálás zajlik, ezért nem használható", _
vbCritical + vbOKOnly, "Zárolt program"
Exit Sub
Else
Me.Szemely.Visible = True
Me.Szemely.SetFocus
End If
End Sub

'felrögzített munkavállaló nyomtatási report funkciója
Private Sub Printing_Click()
Dim surePrint As Integer
Dim printAdo As TempVars
Dim userName As TempVars
If Me.addressDone = False Or Me.jobDone = False Or Me.educationDone = False Then
MsgBox "Addig nem lehet nyomtatni, amíg minden szükséges adatot fel nem viszel!", vbCritical + vbOKOnly, "Hiányzó adatok"
Exit Sub
End If
surePrint = MsgBox("Valóban meg szeretnéd nyitni a " & Me.currentAdo & _
" adószámmal rendelkezõ munkavállaló adatainak nyomtatási képét?", vbYesNo + vbQuestion, "Nyomtatás")
If Me.addressDone.Value = False Or Me.jobDone.Value = False Or Me.educationDone.Value = False Then
MsgBox " A munkavállalónak nincs valamilyen adata rögzítve, a nélkül nem nyomtatható! Ellenõrízd a felrögzített adatokat!" _
, vbCritical + vbOKOnly, "Hiányzó Adatok"
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


'Az aktuálisan szerkesztett adatokból való kilépés
Private Sub Reset_Click()
Dim answer As Integer
answer = MsgBox("A " & Me.currentAdo & " adóazónosítóval rendelkezõ munkavállaló adatai feloldása?", _
vbYesNo + vbQuestion, "Kérdés")
If Me.addressDone.Value = False Or Me.jobDone.Value = False Or Me.educationDone.Value = False Then
MsgBox " A munkavállalónak nincs valamilyen adata rögzítve, a nélkül nem  lehet feloldani! Ellenõrízd a felrögzített adatokat!" _
, vbCritical + vbOKOnly, "Hiányzó Adatok"
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
'kiválasztott adatokhoz tartozó szöveg láthatalna lesz
Me.taj_txt.Visible = False
Me.Ado_txt.Visible = False
Me.szul_txt.Visible = False
Me.Nev_txt.Visible = False

'modosítás reset
Me.isMod.Value = False
Me.modHat = Null
Me.modHat.Visible = False
Me.modHat_txt.Visible = False
'számlát figyelõ reset
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

'végzettség adatok hozzáadása
Private Sub vegzett_Click()
Dim myDate
myDate = DateValue("12/31/3999")
If myDate = DLookup("[Valid]", "[Vegzettseg]", "[Kep_AD_AZON]='" & TempVars!adoazon & "' And [Exported]= False") Then
MsgBox "A munkavállalónak már érvényes Végzettsége! Ellenõrizze a jelenlegi helyességét!", _
vbOKOnly + vbInformation, "Figyelmeztetés"
DoCmd.OpenForm "Vegzettseg", , "[Valid]='" & myDate & "'", _
"[Kep_AD_AZON] ='" & TempVars!adoazon & "' And [Exported]= False", acFormEdit
Forms![Vegzettseg].[ISK_KVEGYEB].Visible = True
Forms![Vegzettseg].[egyeb].Visible = True

Else

DoCmd.OpenForm "Vegzettseg", , , , acFormAdd

End If


End Sub