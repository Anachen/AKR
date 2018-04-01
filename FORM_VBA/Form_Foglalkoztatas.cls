VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Foglalkoztatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub bertablak_Click()
Me.FV_BERTABLA = Me.bertablak
Me.FV_FG_BERTABLA = Me.bertablak
Me.Bert_Kod = Me.FV_BERTABLA.Column(1)

End Sub



Private Sub FG_ELSOJOGV_Click()
If Me.FG_ELSOJOGV = "I" Then
MsgBox "Elsõ jogviszony esetény kérem az iskolai tanulmányok végét megadni!", _
vbOKOnly + vbExclamation, "További adat szükséges!"
Me.iskvegeTxt.Visible = True
Me.FG_ISKVEGE.Visible = True
End If




End Sub

Private Sub Folyoszamla_Click()

Dim myDate
myDate = DateValue("31/12/3999")
If myDate = DLookup("[Valid]", "[SZAMLA]", "[SZAMLA_AD_AZON]='" & TempVars!adoazon & _
"' and [Exported] = false") Then
MsgBox "A munkavállalónak már van rögzítve Folyószámla.", _
vbOKOnly + vbCritical, "Figyelmeztetés"
DoCmd.OpenForm "Szamla", , "[Valid]='" & myDate & "'" _
, "[SZAMLA_AD_AZON] = '" & TempVars!adoazon & "' and [Exported] = false", acFormEdit



Else
DoCmd.OpenForm "Szamla", , , , acFormAdd

End If



End Sub

Private Sub Form_Load()
Dim myDate
myDate = DateValue("12/31/3999")
If Forms![Main].[jobDone].Value = False Then
Me.Fog_AD_AZON = TempVars!adoazon
Me.Rogzito = TempVars!selectedUser
Else
Me.Rogzito = TempVars!selectedUser
Me.Felad_ellat.Visible = True
Me.Szerv_Data = Me.Egyedi_SzerV_AZON
Me.SZERV_NEV = Me.Szerv_Data.Column(2)
Me.TemaNev = Me.Szerv_Data.Column(3)
Me.Rogzito = TempVars!selectedUser
Me.Szervek = Me.Szerv_Data
Me.Cofog_kod = Me.FV_COFOG.Column(0)
Me.bertablak = Me.FV_BERTABLA
Me.FV_FG_BERTABLA = Me.bertablak
Me.Bert_Kod = Me.FV_BERTABLA.Column(1)
Me.FEOR_name = Me.FV_FEOR.Column(1)
End If

If myDate = DLookup("[Valid]", "[HATFOG]", "[ADO_HATFOG]='" & TempVars!adoazon & _
"' And [Exported]= False") Then
Me.Felad_ellat = DLookup("[FV_FG_HATFEL]", "[HATFOG]", "[ADO_HATFOG]='" & TempVars!adoazon _
& "' AND Valid = Datevalue('31/12/3999')")
Forms![Foglalkoztatas].[HatDone].Value = True
Else
Forms![Foglalkoztatas].[HatDone].Value = False
End If
If myDate = DLookup("[Valid]", "[Szamla]", "[SZAMLA_AD_AZON]='" & TempVars!adoazon & _
"' And [Exported]= False") Then
Forms![Foglalkoztatas].[SzamlaDone].Value = True
Else
Forms![Foglalkoztatas].[SzamlaDone].Value = False
End If
If myDate = DLookup("[Valid]", "[Illetmeny]", "[Ill_AD_AZON]='" & TempVars!adoazon & _
"' And [Exported]= False") Then
Forms![Foglalkoztatas].[IlletmenyDone].Value = True
Else
Forms![Foglalkoztatas].[IlletmenyDone].Value = False
End If

End Sub


Private Sub FV_COFOG_Click()
Me.Cofog_kod = Me.FV_COFOG.Column(0)

End Sub

Private Sub FV_FEOR_afterUpdate()
Me.FEOR_name = FV_FEOR.Column(1)
End Sub

Private Sub FV_FG_JIDOTARTAM_Click()
Dim idotart As Integer
idotart = Me.FV_FG_JIDOTARTAM
If idotart = 2 Then
MsgBox "Kérlek állítsd be a határozott idõt", _
vbOKOnly + vbExclamation, "További adat szükséges!"
Me.Hat.Visible = True
Me.Hat.SetFocus
ElseIf idotart <> 1 Then
MsgBox "Csak határozott és határozatlan állítható be!", _
vbOKOnly + vbCritical, "Figyelmeztetés"
Exit Sub
Else
Me.FV_UTMOD.SetFocus
End If

End Sub

Private Sub FV_FG_NEV_Click()
Me.Telhely = Me.FV_FG_NEV
Me.FV_AZON = Me.Telhely.Column(0)
Me.FV_KJAZON = Me.Telhely.Column(2)
Me.FV_KJNEV = Me.Telhely.Column(3)
Me.FV_MJAZON = Me.Telhely.Column(4)
Me.FV_MJNEV = Me.Telhely.Column(5)



End Sub

Private Sub FV_FMIDOJEL_Click()
If Me.FV_FMIDOJEL = 1 Then
MsgBox "Teljes munkaidõs jogviszony esetén az alapértelmezett heti 40 óra beállítva.", _
vbOKOnly + vbInformation, "Információ"
Me.FV_TEHOSZ = 40
Me.FV_FSTATUSZ.SetFocus
Else
MsgBox "Részmunkaidõs esetén kérem a munkaidõ számítás típusát és tényleges hosszát megadni!", _
vbOKOnly + vbExclamation, "További adat szükséges!"
Me.FV_MTIP.SetFocus
End If

End Sub

Private Sub FV_FSTATUSZ_Click()
If Me.FV_FSTATUSZ = "03" Then
    MsgBox "A Feladat ellátásának megnevezését adja meg!", _
    vbOKOnly + vbExclamation, "További adat szükséges!"
    Me.Felad_ellat.Visible = True
    Me.Felad_ellat.SetFocus
    ElseIf Me.FV_FSTATUSZ <> "01" Then
    
    MsgBox "Adatokat rögzíthetõek, de csak 'Állás terhére' vagy 'Feladat ellátására'" _
    & "opciók esetére  lehet csak importálni!", _
    vbOKOnly + vbCritical, "Figyelmeztetés"

    End If


End Sub

Private Sub FV_JELMINTA_Click()
Dim JogKezd
Dim monthText
Dim dateText
Dim dayNumber
Dim jelOraErtek As Integer
Dim vegzettOra As Integer
Me.Jelminta_telj = Me.FV_JELMINTA

jelOraErtek = Me.Jelminta_telj.Column(2)
vegzettOra = Me.FV_TEHOSZ
If jelOraErtek <> vegzettOra Then
MsgBox "A Kiválasztott jelminta nem egyezik a megadott heti óraszámmal! Kérem ellenõrizze!" & vbNewLine & _
"A " & Me.FV_TEHOSZ & " nem egyezik meg a jelminta értékkel: " & jelOraErtek, _
vbOKOnly + vbInformation, "Információ"
Me.FV_JELMINTA = ""
Exit Sub
Else
MsgBox "A Jelminta minta automatikusan kitöltve. Kérem ellenõrizze!" & vbNewLine & _
"A jelenléti kezdete a belépési hónap elsõ napja." & vbNewLine & _
"a jelenléti napja a hónap elsõ napjának naptári napja (Héftõ-Vasárnap)", _
vbOKOnly + vbInformation, "Információ"
JogKezd = Me.FG_JOGVKEZDETE
If month(JogKezd) < 10 Then
monthText = "0" & month(JogKezd)
Else
monthText = month(JogKezd)
End If
dateText = Year(JogKezd) & "." & monthText & ".01"
Me.FV_JELKEZD = dateText
dayNumber = Weekday(Me.FV_JELKEZD, vbMonday)
Me.FV_INDITAS = dayNumber
End If


End Sub

Private Sub FV_JOGTIPKOD_Click()
If Me.FV_JOGTIPKOD <> "06" Then
MsgBox "Az adatfelvétel megtörténik, de importálásra az adat nem lesz alkalmas!", _
vbOKOnly + vbCritical, "Figyelmeztetés"
End If


End Sub

Private Sub FV_MTIP_Click()
If Me.FV_MTIP <> 2 Then
MsgBox "A Tényleges hosszt és törvényes hosszt kérem megadni, óra esetén a havi óraszámot is!", _
vbOKOnly + vbExclamation, "További adat szükséges!"
Me.FV_TEHOSZ.SetFocus
Me.FV_FG_HAVIORA.Visible = True

End If


End Sub

Private Sub FV_UTMOD_Click()
If Me.FV_UTMOD = 2 Then
MsgBox "Állítsa be a folyószámlaszámot!", _
vbOKOnly + vbExclamation, "További adat szükséges!"
Me.Folyoszamla.Visible = True
Me.Folyoszamla.SetFocus
Else
MsgBox "Ha nem nem folyószámlára utalás van, akkor csak a 'Lakcímre', vagy 'Lakcímre, de postán marad' importálható!", _
vbOKOnly + vbCritical, "Figyelmeztetés"
End If

End Sub

Private Sub Hat_Click()
Dim myDate
myDate = DateValue("12/31/3999")
If Me.FV_FSTATUSZ = "03" And IsNull(Me.Felad_ellat) = True Then
MsgBox "Feladat ellátását választottál, de nem adtál meg feladatot!", vbCritical + vbOKOnly _
, "Hiányzó adat"
Exit Sub
ElseIf myDate = DLookup("[Valid]", "[HATFOG]", "[ADO_HATFOG]='" & TempVars!adoazon & _
"'and [Exported] = False") Then
MsgBox "A munkavállalónak már van rögzítve határozott idõ rögzítve!", _
vbOKOnly + vbCritical, "Figyelmeztetés"
DoCmd.OpenForm "Hatido", , "[Valid]='" & myDate & "'", _
"[ADO_HATFOG] = '" & TempVars!adoazon & "' and [Exported] = False", acFormEdit

Else
DoCmd.OpenForm "Hatido", , , , acFormAdd

End If



End Sub

Private Sub Illetm_Click()
Dim myDate
myDate = DateValue("12/31/3999")
If myDate = DLookup("[Valid]", "[Illetmeny]", "[Ill_AD_AZON]='" & TempVars!adoazon & _
"' and [Exported] = false") Then
MsgBox "A munkavállalónak már van Illetményadata. Kérem ellenõrizze!", _
vbOKOnly + vbCritical, "Figyelmeztetés"
DoCmd.OpenForm "Illetmeny", , "[Valid]='" & myDate & "'", _
"[Ill_AD_AZON] = '" & TempVars!adoazon & "'and [Exported] = false", acFormEdit

Else
DoCmd.OpenForm "Illetmeny", , , , acFormAdd

End If

End Sub

Private Sub Save_Click()
Dim myDate
myDate = DateValue("12/31/3999")
Dim jelOraErtek As Integer
Dim vegzettOra As Integer
If Me.IlletmenyDone = False Then
MsgBox "Jelenleg úgy tûnik nem rögzítettél illetményadatot, kérlek ellenõrizd!", _
vbCritical + vbOKOnly, "Hiányzó illetmény"
Exit Sub
ElseIf Me.Felad_ellat <> DLookup("[FV_FG_HATFEL]", "[HATFOG]", _
"[ADO_HATFOG]='" & TempVars!adoazon & "' and [Exported] = False") Then
MsgBox "A Foglalkozatási lapon megadott feladat nem egyezik az érvényes adattal, kérlek frissítsd!" _
, vbCritical + vbOKOnly, "Eltérõ adatok"
ElseIf Me.FV_FSTATUSZ = "03" And IsNull(Me.Felad_ellat) = True Then
MsgBox "Feladat ellátását választottál, de nem adtál meg feladatot!", vbCritical + vbOKOnly _
, "Hiányzó adat"
Exit Sub
ElseIf (Me.FV_FG_JIDOTARTAM <> 2 And Me.HatDone = True) _
Or (Me.FV_FG_JIDOTARTAM = 2 And Me.HatDone = False) Then
MsgBox "Határozott idõs adat van, de te határozatlant akarsz felvinni!" _
, vbCritical + vbOKOnly, "Típuseltérési hiba"
Exit Sub
ElseIf (Me.FV_UTMOD <> 2 And Me.SzamlaDone = True) _
Or (Me.FV_UTMOD = 2 And Me.SzamlaDone = False) Then
MsgBox "Számlaadat van, de te más típusú adatot akarsz felvinni!" _
, vbCritical + vbOKOnly, "Típuseltérési hiba"
Exit Sub
ElseIf Me.FG_JOGVKEZDETE = "" Or Me.FV_SZERVEGYSEG = "" Or _
Me.FV_TEMASZAM = "" Or Me.FV_FKSZLA = "" Or Me.FV_BERTABLA = "" Or Me.FV_FG_BERTABLA = "" Or _
IsNull(Me.FV_COFOG) = True Or Me.FG_MLTIP = "" Or Me.FV_MUNKAREND = "" Or Me.FV_MIDOBEOSZT = "" Or _
Me.FV_FMUSZAKREND = "" Or Me.FV_FMIDOJEL = "" Or Me.FV_FSTATUSZ = "" Or Me.FV_MTIP = "" Or _
IsNull(Me.FV_TEHOSZ) = True Or Me.FV_FG_NEV = "" Or Me.FV_FG_JIDOTARTAM = "" Or Me.FV_UTMOD = "" Or _
IsNull(Me.FV_FG_KULCSSZAM) = True Or IsNull(Me.FV_FEOR) = True Or Me.FV_MKOR = "" Or Me.FV_JELMINTA = "" Or _
Me.FV_JELKEZD = "" Or Me.FV_INDITAS = "" Then
MsgBox "Valamit nem rögzítettél, nézd át az adatokat!", vbCritical + vbOKOnly, "Hiányzó adatok"
Exit Sub
ElseIf myDate <> DLookup("[Valid]", "[Illetmeny]", "[Ill_AD_AZON]='" & TempVars!adoazon & "'") Then
     MsgBox "Illetmény rögzítése kötelezõ!", vbOKOnly + vbCritical, "Figyelmeztetés"
         Exit Sub
    ElseIf Me.FV_UTMOD <> 2 And Me.FV_FG_JIDOTARTAM = 1 Then
         MsgBox "Folyószámla nem lett rögzíte, többi adat mentve! (Határozatlan)", vbOKOnly + vbInformation, "Információ"
            DoCmd.Close acForm, "Foglalkoztatas", acSaveNo
            Forms![Main].[fogldone].Visible = True
            Forms![Main].[jobDone].Visible = True
            Forms![Main].[jobDone].Value = True
    ElseIf Me.FV_UTMOD <> 2 And Me.FV_FG_JIDOTARTAM <> 1 Then
        MsgBox "Folyószámla  és határozott rögzíte!" _
        , vbOKOnly + vbInformation, "Információ"
            DoCmd.Close acForm, "Foglalkoztatas", acSaveNo
            Forms![Main].[fogldone].Visible = True
            Forms![Main].[jobDone].Visible = True
            Forms![Main].[jobDone].Value = True
    ElseIf Me.FV_UTMOD = 2 And Me.FV_FG_JIDOTARTAM <> 1 Then
      MsgBox "Minden szükséges adatt mentve! (Határozott)" _
        , vbOKOnly + vbInformation, "Információ"
            DoCmd.Close acForm, "Foglalkoztatas", acSaveNo
            Forms![Main].[fogldone].Visible = True
            Forms![Main].[jobDone].Visible = True
            Forms![Main].[jobDone].Value = True
    ElseIf Me.FV_UTMOD = 2 And Me.FV_FG_JIDOTARTAM = 1 Then
     MsgBox "Minden szükséges adatt mentve! (Határozatlan)" _
        , vbOKOnly + vbInformation, "Információ"
            DoCmd.Close acForm, "Foglalkoztatas", acSaveNo
            Forms![Main].[fogldone].Visible = True
            Forms![Main].[jobDone].Visible = True
            Forms![Main].[jobDone].Value = True
            Else
            
            Exit Sub
            
            End If
 
    
    
    
    
    

End Sub

Private Sub Szervek_afterUpdate()
Dim sureSzerv
Dim kivalSzerv
kivalSzerv = Mid(Me.Szervek, 1, 6)
If kivalSzerv = "350002" Or kivalSzerv = "350003" Then
    Me.Szerv_Data = Me.Szervek
    Me.Egyedi_SzerV_AZON = Me.Szerv_Data.Column(0)
    Me.FV_SZERVEGYSEG = Me.Szerv_Data.Column(1)
    Me.SZERV_NEV = Me.Szerv_Data.Column(2)
    Me.TemaNev = Me.Szerv_Data.Column(3)
    Me.FV_FKSZLA = Me.Szerv_Data.Column(9)
    Me.FV_TEMASZAM = Me.Szerv_Data.Column(6)
Else
sureSzerv = MsgBox("Valóban erre a szervezeti egységre akarod rögzíteni a munkavállaót?", vbYesNo + vbQuestion, "Szervezeti egyég rögzítése ")
    If sureSzerv = vbNo Then
    Exit Sub
    Else
    Me.Szerv_Data = Me.Szervek
    Me.Egyedi_SzerV_AZON = Me.Szerv_Data.Column(0)
    Me.FV_SZERVEGYSEG = Me.Szerv_Data.Column(1)
    Me.SZERV_NEV = Me.Szerv_Data.Column(2)
    Me.TemaNev = Me.Szerv_Data.Column(3)
    Me.FV_FKSZLA = Me.Szerv_Data.Column(9)
    Me.FV_TEMASZAM = Me.Szerv_Data.Column(6)
    End If
End If
End Sub





