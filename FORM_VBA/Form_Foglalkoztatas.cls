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
MsgBox "Els� jogviszony eset�ny k�rem az iskolai tanulm�nyok v�g�t megadni!", _
vbOKOnly + vbExclamation, "Tov�bbi adat sz�ks�ges!"
Me.iskvegeTxt.Visible = True
Me.FG_ISKVEGE.Visible = True
End If




End Sub

Private Sub Folyoszamla_Click()

Dim myDate
myDate = DateValue("31/12/3999")
If myDate = DLookup("[Valid]", "[SZAMLA]", "[SZAMLA_AD_AZON]='" & TempVars!adoazon & _
"' and [Exported] = false") Then
MsgBox "A munkav�llal�nak m�r van r�gz�tve Foly�sz�mla.", _
vbOKOnly + vbCritical, "Figyelmeztet�s"
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
MsgBox "K�rlek �ll�tsd be a hat�rozott id�t", _
vbOKOnly + vbExclamation, "Tov�bbi adat sz�ks�ges!"
Me.Hat.Visible = True
Me.Hat.SetFocus
ElseIf idotart <> 1 Then
MsgBox "Csak hat�rozott �s hat�rozatlan �ll�that� be!", _
vbOKOnly + vbCritical, "Figyelmeztet�s"
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
MsgBox "Teljes munkaid�s jogviszony eset�n az alap�rtelmezett heti 40 �ra be�ll�tva.", _
vbOKOnly + vbInformation, "Inform�ci�"
Me.FV_TEHOSZ = 40
Me.FV_FSTATUSZ.SetFocus
Else
MsgBox "R�szmunkaid�s eset�n k�rem a munkaid� sz�m�t�s t�pus�t �s t�nyleges hossz�t megadni!", _
vbOKOnly + vbExclamation, "Tov�bbi adat sz�ks�ges!"
Me.FV_MTIP.SetFocus
End If

End Sub

Private Sub FV_FSTATUSZ_Click()
If Me.FV_FSTATUSZ = "03" Then
    MsgBox "A Feladat ell�t�s�nak megnevez�s�t adja meg!", _
    vbOKOnly + vbExclamation, "Tov�bbi adat sz�ks�ges!"
    Me.Felad_ellat.Visible = True
    Me.Felad_ellat.SetFocus
    ElseIf Me.FV_FSTATUSZ <> "01" Then
    
    MsgBox "Adatokat r�gz�thet�ek, de csak '�ll�s terh�re' vagy 'Feladat ell�t�s�ra'" _
    & "opci�k eset�re  lehet csak import�lni!", _
    vbOKOnly + vbCritical, "Figyelmeztet�s"

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
MsgBox "A Kiv�lasztott jelminta nem egyezik a megadott heti �rasz�mmal! K�rem ellen�rizze!" & vbNewLine & _
"A " & Me.FV_TEHOSZ & " nem egyezik meg a jelminta �rt�kkel: " & jelOraErtek, _
vbOKOnly + vbInformation, "Inform�ci�"
Me.FV_JELMINTA = ""
Exit Sub
Else
MsgBox "A Jelminta minta automatikusan kit�ltve. K�rem ellen�rizze!" & vbNewLine & _
"A jelenl�ti kezdete a bel�p�si h�nap els� napja." & vbNewLine & _
"a jelenl�ti napja a h�nap els� napj�nak napt�ri napja (H�ft�-Vas�rnap)", _
vbOKOnly + vbInformation, "Inform�ci�"
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
MsgBox "Az adatfelv�tel megt�rt�nik, de import�l�sra az adat nem lesz alkalmas!", _
vbOKOnly + vbCritical, "Figyelmeztet�s"
End If


End Sub

Private Sub FV_MTIP_Click()
If Me.FV_MTIP <> 2 Then
MsgBox "A T�nyleges hosszt �s t�rv�nyes hosszt k�rem megadni, �ra eset�n a havi �rasz�mot is!", _
vbOKOnly + vbExclamation, "Tov�bbi adat sz�ks�ges!"
Me.FV_TEHOSZ.SetFocus
Me.FV_FG_HAVIORA.Visible = True

End If


End Sub

Private Sub FV_UTMOD_Click()
If Me.FV_UTMOD = 2 Then
MsgBox "�ll�tsa be a foly�sz�mlasz�mot!", _
vbOKOnly + vbExclamation, "Tov�bbi adat sz�ks�ges!"
Me.Folyoszamla.Visible = True
Me.Folyoszamla.SetFocus
Else
MsgBox "Ha nem nem foly�sz�ml�ra utal�s van, akkor csak a 'Lakc�mre', vagy 'Lakc�mre, de post�n marad' import�lhat�!", _
vbOKOnly + vbCritical, "Figyelmeztet�s"
End If

End Sub

Private Sub Hat_Click()
Dim myDate
myDate = DateValue("12/31/3999")
If Me.FV_FSTATUSZ = "03" And IsNull(Me.Felad_ellat) = True Then
MsgBox "Feladat ell�t�s�t v�lasztott�l, de nem adt�l meg feladatot!", vbCritical + vbOKOnly _
, "Hi�nyz� adat"
Exit Sub
ElseIf myDate = DLookup("[Valid]", "[HATFOG]", "[ADO_HATFOG]='" & TempVars!adoazon & _
"'and [Exported] = False") Then
MsgBox "A munkav�llal�nak m�r van r�gz�tve hat�rozott id� r�gz�tve!", _
vbOKOnly + vbCritical, "Figyelmeztet�s"
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
MsgBox "A munkav�llal�nak m�r van Illetm�nyadata. K�rem ellen�rizze!", _
vbOKOnly + vbCritical, "Figyelmeztet�s"
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
MsgBox "Jelenleg �gy t�nik nem r�gz�tett�l illetm�nyadatot, k�rlek ellen�rizd!", _
vbCritical + vbOKOnly, "Hi�nyz� illetm�ny"
Exit Sub
ElseIf Me.Felad_ellat <> DLookup("[FV_FG_HATFEL]", "[HATFOG]", _
"[ADO_HATFOG]='" & TempVars!adoazon & "' and [Exported] = False") Then
MsgBox "A Foglalkozat�si lapon megadott feladat nem egyezik az �rv�nyes adattal, k�rlek friss�tsd!" _
, vbCritical + vbOKOnly, "Elt�r� adatok"
ElseIf Me.FV_FSTATUSZ = "03" And IsNull(Me.Felad_ellat) = True Then
MsgBox "Feladat ell�t�s�t v�lasztott�l, de nem adt�l meg feladatot!", vbCritical + vbOKOnly _
, "Hi�nyz� adat"
Exit Sub
ElseIf (Me.FV_FG_JIDOTARTAM <> 2 And Me.HatDone = True) _
Or (Me.FV_FG_JIDOTARTAM = 2 And Me.HatDone = False) Then
MsgBox "Hat�rozott id�s adat van, de te hat�rozatlant akarsz felvinni!" _
, vbCritical + vbOKOnly, "T�puselt�r�si hiba"
Exit Sub
ElseIf (Me.FV_UTMOD <> 2 And Me.SzamlaDone = True) _
Or (Me.FV_UTMOD = 2 And Me.SzamlaDone = False) Then
MsgBox "Sz�mlaadat van, de te m�s t�pus� adatot akarsz felvinni!" _
, vbCritical + vbOKOnly, "T�puselt�r�si hiba"
Exit Sub
ElseIf Me.FG_JOGVKEZDETE = "" Or Me.FV_SZERVEGYSEG = "" Or _
Me.FV_TEMASZAM = "" Or Me.FV_FKSZLA = "" Or Me.FV_BERTABLA = "" Or Me.FV_FG_BERTABLA = "" Or _
IsNull(Me.FV_COFOG) = True Or Me.FG_MLTIP = "" Or Me.FV_MUNKAREND = "" Or Me.FV_MIDOBEOSZT = "" Or _
Me.FV_FMUSZAKREND = "" Or Me.FV_FMIDOJEL = "" Or Me.FV_FSTATUSZ = "" Or Me.FV_MTIP = "" Or _
IsNull(Me.FV_TEHOSZ) = True Or Me.FV_FG_NEV = "" Or Me.FV_FG_JIDOTARTAM = "" Or Me.FV_UTMOD = "" Or _
IsNull(Me.FV_FG_KULCSSZAM) = True Or IsNull(Me.FV_FEOR) = True Or Me.FV_MKOR = "" Or Me.FV_JELMINTA = "" Or _
Me.FV_JELKEZD = "" Or Me.FV_INDITAS = "" Then
MsgBox "Valamit nem r�gz�tett�l, n�zd �t az adatokat!", vbCritical + vbOKOnly, "Hi�nyz� adatok"
Exit Sub
ElseIf myDate <> DLookup("[Valid]", "[Illetmeny]", "[Ill_AD_AZON]='" & TempVars!adoazon & "'") Then
     MsgBox "Illetm�ny r�gz�t�se k�telez�!", vbOKOnly + vbCritical, "Figyelmeztet�s"
         Exit Sub
    ElseIf Me.FV_UTMOD <> 2 And Me.FV_FG_JIDOTARTAM = 1 Then
         MsgBox "Foly�sz�mla nem lett r�gz�te, t�bbi adat mentve! (Hat�rozatlan)", vbOKOnly + vbInformation, "Inform�ci�"
            DoCmd.Close acForm, "Foglalkoztatas", acSaveNo
            Forms![Main].[fogldone].Visible = True
            Forms![Main].[jobDone].Visible = True
            Forms![Main].[jobDone].Value = True
    ElseIf Me.FV_UTMOD <> 2 And Me.FV_FG_JIDOTARTAM <> 1 Then
        MsgBox "Foly�sz�mla  �s hat�rozott r�gz�te!" _
        , vbOKOnly + vbInformation, "Inform�ci�"
            DoCmd.Close acForm, "Foglalkoztatas", acSaveNo
            Forms![Main].[fogldone].Visible = True
            Forms![Main].[jobDone].Visible = True
            Forms![Main].[jobDone].Value = True
    ElseIf Me.FV_UTMOD = 2 And Me.FV_FG_JIDOTARTAM <> 1 Then
      MsgBox "Minden sz�ks�ges adatt mentve! (Hat�rozott)" _
        , vbOKOnly + vbInformation, "Inform�ci�"
            DoCmd.Close acForm, "Foglalkoztatas", acSaveNo
            Forms![Main].[fogldone].Visible = True
            Forms![Main].[jobDone].Visible = True
            Forms![Main].[jobDone].Value = True
    ElseIf Me.FV_UTMOD = 2 And Me.FV_FG_JIDOTARTAM = 1 Then
     MsgBox "Minden sz�ks�ges adatt mentve! (Hat�rozatlan)" _
        , vbOKOnly + vbInformation, "Inform�ci�"
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
sureSzerv = MsgBox("Val�ban erre a szervezeti egys�gre akarod r�gz�teni a munkav�lla�t?", vbYesNo + vbQuestion, "Szervezeti egy�g r�gz�t�se ")
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





