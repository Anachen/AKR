VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'A programrol nevu urlap tartalma az utolso frissitesekrol szolo informáciok, rovid ismertetese.

Private Sub Form_Load()
Dim Verz As String
Dim inver
Dim imdate
'A legujabb verzio es a hozza tartozo frissitesi informaciok behivasa a tablazatbol
Verz = DMax("Verzio", "Version")
inver = DMax("INTF_VER", "Version")
imdate = DMax("UNIQ_DATA_IMPORT", "Version")
Me.Verzio = Verz
Me.InterVer = inver
Me.importDate = imdate


End Sub

Private Sub Ok_Click()
'az urlap bezarasa
DoCmd.Close
End Sub
