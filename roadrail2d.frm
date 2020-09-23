VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000C000&
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'I made this routine after I decided to start working my
'way towards a civilization type of game. The routine checks
'the area for layed roads/railroads and connects to them.
'
'If you use this routine then please credit me.
'
'
'
Dim x1(1956)  As Integer      ' Coordinaten x
Dim y1(1956)    As Integer        ' ,,          y
Dim blob(1956)   As Integer     ' Data Grid
Dim grids   As Integer         ' Aantal grids (156 = 800*600)
Dim wd   As Integer            ' blokjes x in de breedte
Dim hg     As Integer          ' ,,              hoogte
Dim ic     As Integer          ' Geselecteerde blokje
Dim xpos     As Integer        ' xpositie op het scherm
Dim ypos      As Integer       ' ypositie op het scherm
Dim hig   As Integer
Dim wid  As Integer

Private Sub Form_Load()
ScaleMode = 3
hig = 32
wid = 32
Form1.Caption = "Routine for the building of a road system - By Rudy van Etten in 1999 (Press the mouse to see it in work)"
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X < wd * wid And Y < hg * hig Then
    xpos = Fix(X / wid) + 1                 'Zet de xlokatie om -1
    ypos = Fix(Y / hig) + 1                 ' ,,    y
    ic = xpos + (ypos - 1) * wd               'Zet de waarde om naar een nummer
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X < wd * wid And Y < hg * hig Then
    If Button = 1 Then blob(ic) = True:  Line (x1(ic), y1(ic))-(x1(ic) + wid, y1(ic) + hig), RGB(0, 155, 0), BF: Call drawline(ic)
    If Button = 2 Then blob(ic) = False: Form1.Cls: Call refr
    Else
    If Button = 2 Then End
End If
End Sub

Private Sub drawline(n)
On Error Resume Next
ScaleMode = 3
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
Dim e As Integer
Dim f As Integer
Dim g As Integer
Dim h As Integer

a = True
b = True
c = True
d = True
e = True
f = True
g = True
h = True

If xpos = 1 Then a = False: d = False: f = False
If xpos = wd Then c = False: e = False: h = False
If ypos = 1 Then a = False: b = False: c = False
If ypos = hg Then f = False: g = False: h = False

If a = True Then a = blob(n - (wd + 1))
If b = True Then b = blob(n - wd)
If c = True Then c = blob(n - (wd - 1))
If d = True Then d = blob(n - 1)
If e = True Then e = blob(n + 1)
If f = True Then f = blob(n + (wd - 1))
If g = True Then g = blob(n + wd)
If h = True Then h = blob(n + (wd + 1))

If a = True Then Call connect(x1(n), y1(n), x1(n - (wd + 1)), y1(n - (wd + 1)))
If b = True Then Call connect(x1(n), y1(n), x1(n - wd), y1(n - wd))
If c = True Then Call connect(x1(n), y1(n), x1(n - (wd - 1)), y1(n - (wd - 1)))
If d = True Then Call connect(x1(n), y1(n), x1(n - 1), y1(n - 1))
If e = True Then Call connect(x1(n), y1(n), x1(n + 1), y1(n + 1))
If f = True Then Call connect(x1(n), y1(n), x1(n + (wd - 1)), y1(n + (wd - 1)))
If g = True Then Call connect(x1(n), y1(n), x1(n + wd), y1(n + wd))
If h = True Then Call connect(x1(n), y1(n), x1(n + (wd + 1)), y1(n + (wd + 1)))

End Sub

Private Sub connect(aa, bb, cc, dd)
Line (aa + (wid / 2), bb + (hig / 2))-(cc + (wid / 2), dd + (hig / 2)), RGB(0, 0, 0)
Line (aa + (wid / 2) - 1, bb + (hig / 2 - 1))-(cc + (wid / 2) - 1, dd + (hig / 2) - 1), RGB(0, 0, 0)
Line (aa + (wid / 2) + 1, bb + (hig / 2) + 1)-(cc + (wid / 2) + 1, dd + (hig / 2) + 1), RGB(0, 0, 0)
End Sub


Private Sub Form_Paint()
Call Form_Resize
End Sub

Private Sub Form_Resize()
wd = Fix(800 \ wid)      ' Bereken de breedte
hg = (600 - wid) \ hig    ' Bereken de hoogte
'Teken het grid op het scherm
For b = 0 To hg - 1
    For a = 0 To wd - 1
'        Line (a * wid, b * hig)-((a * wid) + wid - 1, (b * hig) + hig - 1), RGB(255, 255, 255), B
        kt = kt + 1
        x1(kt) = a * wid
        y1(kt) = b * hig
    Next
Next
grids = kt  'Zet het aantal nummers neer

For i = 0 To grids
    blob(i) = False
Next
End Sub

Private Sub refr()
For i = 1 To grids
    If blob(i) = True Then Call drawline(i)
Next
End Sub



