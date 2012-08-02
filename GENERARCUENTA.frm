VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GENERARCUENTA 
   Caption         =   "CREAR CUENTA PRACTICANTE"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   OleObjectBlob   =   "GENERARCUENTA.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "GENERARCUENTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
Dim n As Integer
Hoja1.Select
n = 1


Do Until Hoja1.Cells(n, 6).Value = ComboBox1.Value Or Hoja1.Cells(n, 6).Value = "" Or Hoja1.Cells(n, 6).Value = CStr(ComboBox1.Value)
n = n + 1
Loop

TextBox1.Value = Hoja1.Cells(n, 1).Value
TextBox2.Value = Hoja1.Cells(n, 2).Value
TextBox3.Value = Hoja1.Cells(n, 3).Value
TextBox4.Value = Hoja1.Cells(n, 4).Value
TextBox5.Value = Hoja1.Cells(n, 5).Value
TextBox6.Value = Hoja1.Cells(n, 6).Value
TextBox7.Value = Hoja1.Cells(n, 7).Value

End Sub

Private Sub CommandButton2_Click()
    GENERARCUENTA.Hide
End Sub

Private Sub CommandButton4_Click()
    Dim n As Integer
      Dim sh As Worksheet
    Dim Nombres As String
    
    n = 1
    
    Do Until Hoja1.Cells(n, 6).Value = "" Or Hoja1.Cells(n, 6).Value = CStr(TextBox6)
        n = n + 1
    Loop
    
    If Hoja1.Cells(n, 6).Value = CStr(TextBox6) Then
    
    MsgBox "Usuario ya existe"
    
    Else
    
    
    If (TextBox1.Value <> "" And TextBox2.Value <> "" And TextBox3.Value <> "" And TextBox4.Value <> "" And TextBox5.Value <> "" And TextBox6.Value <> "" And TextBox7.Value <> "") Then
    
    Hoja1.Select
    n = 2
    Do Until Hoja1.Cells(n, 1).Value = ""
        n = n + 1
    Loop

    Cells(n, 1) = TextBox1.Text
    Cells(n, 2) = TextBox2.Text
    Cells(n, 3) = TextBox3.Text
    Cells(n, 4) = TextBox4.Text
    Cells(n, 5) = TextBox5.Text
    Cells(n, 6) = TextBox6.Text
    Cells(n, 7) = TextBox7.Text
    If OptionButton1.Value = True Then
    Cells(n, 8) = OptionButton1.Caption
    Else
    Cells(n, 8) = OptionButton2.Caption
    End If
    
    
A = TextBox6.Value
 Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Select
    Sheets(Sheets.Count).Name = A

    Sheets("Administrador").Select
    Range("A1:A16").Select
    Selection.Copy
    
    Sheets(A).Select
    ActiveSheet.Paste
    
    Sheets("Administrador").Select
    Range("B1:F1").Select
    Application.CutCopyMode = False
    
    Selection.Copy
    Sheets(Sheets.Count).Select
    Range("B1").Select
    ActiveSheet.Paste
    
    Sheets("Administrador").Select
    Application.CutCopyMode = False
    Sheets("hoja1").Select
    
    

    TextBox1 = ""
    TextBox2 = ""
    TextBox3 = ""
    TextBox4 = ""
    TextBox5 = ""
    TextBox6 = ""
    TextBox7 = ""
    OptionButton1.Value = False
    OptionButton2.Value = False
    
    
    Me.ComboBox1.Clear
    Set sh = Sheets("hoja1")
    
    Nombres = "#"
    For I = 2 To sh.Cells.SpecialCells(xlCellTypeLastCell).Row
        Me.ComboBox1.AddItem sh.Cells(I, 6)
        If Me.ComboBox1.Text = "" Then
            Me.ComboBox1.Text = sh.Cells(I, 6)
            Nombres = Nombres & sh.Cells(I, 6) & "#"
        End If
    Next I
    
    Else
   
    MsgBox "Completar todos los datos"
    
    End If
    End If
End Sub

Private Sub CommandButton5_Click()
Dim n As Integer


If ComboBox1 = TextBox6.Value Then
    
Hoja1.Select
n = 1

Do Until Hoja1.Cells(n, 6).Value = ComboBox1.Value Or Hoja1.Cells(n, 6).Value = "" Or Hoja1.Cells(n, 6).Value = CStr(ComboBox1.Value)
n = n + 1
Loop

Hoja1.Cells(n, 1).Value = CStr(TextBox1.Value)
Hoja1.Cells(n, 2).Value = CStr(TextBox2.Value)
Hoja1.Cells(n, 3).Value = CStr(TextBox3.Value)
Hoja1.Cells(n, 4).Value = CStr(TextBox4.Value)
Hoja1.Cells(n, 5).Value = CStr(TextBox5.Value)
Hoja1.Cells(n, 6).Value = CStr(TextBox6.Value)
Hoja1.Cells(n, 7).Value = CStr(TextBox7.Value)

ComboBox1.Clear

n = 2

Do Until Hoja1.Cells(n, 6).Value = ""
ComboBox1.AddItem (Hoja1.Cells(n, 6).Value)
n = n + 1
Loop

Else

MsgBox "Usuario ya existe"
    
End If

End Sub

Private Sub CommandButton6_Click()
Dim n As Integer
Dim A As String
Dim Hoja As Object

Hoja1.Select
n = 1

Do Until Hoja1.Cells(n, 6).Value = ComboBox1.Value Or Hoja1.Cells(n, 6).Value = "" Or Hoja1.Cells(n, 6).Value = CStr(ComboBox1.Value)
n = n + 1
Loop

Range("A" & n).Select
Selection.EntireRow.Delete


A = ComboBox1.Value
    Sheets(A).Select
    ActiveWindow.SelectedSheets.Delete

n = 2

ComboBox1.Clear

Do Until Hoja1.Cells(n, 6).Value = ""
ComboBox1.AddItem (Hoja1.Cells(n, 6).Value)
n = n + 1
Loop
End Sub

Private Sub UserForm_Initialize()
    Dim sh As Worksheet
    Dim Nombres As String
    
    Me.ComboBox1.Clear
    Set sh = Sheets("hoja1")
    
    Nombres = "#"
    For I = 2 To sh.Cells.SpecialCells(xlCellTypeLastCell).Row
        Me.ComboBox1.AddItem sh.Cells(I, 6)
        If Me.ComboBox1.Text = "" Then
            Me.ComboBox1.Text = sh.Cells(I, 6)
            Nombres = Nombres & sh.Cells(I, 6) & "#"
        End If
    Next I
End Sub
