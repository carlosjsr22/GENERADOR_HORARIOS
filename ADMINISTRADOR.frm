VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ADMINISTRADOR 
   Caption         =   "ADMINISTRADOR"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   OleObjectBlob   =   "ADMINISTRADOR.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "ADMINISTRADOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton2_Click()
    GENERARCUENTA.Show
    ADMINISTRADOR.Hide
End Sub

Private Sub CommandButton3_Click()
    ADMINISTRADOR.Hide
End Sub
