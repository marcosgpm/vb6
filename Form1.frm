VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

   Dim objValorMonetario As New clsValorMonetario
   Dim valor As Double
   Dim i As Integer
   
   valor = 0
   Open App.Path & "\valores.txt" For Output As #5
   For i = 0 To 9999
      Print #5, objValorMonetario.RetornarPorExtenso(valor)
      valor = valor + 0.01
   Next i
   Close #5
   
   MsgBox "Valores salvos no arquivo " & App.Path & "\valores.txt"
   
   Unload Me

End Sub


