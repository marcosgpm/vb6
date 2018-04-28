VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsValorMonetario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const REAL As String = "REAL"
Const REAIS As String = "REAIS"
Const CENTAVO As String = "CENTAVO"
Const CENTAVOS As String = "CENTAVOS"
Const MIL As String = "MIL"
Const MILHAO As String = "MILHAO"
Const MILHOES As String = "MILHOES"

Public Function RetornarPorExtenso(pstrNumero As String) As String

   Dim strTexto As String
   Dim lngValorInteiro As Long
   Dim intValorDecimal As Integer
   
   lngValorInteiro = RetornarValorInteiro(pstrNumero)
   intValorDecimal = RetornarValorCentavos(pstrNumero)
   
   Select Case lngValorInteiro
      Case "1": strTexto = "UM"
      Case "2": strTexto = "DOIS"
      Case "3": strTexto = "TRES"
      Case "4": strTexto = "QUATRO"
      Case "5": strTexto = "CINCO"
      Case "6": strTexto = "SEIS"
      Case "7": strTexto = "SETE"
      Case "8": strTexto = "OITO"
      Case "9": strTexto = "NOVE"
      Case "10": strTexto = "DEZ"
      Case "11": strTexto = "ONZE"
      Case "12": strTexto = "DOZE"
      Case "13": strTexto = "TREZE"
      Case "14": strTexto = "QUATORZE"
      Case "15": strTexto = "QUINZE"
      Case "16": strTexto = "DEZESSEIS"
      Case "17": strTexto = "DEZESSETE"
      Case "18": strTexto = "DEZOITO"
      Case "19": strTexto = "DEZENOVE"
      Case "20": strTexto = "VINTE"
      Case "30": strTexto = "TRINTA"
      Case "40": strTexto = "QUARENTA"
      Case "50": strTexto = "CINQUENTA"
      Case "60": strTexto = "SESSENTA"
      Case "70": strTexto = "SETENTA"
      Case "80": strTexto = "OITENTA"
      Case "90": strTexto = "NOVENTA"
      
   End Select

      
   RetornarPorExtenso = strTexto & " " & "REAL"

End Function

Private Function RetornarValorCentavos(pstrNumero As String) As Integer
   If InStr(1, pstrNumero, ",") > 0 Then
      RetornarValorCentavos = Right(pstrNumero, 2)
   End If
End Function

Private Function RetornarValorInteiro(pstrNumero As String) As Long
   RetornarValorInteiro = pstrNumero
   If InStr(1, pstrNumero, ",") > 0 Then
      RetornarValorInteiro = Left(pstrNumero, (InStr(1, pstrNumero, ",") - 1))
   End If
End Function
