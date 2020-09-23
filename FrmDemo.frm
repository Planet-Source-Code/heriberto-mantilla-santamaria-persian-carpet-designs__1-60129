VERSION 5.00
Begin VB.Form frmDemo 
   AutoRedraw      =   -1  'True
   Caption         =   "Persian Carpet Designs - Original concept from Anne Burns"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6735
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "FrmDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   426
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   449
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 Private Der         As Long, Izq  As Long, Abj As Long, k As Long
 Private ColorBorder As Long, Bot  As Long, A   As Long
 
Private Sub Form_DblClick()
 End
End Sub

Private Sub Form_Load()
 Cls
 Show
 Height = 6855
 Width = 6855
 ScaleMode = 3
 AutoRedraw = True
 Izq = 1
 Abj = 1
 Der = ScaleWidth - 1
 Bot = ScaleHeight - 1
 ColorBorder = RGB(0, 0, 0)
 A = 123
 Line (Izq, Abj)-(Der, Abj), ColorBorder
 Line (Izq, Bot)-(Der, Bot), ColorBorder
 Line (Izq, Abj)-(Izq, Bot), ColorBorder
 Line (Der, Abj)-(Der, Bot), ColorBorder
 Call DetermineColor(Izq, Der, Abj, Bot, A)
End Sub

'* Determine the color based on function F.
Private Function DetermineColor(ByVal Izq As Long, ByVal Der As Long, ByVal Abj As Long, ByVal Bot As Long, ByVal A As Long) As Long
 Dim MiddleCol As Long, MiddleRow As Long, C As Long
 
 If (Izq < Der - 1) Then
  C = F(Izq, Der, Abj, Bot, A)
  MiddleCol = (Izq + Der) \ 2
  MiddleRow = (Abj + Bot) \ 2
  Line (Izq + 1, MiddleRow)-(Der - 1, MiddleRow), C
  Line (MiddleCol, Abj + 1)-(MiddleCol, Bot - 1), C
  DetermineColor = DetermineColor(MiddleCol, Der, Abj, MiddleRow, A)
  DetermineColor = DetermineColor(MiddleCol, Der, MiddleRow, Bot, A)
  DetermineColor = DetermineColor(Izq, MiddleCol, Abj, MiddleRow, A)
  DetermineColor = DetermineColor(Izq, MiddleCol, MiddleRow, Bot, A)
  DoEvents
 End If
End Function

'* When b=4, this function takes an average.
Private Function F(ByVal Izq As Long, ByVal Der As Long, ByVal Abj As Long, ByVal Bot As Long, ByVal A As Long) As Long
 Dim P As Long, B As Long
 
 P = Point(Izq, Abj) + Point(Der, Abj) + Point(Izq, Bot) + Point(Der, Bot)
 '* Try values of b = 4 or b = 7.
 B = 4
 F = (P \ B) + A
End Function

Private Sub Form_Unload(Cancel As Integer)
 End
End Sub
