VERSION 5.00
Begin VB.Form ShapeForm 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3075
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   193
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "ShapeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This is the variable that will keep the memory address of the region
Private hRgn As Long, TransColor As SystemColorConstants

Private Sub Form_Load()
'Set the transparent color to White, create the region and modify the Forms Shape
'with it
    TransColor = vbWhite
    SetRegion
'Show the Shaped Form
    ShapeForm.Show
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Free the used memory by the Region and unload the Shaped Form
    If hRgn Then DeleteObject hRgn
    Unload ShapeForm
End Sub

Private Sub SetRegion()
'Free the memory allocated by the previous Region
    If hRgn Then DeleteObject hRgn
'Scan the Bitmap and remove all transparent pixels from it, creating a new region
    hRgn = GetBitmapRegion(ShapeForm.Picture, TransColor)
'Set the Form's new Region
    SetWindowRgn ShapeForm.hwnd, hRgn, True
End Sub

