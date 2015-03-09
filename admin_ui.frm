VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} admin_ui 
   Caption         =   "Administration"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6045
   OleObjectBlob   =   "admin_ui.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "admin_ui"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btn_purge_wb_Click()
Dim a As Variant
Dim answer As Variant

Set a = Worksheets("calculator")

answer = MsgBox("Are you sure you want to purge the workbook?" & vbNewLine & "All recipe workbook will be deleted!", vbYesNo + vbQuestion, "Purge")


If answer = vbYes Then
    Application.DisplayAlerts = False
    For Each ws In Sheets
        If Not ws.Name = "ingredient" And Not ws.Name = "calculator" And Not ws.Name = "tmp" Then
        ws.Delete
        End If
    Next
    Application.DisplayAlerts = True
End If
MsgBox ("All recipe workbook deleted!")

admin_ui.Hide
Unload Me
End Sub

Private Sub Btn_sortingred_Click()
Call Module1.subfunc_ingred_sort
MsgBox ("Ingredient list sorted!")

admin_ui.Hide
Unload Me
End Sub

Private Sub Btn_close_Click()
'exit btn click esemény lekezelése, form elrejtése, unload
admin_ui.Hide
Unload Me
End Sub


