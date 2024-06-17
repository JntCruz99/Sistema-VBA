VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWbName As String
    Dim xlWbPath As String

    On Error Resume Next
    
    Set xlApp = CreateObject("Excel.Application")
    
    xlWbName = "sistema.xlsm"
    xlWbPath = App.Path
    
    Me.Hide
    Set xlWb = xlApp.workbooks.open(xlWbPath & "\" & xlWbName)

    Set xlWb = Nothing
    Unload Me

End Sub
