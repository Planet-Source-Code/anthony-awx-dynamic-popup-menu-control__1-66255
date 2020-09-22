VERSION 5.00
Begin VB.Form frmDynamicMenuSample 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dynamic Menu Control Sample"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Show Sample Menu 1 (long)"
      Height          =   795
      Left            =   435
      TabIndex        =   1
      Top             =   1245
      Width           =   2400
   End
   Begin Project1.DynamicPopupMenu dpMenu1 
      Left            =   2715
      Top             =   2115
      _ExtentX        =   847
      _ExtentY        =   503
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Sample Menu 1 (short)"
      Height          =   795
      Left            =   435
      TabIndex        =   0
      Top             =   345
      Width           =   2400
   End
End
Attribute VB_Name = "frmDynamicMenuSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    ' SHOW A SAMPLE DYNAMIC MENU AND DISPLAY WHAT WAS CHOSEN
    Dim SelectedChoice As String
    SelectedChoice = dpMenu1.Popup("item1, item2, item3, item4, -, item5, item6, item7")
    If SelectedChoice <> "" Then MsgBox SelectedChoice
    
End Sub

Private Sub Command2_Click()

    ' SHOW A SAMPLE DYNAMIC MENU AND DISPLAY WHAT WAS CHOSEN
    Dim SelectedChoice As String
    SelectedChoice = dpMenu1.Popup("item1, item2, item3, item4, -, item5, item6, item7,-,item8,item9,item10,item11,item12,item13,item14,item15,item16,-,item17,item18,item19,item20")
    If SelectedChoice <> "" Then MsgBox SelectedChoice

End Sub
