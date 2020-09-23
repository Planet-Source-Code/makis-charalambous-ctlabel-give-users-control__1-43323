VERSION 5.00
Begin VB.Form cform 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Custom Text"
   ClientHeight    =   1170
   ClientLeft      =   4755
   ClientTop       =   6090
   ClientWidth     =   4170
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3300
      TabIndex        =   5
      Top             =   810
      Width           =   855
   End
   Begin VB.CheckBox chkUnder 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Underline"
      Height          =   225
      Left            =   930
      TabIndex        =   4
      Top             =   810
      Width           =   1275
   End
   Begin VB.CheckBox chkItalic 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Italic"
      Height          =   225
      Left            =   930
      TabIndex        =   3
      Top             =   600
      Width           =   795
   End
   Begin VB.CheckBox chkBold 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Bold"
      Height          =   225
      Left            =   930
      TabIndex        =   2
      Top             =   360
      Width           =   705
   End
   Begin VB.TextBox cText 
      Height          =   285
      Left            =   930
      TabIndex        =   0
      Top             =   30
      Width           =   3225
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Text"
      Height          =   255
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   795
   End
End
Attribute VB_Name = "cform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
  Me.Hide
End Sub
