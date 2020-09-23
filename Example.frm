VERSION 5.00
Begin VB.Form Example 
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   2085
   ClientTop       =   1725
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   8625
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1290
      TabIndex        =   7
      Top             =   1590
      Width           =   2205
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1290
      TabIndex        =   6
      Top             =   1200
      Width           =   2205
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1290
      TabIndex        =   5
      Top             =   810
      Width           =   2205
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1290
      TabIndex        =   4
      Top             =   450
      Width           =   2205
   End
   Begin Project1.cLabel cLabel1 
      Height          =   225
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   397
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Name"
      id              =   "1"
   End
   Begin Project1.cLabel cLabel1 
      Height          =   225
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   855
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   397
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      Caption         =   "Address"
      id              =   "2"
   End
   Begin Project1.cLabel cLabel1 
      Height          =   225
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   1215
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   397
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      Caption         =   "Telephone"
      id              =   "3"
   End
   Begin Project1.cLabel cLabel1 
      Height          =   225
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   1590
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   397
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      Caption         =   "Country"
      id              =   "4"
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   270
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   300
      Y1              =   1290
      Y2              =   1290
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   270
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   270
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   1590
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   120
      Y1              =   570
      Y2              =   2790
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Example.frx":0000
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   1590
      TabIndex        =   8
      Top             =   2550
      Width           =   2835
   End
End
Attribute VB_Name = "Example"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Use the following declarations in a module of your program
Public dbCustom As ADODB.Connection
Public tmpAdo As recordset

Private Sub Form_Load()
    
    ' Use the following code at the beginning of your program ONCE
    '-------------------------------------------------------------
    Set dbCustom = New ADODB.Connection
    dbCustom.CursorLocation = adUseClient
    dbCustom.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\mscustom.mdb" & ";Persist Security Info=False"
    
    Set tmpAdo = New recordset
    tmpAdo.Open "msCaptions", dbCustom, adOpenStatic, adLockOptimistic
    '-------------------------------------------------------------
    
    
    
    ' Use the following code at the load of every form that uses ctlabel
    ' This all the code necessary to add to your forms.
    ' Of course make sure that the ctlabel.id is unique for every label in your forms
    ' or use the same id for labels with the same caption that you want to change globally.
        
    For i = 0 To cLabel1.Count - 1
      cLabel1(i).recordset = tmpAdo
    Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  ' Following code is not really necessary.
  ' If you use it then do it when you exit your program
  ' I use it here as good programming practice.
  
  tmpAdo.Close
  Set tmpAdo = Nothing
  
  dbCustom.Close
  Set dbCustom = Nothing
  
End Sub
