VERSION 5.00
Begin VB.UserControl cLabel 
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2895
   ScaleHeight     =   615
   ScaleWidth      =   2895
   Begin VB.Label Label1 
      BackColor       =   &H00D8E9EC&
      Caption         =   "cLabel1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2385
   End
End
Attribute VB_Name = "cLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private ctrlDown As Integer
Private isOpen As Integer
Private defaultCaption As Variant
Private tmpAdo As recordset
Private m_def_recordset As recordset

Private tmpNewText As String
Private m_id As Variant
Const m_def_id = 0

'Property Variables:
'Event Declarations:
Event Click() 'MappingInfo=Label1,Label1,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=Label1,Label1,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Label1,Label1,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Label1,Label1,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Label1,Label1,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event error()
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Label1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Label1.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Label1.ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = Label1.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Label1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    PropertyChanged "Font"
End Property
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = Label1.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    Label1.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = Label1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    Label1.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    Label1.Refresh
End Sub
Private Sub Label1_DblClick()

    On Error Resume Next
    
    If Not ctrlDown Then Exit Sub ' Activate with doublie AND control button pressed
    If Not isOpen Then Exit Sub ' if no recordset defined then exit
    
    'tmpNewText = InputBox("Enternew Name", , (Label1.Caption))
    
    cform.cText.Text = Label1.Caption
    cform.Show vbModal
    tmpNewText = cform.cText
        
'    If Label1.Caption = tmpNewText Then
'      Unload cform  ' Throw away the form after we got our parameters
'      Exit Sub      ' Text not changed, no need to do anything
'    End If
        
    tmpAdo.MoveFirst
    tmpAdo.Find "txtID=" & Val(m_id)
    
    If Not tmpAdo.EOF Then      ' check if id exists
      If tmpNewText = "" Then   ' If no text then delete previous custom text and restore default text
            tmpAdo.Delete
            Label1.Caption = defaultCaption
      Else                      ' else modify previous text
            SetNewCaption
            tmpAdo.Update
      End If
    Else
        tmpAdo.AddNew
        SetNewCaption
        tmpAdo.Update
    End If
    
    Unload cform  ' Throw away the form after we got our parameters

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    PropertyChanged "Caption"
End Property
Public Property Get id() As Variant
    id = m_id
End Property

Public Property Let id(ByVal New_id As Variant)
    m_id = New_id
    PropertyChanged "id"
End Property

Public Property Let recordset(ByVal New_recordset As recordset)
    Set tmpAdo = New_recordset
    isOpen = True
    PropertyChanged "recordset"
End Property


'Initialize Properties for User Control

Private Sub UserControl_InitProperties()
    m_id = m_def_id
    ctrlDown = False
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 2 Then
     ctrlDown = True
   End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
   ctrlDown = False
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Label1.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Label1.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Label1.BackStyle = PropBag.ReadProperty("BackStyle", 0)
    Label1.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Label1.Caption = PropBag.ReadProperty("Caption", "Label1")
    defaultCaption = Label1.Caption
    m_id = PropBag.ReadProperty("id", m_def_id)
    Set tmpAdo = PropBag.ReadProperty("recordset", m_def_recordset)
    
End Sub

Private Sub UserControl_Show()
  If Not isOpen Then Exit Sub
  FetchCaption
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", Label1.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", Label1.Enabled, True)
    Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", Label1.BackStyle, 0)
    Call PropBag.WriteProperty("BorderStyle", Label1.BorderStyle, 0)
    Call PropBag.WriteProperty("Caption", Label1.Caption, "cLabel1")
    Call PropBag.WriteProperty("id", m_id, m_def_id) ' Holds the id for the customized text
    Call PropBag.WriteProperty("recordset", m_recordset, tmpAdo)
    
End Sub

Private Sub UserControl_Resize()
  
  Label1.Move 0, 0, UserControl.Width, UserControl.Height
  
End Sub

Private Sub FetchCaption()
    
    On Error Resume Next
    tmpAdo.MoveFirst
    tmpAdo.Find "txtID=" & Val(m_id)
    If Not tmpAdo.EOF Then
      
      Label1.Caption = tmpAdo.Fields("txtCaption")

      If (tmpAdo.Fields("txtBold") & "") = True Then
        Label1.FontBold = True
      Else
        Label1.FontBold = False
      End If
      
      If (tmpAdo.Fields("txtItalic") & "") = True Then
        Label1.FontItalic = True
      Else
        Label1.FontItalic = False
      End If
     
      If (tmpAdo.Fields("txtUnder") & "") = True Then
        Label1.FontUnderline = True
      Else
        Label1.FontUnderline = False
      End If
      
    End If
    
End Sub
Private Sub SetNewCaption()

            tmpAdo.Fields("txtID") = m_id
            
            tmpAdo.Fields("txtCaption") = tmpNewText
            
            Label1.FontBold = cform.chkBold.Value
            tmpAdo.Fields("txtBold") = Label1.FontBold
            
            Label1.FontItalic = cform.chkItalic.Value
            tmpAdo.Fields("txtItalic") = Label1.FontItalic
            
            Label1.FontUnderline = cform.chkUnder.Value
            tmpAdo.Fields("txtUnder") = Label1.FontUnderline
            Label1.Caption = tmpNewText

End Sub
            
