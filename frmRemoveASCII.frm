VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRemoveASCII 
   Caption         =   "Replace ASCII Characters"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5190
   Icon            =   "frmRemoveASCII.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstabContainer 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   5106
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   529
      BackColor       =   -2147483639
      TabCaption(0)   =   "Preset Text"
      TabPicture(0)   =   "frmRemoveASCII.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "chkPresetControlCharacters"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "frameTextFrame"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Custom Text"
      TabPicture(1)   =   "frmRemoveASCII.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "chkCustomControlCharacters"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.CheckBox chkCustomControlCharacters 
         Caption         =   "Remove Control Characters?"
         Height          =   435
         Left            =   2160
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Do ASCII control characters need to be replaced?"
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CheckBox chkPresetControlCharacters 
         Caption         =   "Remove Control Characters?"
         Height          =   435
         Left            =   -72840
         TabIndex        =   3
         ToolTipText     =   "Do ASCII control characters need to be replaced?"
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Frame Frame3 
         Caption         =   "Text"
         Height          =   2295
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1815
         Begin VB.TextBox txtCustomText 
            Height          =   615
            Left            =   120
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "This text will be replaced."
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label9 
            Caption         =   "Replaced Text"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Original Text"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblCustomReplacedText 
            Height          =   495
            Left            =   120
            TabIndex        =   20
            ToolTipText     =   "The final text after replacements"
            Top             =   1680
            Width           =   1575
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Replace"
         Height          =   2295
         Left            =   2040
         TabIndex        =   17
         Top             =   480
         Width           =   2775
         Begin VB.CheckBox chkCustomSpecialCharacters 
            Caption         =   "Remove Special Characters?"
            Height          =   375
            Left            =   120
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Do special characters need to be replaced?"
            Top             =   1200
            Width           =   2535
         End
         Begin VB.TextBox txtCustomReplacements 
            Height          =   615
            Left            =   120
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Enter text here to replace."
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label Label5 
            Caption         =   "Characters to Replace"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   18
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Replace"
         Height          =   2295
         Left            =   -72960
         TabIndex        =   15
         Top             =   480
         Width           =   2775
         Begin VB.TextBox txtPresetReplacements 
            Height          =   615
            Left            =   120
            TabIndex        =   1
            ToolTipText     =   "Enter text here to replace."
            Top             =   480
            Width           =   2535
         End
         Begin VB.CheckBox chkPresetSpecialCharacters 
            Caption         =   "Remove Special Characters?"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            ToolTipText     =   "Do special characters need to be removed?"
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Label3 
            Caption         =   "Characters to Replace"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   16
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame frameTextFrame 
         Caption         =   "Text"
         Height          =   2295
         Left            =   -74880
         TabIndex        =   10
         Top             =   480
         Width           =   1815
         Begin VB.Label lblPresetReplacedText 
            Height          =   495
            Left            =   120
            TabIndex        =   14
            ToolTipText     =   "The final text after replacements"
            Top             =   1680
            Width           =   1575
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblPresetOriginalText 
            Height          =   495
            Left            =   120
            TabIndex        =   13
            ToolTipText     =   "The original text that will have something replaced."
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Original Text"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   12
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Replaced Text"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   1320
            Width           =   1335
         End
      End
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Replace Characters"
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      ToolTipText     =   "Clicking will replace the characters in the Original Text."
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5520
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Replacing Characters"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   120
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmRemoveASCII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSubmit_Click()
            
    Dim finalResult As String
    Dim placeholder As String
    
    If sstabContainer.Tab = 0 Then

        placeholder = lblPresetOriginalText.Caption
        
        If chkPresetSpecialCharacters.Value = 1 Then
            
            placeholder = RemoveSpecialCharacters(lblPresetOriginalText.Caption)
                
        End If
        
        If chkPresetControlCharacters.Value = 1 Then
            
            placeholder = RemoveControlCharacters(placeholder)
            
        End If
        
        finalResult = RemoveCharacters(placeholder, txtPresetReplacements.Text)
        
        
        lblPresetReplacedText.Caption = finalResult
        
    ElseIf sstabContainer.Tab = 1 Then
    
        placeholder = txtCustomText.Text
        
        If chkCustomSpecialCharacters.Value = 1 Then
            
            placeholder = RemoveSpecialCharacters(txtCustomText.Text)
                
        End If
        
        If chkCustomControlCharacters.Value = 1 Then
            
            placeholder = RemoveControlCharacters(placeholder)
            
        End If
        
        finalResult = RemoveCharacters(placeholder, txtCustomReplacements.Text)
        
        
        lblCustomReplacedText.Caption = finalResult
        
    End If
    
End Sub

Private Sub Form_Load()

    lblPresetOriginalText.Caption = "Sample " + Chr(28) + Chr(29) + Chr(30) + Chr(31) + " Text"
    
    sstabContainer.Tab = 0
    
End Sub

Private Sub sstabContainer_Click(PreviousTab As Integer)

    If PreviousTab = 0 Then
        
        txtCustomText.TabStop = True
        txtCustomReplacements.TabStop = True
        chkCustomSpecialCharacters.TabStop = True
        chkCustomControlCharacters.TabStop = True
        
        txtPresetReplacements.TabStop = False
        chkPresetSpecialCharacters.TabStop = False
        chkPresetControlCharacters.TabStop = False
        
    Else
        
        txtCustomText.TabStop = False
        txtCustomReplacements.TabStop = False
        chkCustomSpecialCharacters.TabStop = False
        chkCustomControlCharacters.TabStop = False
        
        txtPresetReplacements.TabStop = True
        chkPresetSpecialCharacters.TabStop = True
        chkPresetControlCharacters.TabStop = True
    End If
    
End Sub


Function RemoveControlCharacters(source As String) As String

    Dim result As String
    Dim charnumber As Integer
    charnumber = 28
    
    result = source
        
    Do While charnumber < 32
        result = Replace$(result, Chr(charnumber), "")
        charnumber = charnumber + 1
    Loop

    RemoveControlCharacters = result
End Function

Function RemoveSpecialCharacters(source As String) As String

    Dim result As String
    Dim charnumber As Integer
    charnumber = 33
    
    result = source
        
    Do While charnumber < 127
        If charnumber >= 33 And charnumber <= 47 Then
            result = Replace$(result, Chr(charnumber), "")
            charnumber = charnumber + 1
        ElseIf charnumber >= 58 And charnumber <= 64 Then
            result = Replace$(result, Chr(charnumber), "")
            charnumber = charnumber + 1
        ElseIf charnumber >= 91 And charnumber <= 96 Then
            result = Replace$(result, Chr(charnumber), "")
            charnumber = charnumber + 1
        ElseIf charnumber >= 91 And charnumber <= 96 Then
            result = Replace$(result, Chr(charnumber), "")
            charnumber = charnumber + 1
        ElseIf charnumber >= 123 And charnumber <= 126 Then
            result = Replace$(result, Chr(charnumber), "")
            charnumber = charnumber + 1
        Else
            charnumber = charnumber + 1
        End If
    Loop
    
    RemoveSpecialCharacters = result
End Function

Function RemoveCharacters(source As String, characters As String) As String

    Dim result As String

    result = Replace$(source, characters, "")

    RemoveCharacters = result
End Function
