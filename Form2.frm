VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   10095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   10095
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   9735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      Begin VB.Timer Timer1 
         Left            =   240
         Top             =   480
      End
      Begin VB.CommandButton btnClear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   12
         Top             =   8880
         Width           =   1815
      End
      Begin VB.TextBox txtTin 
         Height          =   375
         Left            =   960
         TabIndex        =   10
         Text            =   "TIN NUMBER"
         Top             =   8160
         Width           =   2535
      End
      Begin VB.TextBox txtPoscode 
         Height          =   375
         Left            =   960
         TabIndex        =   9
         Text            =   "POSCODE"
         Top             =   7440
         Width           =   2535
      End
      Begin VB.TextBox txtAddress2 
         Height          =   375
         Left            =   960
         TabIndex        =   8
         Text            =   "ADDRESS2"
         Top             =   6720
         Width           =   2535
      End
      Begin VB.TextBox txtAddress1 
         Height          =   375
         Left            =   960
         TabIndex        =   7
         Text            =   "ADDRESS1"
         Top             =   6000
         Width           =   2535
      End
      Begin VB.TextBox txtEmail 
         Height          =   375
         Left            =   960
         TabIndex        =   6
         Text            =   "EMAIL"
         Top             =   5280
         Width           =   2535
      End
      Begin VB.TextBox txtPhoneNumber 
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Text            =   "PHONE NUMBER"
         Top             =   4560
         Width           =   2535
      End
      Begin VB.TextBox txtGender 
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Text            =   "GENDER"
         Top             =   3840
         Width           =   2535
      End
      Begin VB.TextBox txtDOB 
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Text            =   "DOB"
         Top             =   3120
         Width           =   2535
      End
      Begin VB.TextBox txtNRIC 
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Text            =   "NRIC"
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox txtName 
         Height          =   375
         HideSelection   =   0   'False
         Left            =   960
         TabIndex        =   1
         Text            =   "NAME"
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label txtSingnal 
         Caption         =   "Ready To Scan"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1200
         TabIndex        =   11
         Top             =   720
         Width           =   2535
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Private Sub Form_Load()
    ' Start the hook when the form opens [cite: 23]
    hHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0)
    txtSingnal.Caption = "READY"
End Sub

Public Sub ProcessGlobalScan(ByVal rawData As String)
    On Error GoTo BadJson
    
    Dim Json As Object
    Set Json = JsonConverter.ParseJson(rawData)
    
    ' Fill boxes - This works even if the cursor is inside one! [cite: 73]
    txtName.Text = Json("name")
    txtNRIC.Text = Json("nric")
    txtDOB.Text = Json("dob")
    txtGender.Text = Json("gender")
    txtPhoneNumber.Text = Json("phone")
    txtEmail.Text = Json("email")
    txtAddress1.Text = Json("address1")
    txtAddress2.Text = Json("address2")
    txtPoscode.Text = Json("postcode")
    txtTin.Text = Json("tin")

    txtSingnal.BackColor = vbGreen
    txtSingnal.Caption = "SCAN SUCCESS"
    Beep 1200, 200
    Exit Sub

BadJson:
    txtSingnal.Caption = "RE-SCAN QR"
    txtSingnal.BackColor = vbRed
End Sub

Private Sub btnClear_Click()
    txtName.Text = "": txtNRIC.Text = "": txtDOB.Text = ""
    txtGender.Text = "": txtPhoneNumber.Text = "": txtEmail.Text = ""
    txtAddress1.Text = "": txtAddress2.Text = "": txtPoscode.Text = ""
    txtTin.Text = ""
    txtSingnal.Caption = "READY": txtSingnal.BackColor = vbButtonFace
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnhookWindowsHookEx hHook ' Clean up [cite: 27]
End Sub
