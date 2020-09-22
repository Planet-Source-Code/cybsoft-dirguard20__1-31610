VERSION 5.00
Begin VB.Form FrmNote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert a remark"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "FrmNote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   6165
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   255
         Left            =   4560
         TabIndex        =   4
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   255
         Left            =   3720
         TabIndex        =   3
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdDoIt 
         Caption         =   "&Ok"
         Height          =   255
         Left            =   5400
         TabIndex        =   2
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox chkStamp 
         Caption         =   "Add timestamp"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   680
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5895
      End
   End
End
Attribute VB_Name = "FrmNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
txtNote.SetFocus
txtNote.SelStart = 0                            ' select previous tekst from 0
txtNote.SelLength = Len(txtNote.Text)           ' to its length
End Sub
Private Sub cmdDoIt_Click()
If Len(Trim(txtNote.Text)) > 0 Then
    If chkStamp.Value = 1 Then
        frmGuard.rtbChangedfiles.Text = frmGuard.rtbChangedfiles.Text & Date & "  " & Time & "   " & txtNote.Text & vbCrLf
    Else
        frmGuard.rtbChangedfiles.Text = frmGuard.rtbChangedfiles.Text & txtNote.Text & vbCrLf
    End If
End If
FrmNote.Hide
frmGuard.Refresh
End Sub
Private Sub cmdClear_Click()
txtNote.Text = ""
frmGuard.Refresh
End Sub
Private Sub CmdCancel_Click()
txtNote.Text = ""
FrmNote.Hide
frmGuard.Refresh
End Sub
