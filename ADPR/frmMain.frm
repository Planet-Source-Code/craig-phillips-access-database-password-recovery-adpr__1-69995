VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Access Database Password Recovery (ADPR)"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4230
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   0
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt2000Password 
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   3375
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txt9597Password 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   3375
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Default         =   -1  'True
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3720
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label5 
      Caption         =   "Access 2000+ Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "© Copyright Craig Phillips ADPR 2008"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Only for passwords < 18 characters"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "File:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Access 95/97 Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem ------------------------------------------------------------------------------------------------------------
Rem Access Database Password Recovery (ADPR)
Rem Recovers the passwords of most .mdb files
Rem © Copyright Craig Phillips, All rights reserved 2008-2009
Rem
Rem This program is free software: you can redistribute it and/or modify it under the terms of the GNU
Rem General Public License version 3 as published by the Free Software Foundation.
Rem
Rem This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without
Rem even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
Rem General Public License for more details.
Rem http://www.gnu.org/licenses/
Rem ------------------------------------------------------------------------------------------------------------
Rem Please do not use this code for any malicious activity
Rem I will not accept responsibility for any criminal act
Rem This code is purely for forgotten password recovery
Option Explicit

Private Sub cmdBrowse_Click()

    cd.Filter = "Microsoft Access Files (*.mdb)|*.mdb|All Files (*.*)|*.*"
    cd.DialogTitle = App.FileDescription
    cd.ShowOpen                                     ' Show open dialog
    
    If Not Len(cd.FileName) = 0 Then
        txtFile.Text = cd.FileName                  ' Put the filename into the textbox
        
        Call GetPassword                            ' Get the password
    End If

End Sub

Private Sub cmdClose_Click()

    End                                             ' End the program
    
End Sub

Private Function GetPassword()

    On Error GoTo ErrHand

    Dim Access2000Decode As Variant                 ' Decode Array (Access 2000)
    Dim Access9597Decode As Variant                 ' Decode Array (Access 95/97)
    
    Dim fFile       As Integer                      ' File Number
    Dim bCnt        As Integer                      ' Loop Count
    
    Dim ret95wd(17) As Byte                         ' Return 95/97 Password (max 18 chars)
    Dim retXPwd(17) As Integer                      ' Return File Password (max 18 chars)

    Dim wkCode      As Integer                      ' Working Code
    Dim mgCode      As Integer                      ' Magic Code
    
    'Define the Access 95/97 decode array
    Access9597Decode = Array(&H86, &HFB, &HEC, &H37, &H5D, &H44, &H9C, &HFA, &HC6, _
                             &H5E, &H28, &HE6, &H13, &HB6, &H8A, &H60, &H54, &H94)
    
    'Define the Access 2000 decode array
    Access2000Decode = Array(&H6ABA, &H37EC, &HD561, &HFA9C, &HCFFA, _
                      &HE628, &H272F, &H608A, &H568, &H367B, _
                      &HE3C9, &HB1DF, &H654B, &H4313, &H3EF3, _
                      &H33B1, &HF008, &H5B79, &H24AE, &H2A7C)

    If Len(txtFile.Text) > 0 Then                   ' If theres text in the file
    
        fFile = FreeFile                            ' Free File Channel
    
        Open txtFile.Text For Binary As #fFile      ' Open the file
            Get #fFile, 67, retXPwd                 ' Get Encoded Access 2000+ Password
            Get #fFile, 67, ret95wd                 ' Get Encoded Access 95/97 Password
            Get #fFile, 103, mgCode                 ' Get Magic code
        Close #fFile
        
        mgCode = mgCode Xor Access2000Decode(18)    ' Xor magic code

        txt9597Password.Text = vbNullString         ' Clear the 95/97 Password textbox
        txt2000Password.Text = vbNullString         ' Clear the 2000+ textbox

        For bCnt = 0 To 17
        
            ' Decode Access 95/97 Password
            wkCode = ret95wd(bCnt) Xor Access9597Decode(bCnt)
            txt9597Password.Text = txt9597Password.Text & Chr(wkCode)
        
            ' Decode Access 2000+ Password
            wkCode = retXPwd(bCnt) Xor Access2000Decode(bCnt)
            
            If wkCode < 256 Then                    ' Normal ASCII Code
                txt2000Password.Text = txt2000Password.Text & Chr(wkCode)
            Else                                    ' Un-normal; XOR with Magic Code
                txt2000Password.Text = txt2000Password.Text & Chr(wkCode Xor mgCode)
            End If
            
        Next bCnt
        
    Else
    
        txt2000Password.Text = "No file Selected"       ' No file
    
    End If
    
Exit Function
ErrHand:
    MsgBox "Error with opening file", vbCritical, App.Title


End Function

