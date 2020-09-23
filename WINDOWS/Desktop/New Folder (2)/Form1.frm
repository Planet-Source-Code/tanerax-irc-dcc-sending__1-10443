VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form DccSend 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock DCC 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "DccSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// Programming By: Tanerax [Tanerax@nbnet.nb.ca] [50298342]
'// Program: DCC Sending v1.0
'// Comments: Complete
'// This Program is an example that works for sending files
'// on the IRC with Mirc. For the use of this program you
'// require the ctcp commands that are needed to send to
'// the person you are attempting to send the file to.
'// Anyone that has worked with IRC before should know how
'// or where to find these.
'// ********************************************************
'// If You Use This At All In A Program Please Mention Me
'// If You Can Optimize This Please Send Me A Optimized Version
'// Via E-Mail
'// Thank You.

Private Sub Form_Load()
    DCC.LocalPort = 1560                          '// Sets Port to 1560
    DCC.Listen                                    '// Sets Winsock To Listen
    DccSend.Caption = "1560 Listen....."          '// Changes Form Caption
End Sub

Function SendFile(sFile As String) As Boolean
Dim I, fLength, ret                               '// Declare Variables
Dim Buffer As String                              '// Declare Buffer
Dim bSize As Long                                 '// Declare BufferSize
bSize = 1024                                      '// Set BufferSize
I = FreeFile                                      '// Set I As FreeFile
    
    Open sFile For Binary Access Read As I        '// Open File For Binary Read
    fLength = LOF(I)                              '// Gets The File Length

        Do Until EOF(I)                           '// Begin A Loop Until EOF
                                                       
            If fLength - Loc(I) <= bSize Then     '// If The Buffer Is Larger Than
                bSize = fLength - Loc(I)          '// The Rest Of the File. Make The
            End If                                '// New Buffer Size The Rest Of The
                                                  '// File
            If bSize = 0 Then Exit Do             '// If Buffer Size Is 0 Send Done
        
            bytesent = bytesent + bSize           '// Adds The Buffer To Bytes Sent
            Buffer = Space$(bSize)                '// Get The Buffer From The BlockSize
            Get I, , Buffer                       '// Take Block From File
            DCC.SendData Buffer                   '// Send Block
        Loop                                      '// Loop
    Close I                                       '// Close File

DCC.Close                                         '// Close The Connectio
DccSend.Caption = "File Sent"                     '// Change Caption
SendFile = True                                   '// Return A True
End Function

Private Sub DCC_ConnectionRequest(ByVal requestID As Long)
    Dim Retval As Boolean
    If DCC.State <> sckClosed Then DCC.Close       '// If The State Is Not Close Close It
    DCC.Accept requestID                           '// Accept The Request ID
    DccSend.Caption = "Connection Established"     '// Change Caption
    Retval = SendFile("C:\command.com")            '// Begin FileSend
    If Retval = True Then DccSend.Caption = "File Send Successfull"
End Sub
