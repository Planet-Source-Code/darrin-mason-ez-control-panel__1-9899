VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Call Control Panel Functions From Your Program"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   2340
      TabIndex        =   14
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdPanel 
      Caption         =   "Game Controller Properties"
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton cmdPanel 
      Caption         =   "Dialing Properties"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton cmdPanel 
      Caption         =   "ODBC Data Source Administrator"
      Height          =   375
      Index           =   9
      Left            =   3000
      TabIndex        =   9
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton cmdPanel 
      Caption         =   "Regional Settings Properties"
      Height          =   375
      Index           =   12
      Left            =   3000
      TabIndex        =   12
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton cmdPanel 
      Caption         =   "Internet Explorer Properties"
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton cmdPanel 
      Caption         =   "Password Properties"
      Height          =   375
      Index           =   10
      Left            =   3000
      TabIndex        =   10
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton cmdPanel 
      Caption         =   "Network Properties"
      Height          =   375
      Index           =   8
      Left            =   3000
      TabIndex        =   8
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton cmdPanel 
      Caption         =   "System Properties"
      Height          =   375
      Index           =   13
      Left            =   3000
      TabIndex        =   13
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton cmdPanel 
      Caption         =   "Power Management Properties"
      Height          =   375
      Index           =   11
      Left            =   3000
      TabIndex        =   11
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton cmdPanel 
      Caption         =   "Multimedia Properties"
      Height          =   375
      Index           =   7
      Left            =   3000
      TabIndex        =   7
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton cmdPanel 
      Caption         =   "Modems Properties"
      Height          =   375
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton cmdPanel 
      Caption         =   "Display Settings"
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton cmdPanel 
      Caption         =   "Add/Remove Programs"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton cmdPanel 
      Caption         =   "Date && Time"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Source Code By Darrin P. Mason

' This source code may be freely used by all for any purpose.

' See below for more documentation

Option Explicit
                
Public CPL As String


Private Sub cmdPanel_Click(Index As Integer)
                                             
    Select Case Index
        Case 0
            CPL = "Appwiz.cpl" ' Add/Remove Programs
        Case 1
            CPL = "Timedate.cpl" ' Date & Time
        Case 2
            CPL = "Telephon.cpl" 'Dialing Properties
        Case 3
            CPL = "Desk.cpl" ' Display Settings
        Case 4
            CPL = "Joy.cpl" ' Game Controller Properties
        Case 5
            CPL = "Inetcpl.cpl" ' Internet Explorer Properties
        Case 6
            CPL = "Modem.cpl" ' Modems Properties
        Case 7
            CPL = "MMSys.cpl" ' Multimedia Properties
        Case 8
            CPL = "Netcpl.cpl" ' Network Properties
        Case 9
            CPL = "ODBCcp32.cpl" 'ODBC Data Source Administrator
        Case 10
            CPL = "Password.cpl" ' Password Properties
        Case 11
            CPL = "Powercfg.cpl" ' Power Management Properties
        Case 12
            CPL = "Intl.cpl" ' Regional Settings Properties
        Case 13
            CPL = "Sysdm.cpl" ' System Properties
    End Select
    OpenCPL
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmMain = Nothing
End Sub
Public Sub OpenCPL()
                     
    Dim MyShell As Double
    MyShell = Shell("Control " & CPL, vbNormalFocus)
End Sub
                                                     
    ' You do not have to give the path to Control.exe because it is
    ' a Windows system file. .cpl files are located in the folder
    ' C:\Windows\System or C:\Winnt\System, depending on your
    ' Operating System.
    
    ' This was written on the Windows 98 Second Edition platform. I'm not sure
    ' if all of the .cpl files in Windows NT are named the same as Windows 98.
    ' If one of these buttons doesn't work, you can test the .cpl files on your
    ' system by clicking Start, going to Run, and typing 'Control [cplname].cpl'
    ' (Example: Control Appwiz.cpl) Just keep in mind, if the .cpl didn't come with
    ' the operating system, you cannot guarantee that another user has that .cpl and
    ' therefore the shell won't work. You probably won't get an error, you just won't
    ' see anything happen! :)
    
    ' You can use this routine to open ANY .cpl (Control Panel) file. I've used it
    ' to open my Scanner Properties and Tweak UI for example.
    
    ' You can also shell without a variable like the example below:

    ' Dim MyShell
    ' MyShell = Shell("Control Appwiz.cpl", vbNormalFocus)

    ' Enjoy!
