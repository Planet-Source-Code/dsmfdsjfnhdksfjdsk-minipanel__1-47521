VERSION 5.00
Begin VB.Form frmControlPanel 
   Caption         =   "MiniPanel"
   ClientHeight    =   3420
   ClientLeft      =   2400
   ClientTop       =   1425
   ClientWidth     =   5265
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   5265
   Begin VB.CommandButton cmdKeyboard 
      Caption         =   "Keyboard Properties"
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton cmdMouse 
      Caption         =   "Mouse Properties"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton cmdDateTime 
      Caption         =   "Date/Time Properties"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton cmdDial 
      Caption         =   "Dialing Properties"
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton cmdInternet 
      Caption         =   "Internet"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton cmdAddRemoveProg 
      Caption         =   "Add/Remove Programs"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton cmdAccessibility 
      Caption         =   "Accessibility Properties"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmdMultimedia 
      Caption         =   "Multimedia Properties"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmdRegional 
      Caption         =   "Regional Settings"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display Properties"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''
'   MiniPanel                                       '
'      Please vote for this project, like you don't '
'   care...                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''



Private Type CPLINFO
    idIcon          As Long
    idName          As Long
    idInfo          As Long
    lData           As Long
End Type


Private Declare Function CPlApplet_Desk Lib "desk.cpl" Alias "CPlApplet" (ByVal hwndCPl As Long, ByVal uMsg As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Private Declare Function CPlApplet_Intl Lib "intl.cpl" Alias "CPlApplet" (ByVal hwndCPl As Long, ByVal uMsg As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Private Declare Function CPlApplet_MMSys Lib "mmsys.cpl" Alias "CPlApplet" (ByVal hwndCPl As Long, ByVal uMsg As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Private Declare Function CPlApplet_Access Lib "access.cpl" Alias "CPlApplet" (ByVal hwndCPl As Long, ByVal uMsg As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Private Declare Function CPlApplet_AppWiz Lib "appwiz.cpl" Alias "CPlApplet" (ByVal hwndCPl As Long, ByVal uMsg As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Private Declare Function CPlApplet_InetCpl Lib "inetcpl.cpl" Alias "CPlApplet" (ByVal hwndCPl As Long, ByVal uMsg As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Private Declare Function CPlApplet_Telephon Lib "telephon.cpl" Alias "CPlApplet" (ByVal hwndCPl As Long, ByVal uMsg As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Private Declare Function CPlApplet_TimeDate Lib "timedate.cpl" Alias "CPlApplet" (ByVal hwndCPl As Long, ByVal uMsg As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Private Declare Function CPlApplet_Main Lib "main.cpl" Alias "CPlApplet" (ByVal hwndCPl As Long, ByVal uMsg As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long

'Control Panel Messages Constants:
Private Const CPL_INIT = 1
Private Const CPL_GETCOUNT = 2
Private Const CPL_INQUIRE = 3
Private Const CPL_SELECT = 4
Private Const CPL_DBLCLK = 5
Private Const CPL_STOP = 6
Private Const CPL_EXIT = 7
Private Const CPL_NEWINQUIRE = 8

Private ci      As CPLINFO

Private Sub MainCplApplet(lParam1 As Long)
    If CPlApplet_Main(hWnd, CPL_INIT, 0, 0) <> 0 Then
        CPlApplet_Main hWnd, CPL_INQUIRE, lParam1, VarPtr(ci)
        CPlApplet_Main hWnd, CPL_DBLCLK, lParam1, ci.lData
        CPlApplet_Main hWnd, CPL_STOP, 0, ci.lData
        CPlApplet_Main hWnd, CPL_EXIT, 0, 0
    End If
End Sub

'Accessibility Properties
Private Sub cmdAccessibility_Click()
    If CPlApplet_Access(hWnd, CPL_INIT, 0, 0) <> 0 Then
        CPlApplet_Access hWnd, CPL_INQUIRE, 0, VarPtr(ci)
        CPlApplet_Access hWnd, CPL_DBLCLK, 0, ci.lData
        CPlApplet_Access hWnd, CPL_STOP, 0, ci.lData
        CPlApplet_Access hWnd, CPL_EXIT, 0, 0
    End If
End Sub

'Add/Remove Programs
Private Sub cmdAddRemoveProg_Click()
    If CPlApplet_AppWiz(hWnd, CPL_INIT, 0, 0) <> 0 Then
        CPlApplet_AppWiz hWnd, CPL_INQUIRE, 0, VarPtr(ci)
        CPlApplet_AppWiz hWnd, CPL_DBLCLK, 0, ci.lData
        CPlApplet_AppWiz hWnd, CPL_STOP, 0, ci.lData
        CPlApplet_AppWiz hWnd, CPL_EXIT, 0, 0
    End If
End Sub

'Date/Time Properties
Private Sub cmdDateTime_Click()
    If CPlApplet_TimeDate(hWnd, CPL_INIT, 0, 0) <> 0 Then
        CPlApplet_TimeDate hWnd, CPL_INQUIRE, 0, VarPtr(ci)
        CPlApplet_TimeDate hWnd, CPL_DBLCLK, 0, ci.lData
        CPlApplet_TimeDate hWnd, CPL_STOP, 0, ci.lData
        CPlApplet_TimeDate hWnd, CPL_EXIT, 0, 0
    End If
End Sub

'Dialing Properties
Private Sub cmdDial_Click()
    If CPlApplet_Telephon(hWnd, CPL_INIT, 0, 0) <> 0 Then
        CPlApplet_Telephon hWnd, CPL_INQUIRE, 0, VarPtr(ci)
        CPlApplet_Telephon hWnd, CPL_DBLCLK, 0, ci.lData
        CPlApplet_Telephon hWnd, CPL_STOP, 0, ci.lData
        CPlApplet_Telephon hWnd, CPL_EXIT, 0, 0
    End If
End Sub

'Display Properties
Private Sub cmdDisplay_Click()
    If CPlApplet_Desk(hWnd, CPL_INIT, 0, 0) <> 0 Then
        CPlApplet_Desk hWnd, CPL_INQUIRE, 0, VarPtr(ci)
        CPlApplet_Desk hWnd, CPL_DBLCLK, 0, ci.lData
        CPlApplet_Desk hWnd, CPL_STOP, 0, ci.lData
        CPlApplet_Desk hWnd, CPL_EXIT, 0, 0
    End If
End Sub

'Internet
Private Sub cmdInternet_Click()
    If CPlApplet_InetCpl(hWnd, CPL_INIT, 0, 0) <> 0 Then
        CPlApplet_InetCpl hWnd, CPL_INQUIRE, 0, VarPtr(ci)
        CPlApplet_InetCpl hWnd, CPL_DBLCLK, 0, ci.lData
        CPlApplet_InetCpl hWnd, CPL_STOP, 0, ci.lData
        CPlApplet_InetCpl hWnd, CPL_EXIT, 0, 0
    End If
End Sub

'Keyboard Properties
Private Sub cmdKeyboard_Click()
    MainCplApplet 1
End Sub


'Mouse Properties
Private Sub cmdMouse_Click()
    MainCplApplet 0
End Sub

'Multimedia Properties
Private Sub cmdMultimedia_Click()
    If CPlApplet_MMSys(hWnd, CPL_INIT, 0, 0) <> 0 Then
        CPlApplet_MMSys hWnd, CPL_INQUIRE, 0, VarPtr(ci)
        CPlApplet_MMSys hWnd, CPL_DBLCLK, 0, ci.lData
        CPlApplet_MMSys hWnd, CPL_STOP, 0, ci.lData
        CPlApplet_MMSys hWnd, CPL_EXIT, 0, 0
    End If
End Sub

'Regional Settings
Private Sub cmdRegional_Click()
    If CPlApplet_Intl(hWnd, CPL_INIT, 0, 0) <> 0 Then
        CPlApplet_Intl hWnd, CPL_INQUIRE, 0, VarPtr(ci)
        CPlApplet_Intl hWnd, CPL_DBLCLK, 0, ci.lData
        CPlApplet_Intl hWnd, CPL_STOP, 0, ci.lData
        CPlApplet_Intl hWnd, CPL_EXIT, 0, 0
    End If
End Sub



Private Sub Form_Load()
    MsgBox "WARNING: To reduce the risk of system crashes during code execution, please compile this application", vbExclamation, "Programmer's Heed"
End Sub
