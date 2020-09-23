VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmRank 
   Caption         =   "Google PageRank ™"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   144
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   504
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrLine 
      Height          =   45
      Left            =   30
      TabIndex        =   11
      Top             =   1830
      Width           =   7515
   End
   Begin VB.PictureBox PicStars 
      BorderStyle     =   0  'None
      Height          =   1980
      Left            =   1830
      Picture         =   "frmRank.frx":0000
      ScaleHeight     =   132
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   8
      Top             =   1740
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Timer Tim 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6750
      Top             =   1830
   End
   Begin VB.Frame FrRank 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   2828
      TabIndex        =   5
      Top             =   1050
      Width           =   1905
      Begin VB.PictureBox Pict 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   60
         Picture         =   "frmRank.frx":169B
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   9
         Top             =   390
         Width           =   960
      End
      Begin VB.Label lbTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "  PageRank"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   7
         Top             =   30
         Width           =   810
      End
      Begin VB.Label lbRank 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0/10"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1050
         TabIndex        =   6
         Top             =   300
         Width           =   810
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   0
         Top             =   0
         Width           =   2000
      End
   End
   Begin InetCtlsObjects.Inet Internet 
      Left            =   1140
      Top             =   1230
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Fr 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   7305
      Begin VB.TextBox txtHttp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   810
         TabIndex        =   10
         Text            =   "http://"
         Top             =   270
         Width           =   645
      End
      Begin VB.TextBox txtUrl 
         Height          =   285
         Left            =   1470
         TabIndex        =   1
         Top             =   270
         Width           =   4035
      End
      Begin VB.CommandButton cmdRank 
         Caption         =   "Get Rank!"
         Height          =   405
         Left            =   5640
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lbUrl 
         AutoSize        =   -1  'True
         Caption         =   "Some Url:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   690
      End
   End
   Begin VB.Label lbMe 
      AutoSize        =   -1  'True
      Caption         =   "Developed by Int21"
      Height          =   195
      Left            =   5850
      TabIndex        =   12
      Top             =   1530
      Width           =   1425
   End
   Begin VB.Label lbStatus 
      Alignment       =   2  'Center
      Caption         =   "Idle"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   1920
      Width           =   7335
   End
End
Attribute VB_Name = "frmRank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRank_Click()
'*******************************************
'Make the magic
'*******************************************
Dim chSum, strRes$, Rank%, strTmp, sUrl$
sUrl = txtHttp & txtUrl
    Tim.Enabled = True
'calculate Google Checksum
chSum = CalculateChecksum(sUrl)
'*************************************************
'Read from Google
    strRes = Internet.OpenURL("http://www.google.com/search?client=navclient-auto&ch=" & chSum & "&features=Rank&q=info:" & sUrl)
        
    
    strTmp = Split(strRes, ":")
    Rank = CInt(Mid(strTmp(2), 1, 1))
    Tim.Enabled = False
    lbRank = Rank & "/10"
    Pict.PaintPicture PicStars.Picture, 0, 0, 64, 12, 0, 12 * Rank, 64, 12
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Internet.Cancel
End Sub

Private Sub Internet_StateChanged(ByVal State As Integer)
    Select Case State
        Case 1
            lbStatus = "Resolving host..."
        Case 2
            lbStatus = "Host Resolved."
        Case 3
            lbStatus = "Connecting..."
        Case 4
            lbStatus = "Connected."
        Case 5
            lbStatus = "Requesting..."
        Case 6
            lbStatus = "Request send."
        Case 7
            lbStatus = "Receiving Response..."
        Case 8
            lbStatus = "Response received."
        Case 9
            lbStatus = "Disconnecting..."
        Case 10
            lbStatus = "Disconnected."
        Case 11
            lbStatus = "Some Error."
        Case 12
            lbStatus = "Response completed."
            
    End Select
End Sub

Private Sub Tim_Timer()
Dim rand%
rand = Int((10 - 0 + 1) * Rnd + 0) 'get random number
    lbRank = rand & "/10"
    Pict.PaintPicture PicStars.Picture, 0, 0, 64, 12, 0, 12 * rand, 64, 12
End Sub
