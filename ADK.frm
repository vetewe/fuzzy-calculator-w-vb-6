VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Apk Derajat Keanggotaan"
   ClientHeight    =   11490
   ClientLeft      =   300
   ClientTop       =   2055
   ClientWidth     =   22485
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ADK.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11490
   ScaleWidth      =   22485
   Begin VB.CommandButton Command10 
      Caption         =   "Ulang"
      Height          =   495
      Left            =   17160
      TabIndex        =   58
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Ulang"
      Height          =   495
      Left            =   17160
      TabIndex        =   57
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Ulang"
      Height          =   495
      Left            =   17160
      TabIndex        =   56
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Ulang"
      Height          =   495
      Left            =   17160
      TabIndex        =   55
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Reset"
      Height          =   495
      Left            =   18480
      TabIndex        =   54
      Top             =   9480
      Width           =   1215
   End
   Begin VB.TextBox Text25 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   53
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox Text24 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   52
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox Text23 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   51
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox Text22 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   50
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox Text21 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   49
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox Text20 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   48
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox Text19 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   47
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox Text18 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   46
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Keluar"
      Height          =   495
      Left            =   19800
      MaskColor       =   &H8000000E&
      TabIndex        =   40
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Hitung"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14040
      TabIndex        =   39
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Hitung"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14040
      TabIndex        =   38
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hitung"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14040
      TabIndex        =   37
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox Text17 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   17520
      TabIndex        =   35
      Top             =   7200
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hitung"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14160
      TabIndex        =   34
      Top             =   6240
      Width           =   1095
   End
   Begin VB.TextBox Text16 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   33
      Top             =   8160
      Width           =   1215
   End
   Begin VB.TextBox Text15 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   31
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   29
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   27
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   26
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   25
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   24
      Top             =   6240
      Width           =   1215
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Trapesium"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   19
      Top             =   8280
      Width           =   2655
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Segitiga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   18
      Top             =   7680
      Width           =   2655
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Linear Turun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   17
      Top             =   7200
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Linear Naik"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   16
      Top             =   6720
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   15
      Top             =   3360
      Width           =   2535
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   13
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1440
      TabIndex        =   8
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   2880
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3840
      TabIndex        =   6
      Top             =   2400
      Width           =   6495
   End
   Begin VB.Label Label20 
      Caption         =   "®vtw"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15000
      TabIndex        =   59
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label19 
      Caption         =   "Fungsi Keanggotaan : Trapesium"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   45
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Label Label18 
      Caption         =   "Fungsi Keanggotaan : Segitiga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11640
      TabIndex        =   44
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label Label17 
      Caption         =   "Fungsi Keanggotaan : Linear Turun"
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
      Left            =   11400
      TabIndex        =   43
      Top             =   1800
      Width           =   4335
   End
   Begin VB.Line Line9 
      X1              =   10800
      X2              =   21000
      Y1              =   9240
      Y2              =   9240
   End
   Begin VB.Line Line8 
      X1              =   21000
      X2              =   21000
      Y1              =   1920
      Y2              =   9240
   End
   Begin VB.Line Line7 
      X1              =   15600
      X2              =   21000
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Fungsi Keanggotaan : Linear Naik"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11400
      TabIndex        =   42
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Line Line6 
      X1              =   10800
      X2              =   11280
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label15 
      Caption         =   "DATA INPUT"
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
      Left            =   1440
      TabIndex        =   41
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Line Line5 
      X1              =   3120
      X2              =   10800
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line4 
      X1              =   840
      X2              =   10800
      Y1              =   9240
      Y2              =   9240
   End
   Begin VB.Line Line3 
      X1              =   840
      X2              =   1320
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line2 
      X1              =   840
      X2              =   840
      Y1              =   1920
      Y2              =   9240
   End
   Begin VB.Label Label14 
      Caption         =   "Nilai Derajat Keanggotaan Nilai X Adalah"
      Height          =   375
      Left            =   13920
      TabIndex        =   36
      Top             =   7320
      Width           =   4695
   End
   Begin VB.Label Label13 
      Caption         =   "X :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   32
      Top             =   8160
      Width           =   375
   End
   Begin VB.Label Label12 
      Caption         =   "X :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   30
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label Label11 
      Caption         =   "X :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   28
      Top             =   7200
      Width           =   375
   End
   Begin VB.Label Label10 
      Caption         =   "d :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11520
      TabIndex        =   23
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "c :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11520
      TabIndex        =   22
      Top             =   7200
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "b :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11520
      TabIndex        =   21
      Top             =   6720
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "a :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   20
      Top             =   6240
      Width           =   375
   End
   Begin VB.Image Image4 
      Height          =   1950
      Left            =   11400
      Picture         =   "ADK.frx":2DD30
      Top             =   2760
      Width           =   4800
   End
   Begin VB.Image Image3 
      Height          =   2070
      Left            =   11400
      Picture         =   "ADK.frx":2F100
      Top             =   2760
      Width           =   4800
   End
   Begin VB.Image Image2 
      Height          =   2325
      Left            =   11400
      Picture         =   "ADK.frx":30573
      Top             =   2760
      Width           =   4800
   End
   Begin VB.Image Image1 
      Height          =   3420
      Left            =   11400
      Picture         =   "ADK.frx":31941
      Top             =   2280
      Width           =   4800
   End
   Begin VB.Line Line1 
      X1              =   10800
      X2              =   10800
      Y1              =   1920
      Y2              =   9240
   End
   Begin VB.Label Label6 
      Caption         =   "Silahkan Pilih Fungsi Keanggotaan Yang Ingin Digunakan"
      Height          =   255
      Left            =   1440
      TabIndex        =   14
      Top             =   6240
      Width           =   5295
   End
   Begin VB.Label Label5 
      Caption         =   "Domain"
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Himpunan "
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Semesta Pembicaraan :"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Variabel            :"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Nama Kasus               :"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label A 
      Caption         =   "FUZZY CALCULATOR"
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   0
      Top             =   720
      Width           =   8175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    If Val(Text18) <= Val(Text10) Then
        Text17.Text = 0
        Text17.Visible = True
        Label14.Visible = True
    Else
    If Val(Text18) >= Val(Text11) Then
        Text17.Text = 1
        Text17.Visible = True
        Label14.Visible = True
    Else
        Text17.Text = (Val(Text18) - Val(Text10)) / (Val(Text11) - Val(Text10))
        Text17.Visible = True
        Label14.Visible = True
    End If
    End If
End Sub

Private Sub Command2_Click()
    If Val(Text14) <= Val(Text19) Then
        Text17.Text = 1
        Text17.Visible = True
        Label14.Visible = True
    Else
    If Val(Text14) >= Val(Text20) Then
        Text17.Text = 0
        Text17.Visible = True
        Label14.Visible = True
    Else
        Text17.Text = (Val(Text20) - Val(Text14)) / (Val(Text20) - Val(Text19))
        Text17.Visible = True
        Label14.Visible = True
    End If
    End If
End Sub
Private Sub Command3_Click()
    If Val(Text15) <= Val(Text21) Or Val(Text15) >= Val(Text12) Then
        Text17.Text = 0
        Text17.Visible = True
        Label14.Visible = True
    Else
    If Val(Text15) >= Val(Text22) And Val(Text15) <= Val(Text12) Then
        Text17.Text = (Val(Text12) - Val(Text15)) / (Val(Text12) - Val(Text22))
        Text17.Visible = True
        Label14.Visible = True
    Else
    If Val(Text15) >= Val(Text21) And Val(Text15) <= Val(Text22) Then
        Text17.Text = (Val(Text15) - Val(Text21)) / (Val(Text22) - Val(Text21))
        Text17.Visible = True
        Label14.Visible = True
    Else
        Text17.Text = 1
        Text17.Visible = True
        Label14.Visible = True
    End If
    End If
    End If
End Sub
Private Sub Command4_Click()
    If Val(Text16) <= Val(Text23) Or Val(Text16) >= Val(Text13) Then
        Text17.Text = 0
        Text17.Visible = True
        Label14.Visible = True
    Else
    If Val(Text16) >= Val(Text23) And Val(Text16) <= Val(Text24) Then
        Text17.Text = (Val(Text16) - Val(Text23)) / (Val(Text24) - Val(Text23))
        Text17.Visible = True
        Label14.Visible = True
    Else
    If Val(Text16) >= Val(Text24) And Val(Text16) <= Val(Text25) Then
        Text17.Text = 1
        Text17.Visible = True
        Label14.Visible = True
    Else
        Text17.Text = (Val(Text13) - Val(Text16)) / (Val(Text13) - Val(Text25))
        Text17.Visible = True
        Label14.Visible = True
    End If
    End If
    End If
End Sub

Private Sub Command5_Click()
     If MsgBox("Apakah Anda Yakin Ingin Keluar?", vbExclamation + vbYesNo + vbDefaultButton2, "Konfirmasi") = vbYes Then End
End Sub

Private Sub Command7_Click()
    Text10.Text = Empty
    Text11.Text = Empty
    Text17.Text = Empty
    Text18.Text = Empty
    Text17.Visible = False
    Label14.Visible = False
    Command1.Enabled = False
End Sub
Private Sub Command8_Click()
    Text14.Text = Empty
    Text17.Text = Empty
    Text19.Text = Empty
    Text20.Text = Empty
    Text17.Visible = False
    Label14.Visible = False
    Command2.Enabled = False
End Sub
Private Sub Command9_Click()
    Text12.Text = Empty
    Text16.Text = Empty
    Text17.Text = Empty
    Text21.Text = Empty
    Text22.Text = Empty
    Text17.Visible = False
    Label14.Visible = False
    Command3.Enabled = False
End Sub
Private Sub Command10_Click()
    Text13.Text = Empty
    Text16.Text = Empty
    Text17.Text = Empty
    Text23.Text = Empty
    Text24.Text = Empty
    Text25.Text = Empty
    Text17.Visible = False
    Label14.Visible = False
    Command4.Enabled = False
End Sub
Sub Reset()
    Line6.Visible = False
    Line7.Visible = False
    Line8.Visible = False
    Line9.Visible = False
    Image1.Visible = False
    Image2.Visible = False
    Image3.Visible = False
    Image4.Visible = False
    Label7.Visible = False
    Label8.Visible = False
    Label9.Visible = False
    Label10.Visible = False
    Label11.Visible = False
    Label12.Visible = False
    Label13.Visible = False
    Label14.Visible = False
    Label16.Visible = False
    Label17.Visible = False
    Label18.Visible = False
    Label19.Visible = False
    Text1.Text = Empty
    Text2.Text = Empty
    Text3.Text = Empty
    Text4.Text = Empty
    Text5.Text = Empty
    Text6.Text = Empty
    Text7.Text = Empty
    Text8.Text = Empty
    Text9.Text = Empty
    Text10.Text = Empty
    Text11.Text = Empty
    Text12.Text = Empty
    Text13.Text = Empty
    Text14.Text = Empty
    Text15.Text = Empty
    Text16.Text = Empty
    Text17.Text = Empty
    Text18.Text = Empty
    Text19.Text = Empty
    Text20.Text = Empty
    Text21.Text = Empty
    Text22.Text = Empty
    Text23.Text = Empty
    Text24.Text = Empty
    Text25.Text = Empty
    Text10.Visible = False
    Text11.Visible = False
    Text12.Visible = False
    Text13.Visible = False
    Text14.Visible = False
    Text15.Visible = False
    Text16.Visible = False
    Text17.Visible = False
    Text18.Visible = False
    Text19.Visible = False
    Text20.Visible = False
    Text21.Visible = False
    Text22.Visible = False
    Text23.Visible = False
    Text24.Visible = False
    Text25.Visible = False
    Option1.Value = False
    Option2.Value = False
    Option3.Value = False
    Option4.Value = False
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Command1.Visible = False
    Command2.Visible = False
    Command3.Visible = False
    Command4.Visible = False
    Text1.SetFocus
End Sub

Private Sub Command6_Click()
    Reset
End Sub

Private Sub Form_Load()
    Line6.Visible = False
    Line7.Visible = False
    Line8.Visible = False
    Line9.Visible = False
    Image1.Visible = False
    Image2.Visible = False
    Image3.Visible = False
    Image4.Visible = False
    Label7.Visible = False
    Label8.Visible = False
    Label9.Visible = False
    Label10.Visible = False
    Label11.Visible = False
    Label12.Visible = False
    Label13.Visible = False
    Label14.Visible = False
    Label16.Visible = False
    Label17.Visible = False
    Label18.Visible = False
    Label19.Visible = False
    Text10.Visible = False
    Text11.Visible = False
    Text12.Visible = False
    Text13.Visible = False
    Text14.Visible = False
    Text15.Visible = False
    Text16.Visible = False
    Text17.Visible = False
    Text18.Visible = False
    Text19.Visible = False
    Text20.Visible = False
    Text21.Visible = False
    Text22.Visible = False
    Text23.Visible = False
    Text24.Visible = False
    Text25.Visible = False
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
    Command8.Enabled = False
    Command9.Enabled = False
    Command10.Enabled = False
    Command1.Visible = False
    Command2.Visible = False
    Command3.Visible = False
    Command4.Visible = False
    Command7.Visible = False
    Command8.Visible = False
    Command9.Visible = False
    Command10.Visible = False
End Sub


Private Sub Option1_Click()
    If Option1.Value = True Then
        Line6.Visible = True
        Line7.Visible = True
        Line8.Visible = True
        Line9.Visible = True
        Image1.Visible = True
        Image2.Visible = False
        Image3.Visible = False
        Image4.Visible = False
        Label7.Visible = True
        Label8.Visible = True
        Label9.Visible = False
        Label10.Visible = False
        Label11.Visible = True
        Label12.Visible = False
        Label13.Visible = False
        Label14.Visible = False
        Label16.Visible = True
        Label17.Visible = False
        Label18.Visible = False
        Label19.Visible = False
        Text10.Visible = True
        Text11.Visible = True
        Text12.Visible = False
        Text13.Visible = False
        Text14.Visible = False
        Text15.Visible = False
        Text16.Visible = False
        Text17.Visible = False
        Text18.Visible = True
        Text19.Visible = False
        Text20.Visible = False
        Text21.Visible = False
        Text22.Visible = False
        Text23.Visible = False
        Text24.Visible = False
        Text25.Visible = False
        Command1.Visible = True
        Command2.Visible = False
        Command3.Visible = False
        Command4.Visible = False
        Command7.Visible = True
        Command8.Visible = False
        Command9.Visible = False
        Command10.Visible = False
    End If
End Sub

Private Sub Option2_Click()
    If Option2.Value = True Then
        Line6.Visible = True
        Line7.Visible = True
        Line8.Visible = True
        Line9.Visible = True
        Image1.Visible = False
        Image2.Visible = True
        Image3.Visible = False
        Image4.Visible = False
        Label7.Visible = True
        Label8.Visible = True
        Label9.Visible = False
        Label10.Visible = False
        Label11.Visible = True
        Label12.Visible = False
        Label13.Visible = False
        Label14.Visible = False
        Label16.Visible = False
        Label17.Visible = True
        Label18.Visible = False
        Label19.Visible = False
        Text10.Visible = False
        Text11.Visible = False
        Text12.Visible = False
        Text13.Visible = False
        Text14.Visible = True
        Text15.Visible = False
        Text16.Visible = False
        Text17.Visible = False
        Text18.Visible = False
        Text19.Visible = True
        Text20.Visible = True
        Text21.Visible = False
        Text22.Visible = False
        Text23.Visible = False
        Text24.Visible = False
        Text25.Visible = False
        Command1.Visible = False
        Command2.Visible = True
        Command3.Visible = False
        Command4.Visible = False
        Command7.Visible = False
        Command8.Visible = True
        Command9.Visible = False
        Command10.Visible = False
    End If
End Sub

Private Sub Option3_Click()
    If Option3.Value = True Then
        Line6.Visible = True
        Line7.Visible = True
        Line8.Visible = True
        Line9.Visible = True
        Image1.Visible = False
        Image2.Visible = False
        Image3.Visible = True
        Image4.Visible = False
        Label7.Visible = True
        Label8.Visible = True
        Label9.Visible = True
        Label10.Visible = False
        Label11.Visible = False
        Label12.Visible = True
        Label13.Visible = False
        Label14.Visible = False
        Label16.Visible = False
        Label17.Visible = False
        Label18.Visible = True
        Label19.Visible = False
        Text10.Visible = False
        Text11.Visible = False
        Text12.Visible = True
        Text13.Visible = False
        Text14.Visible = False
        Text15.Visible = True
        Text16.Visible = False
        Text17.Visible = False
        Text18.Visible = False
        Text19.Visible = False
        Text20.Visible = False
        Text21.Visible = True
        Text22.Visible = True
        Text23.Visible = False
        Text24.Visible = False
        Text25.Visible = False
        Command1.Visible = False
        Command2.Visible = False
        Command3.Visible = True
        Command4.Visible = False
        Command7.Visible = False
        Command8.Visible = False
        Command9.Visible = True
        Command10.Visible = False
    End If
End Sub

Private Sub Option4_Click()
    If Option4.Value = True Then
        Line6.Visible = True
        Line7.Visible = True
        Line8.Visible = True
        Line9.Visible = True
        Image1.Visible = False
        Image2.Visible = False
        Image3.Visible = False
        Image4.Visible = True
        Label7.Visible = True
        Label8.Visible = True
        Label9.Visible = True
        Label10.Visible = True
        Label11.Visible = False
        Label12.Visible = False
        Label13.Visible = True
        Label14.Visible = False
        Label16.Visible = False
        Label17.Visible = False
        Label18.Visible = False
        Label19.Visible = True
        Text10.Visible = False
        Text11.Visible = False
        Text12.Visible = False
        Text13.Visible = True
        Text14.Visible = False
        Text15.Visible = False
        Text16.Visible = True
        Text17.Visible = False
        Text18.Visible = False
        Text20.Visible = False
        Text21.Visible = False
        Text22.Visible = False
        Text23.Visible = True
        Text24.Visible = True
        Text25.Visible = True
        Command1.Visible = False
        Command2.Visible = False
        Command3.Visible = False
        Command4.Visible = True
        Command7.Visible = False
        Command8.Visible = False
        Command9.Visible = False
        Command10.Visible = True
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text2.SetFocus
        Command6.Enabled = True
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text3.SetFocus
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text4.SetFocus
    End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text7.SetFocus
    End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text8.SetFocus
    End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text9.SetFocus
    End If
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text5.SetFocus
    End If
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text6.SetFocus
    End If
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Text9.Text = Empty Then
            MsgBox "Mohon Lengkapi Domain"
            Text9.SetFocus
        End If
    End If
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text11.SetFocus
        Command7.Visible = True
        Command7.Enabled = True
    End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text18.SetFocus
    End If
End Sub
Private Sub Text12_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text15.SetFocus
    End If
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text16.SetFocus
    End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Text14.Text = Empty Then
        MsgBox "Niai x Harus Diisi!!!"
        Text14.SetFocus
        Else
        Command2.Enabled = True
        Command2.SetFocus
        End If
    End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Text15.Text = Empty Then
        MsgBox "Niai x Harus Diisi!!!"
        Text15.SetFocus
        Else
        Command3.Enabled = True
        Command3.SetFocus
        End If
    End If
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Text16.Text = Empty Then
        MsgBox "Niai x Harus Diisi!!!"
        Text16.SetFocus
        Else
        Command4.Enabled = True
        Command4.SetFocus
        End If
    End If
End Sub
Private Sub Text18_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Text18.Text = Empty Then
        MsgBox "Niai x Harus Diisi!!!"
        Text18.SetFocus
        Else
        Command1.Enabled = True
        Command1.SetFocus
        End If
    End If
End Sub
Private Sub Text19_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text20.SetFocus
        Command8.Visible = True
        Command8.Enabled = True
    End If
End Sub
Private Sub Text20_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text14.SetFocus
    End If
End Sub
Private Sub Text21_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text22.SetFocus
        Command9.Visible = True
        Command9.Enabled = True
    End If
End Sub
Private Sub Text22_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text12.SetFocus
    End If
End Sub
Private Sub Text23_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text24.SetFocus
        Command10.Visible = True
        Command10.Enabled = True
    End If
End Sub
Private Sub Text24_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text25.SetFocus
    End If
End Sub
Private Sub Text25_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text13.SetFocus
    End If
End Sub
Private Sub Text2_GotFocus()
    If Text1.Text = Empty Then
        MsgBox "Mohon isi nama kasus"
        Text1.SetFocus
    End If
End Sub
Private Sub Text3_GotFocus()
    If Text2.Text = Empty Then
        MsgBox "Mohon isi nama variabel"
        Text2.SetFocus
    End If
End Sub
Private Sub Text4_GotFocus()
    If Text3.Text = Empty Then
        MsgBox "Semesta Pembicaraan Harus Diiai"
        Text3.SetFocus
    End If
End Sub
Private Sub Text7_GotFocus()
    If Text4.Text = Empty Then
        MsgBox "Mohon Lengkapi Himpunan"
        Text4.SetFocus
    End If
End Sub
Private Sub Text5_GotFocus()
    If Text7.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text7.SetFocus
    End If
End Sub
Private Sub Text8_GotFocus()
    If Text5.Text = Empty Then
        MsgBox "Mohon Lengkapi Himpunan"
        Text5.SetFocus
    End If
End Sub
Private Sub Text6_GotFocus()
    If Text8.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text8.SetFocus
    End If
End Sub
Private Sub Text9_GotFocus()
    If Text6.Text = Empty Then
        MsgBox "Mohon Lengkapi Himpunan"
        Text6.SetFocus
    End If
End Sub
Private Sub Text10_GotFocus()
    If Text1.Text = Empty Or Text2.Text = Empty Or Text3.Text = Empty Or Text4.Text = Empty Or Text5.Text = Empty Or Text6.Text = Empty Or Text7.Text = Empty Or Text8.Text = Empty Or Text9.Text = Empty Then
            MsgBox "Mohon Lengkapi Data Input"
            Text1.SetFocus
    End If
End Sub
Private Sub Text11_GotFocus()
    If Text10.Text = Empty Then
        MsgBox "Niai a Harus Diisi!!!"
        Text10.SetFocus
    End If
End Sub
Private Sub Text12_GotFocus()
    If Text22.Text = Empty Then
        MsgBox "Niai b Harus Diisi!!!"
        Text22.SetFocus
    End If
End Sub
Private Sub Text13_GotFocus()
    If Text25.Text = Empty Then
        MsgBox "Niai c Harus Diisi!!!"
        Text25.SetFocus
    End If
End Sub
Private Sub Text15_GotFocus()
    If Text12.Text = Empty Then
        MsgBox "Niai c Harus Diisi!!!"
        Text12.SetFocus
    End If
End Sub
Private Sub Text14_GotFocus()
    If Text20.Text = Empty Then
        MsgBox "Niai b Harus Diisi!!!"
        Text20.SetFocus
    End If
End Sub
Private Sub Text16_GotFocus()
    If Text13.Text = Empty Then
        MsgBox "Niai d Harus Diisi!!!"
        Text13.SetFocus
    End If
End Sub
Private Sub Text18_GotFocus()
    If Text11.Text = Empty Then
        MsgBox "Niai b Harus Diisi!!!"
        Text11.SetFocus
    End If
End Sub
Private Sub Text19_GotFocus()
    If Text1.Text = Empty Or Text2.Text = Empty Or Text3.Text = Empty Or Text4.Text = Empty Or Text5.Text = Empty Or Text6.Text = Empty Or Text7.Text = Empty Or Text8.Text = Empty Or Text9.Text = Empty Then
            MsgBox "Mohon Lengkapi Data Input"
            Text1.SetFocus
    End If
End Sub
Private Sub Text20_GotFocus()
    If Text19.Text = Empty Then
        MsgBox "Niai a Harus Diisi!!!"
        Text19.SetFocus
    End If
End Sub
Private Sub Text21_GotFocus()
    If Text1.Text = Empty Or Text2.Text = Empty Or Text3.Text = Empty Or Text4.Text = Empty Or Text5.Text = Empty Or Text6.Text = Empty Or Text7.Text = Empty Or Text8.Text = Empty Or Text9.Text = Empty Then
            MsgBox "Mohon Lengkapi Data Input"
            Text1.SetFocus
    End If
End Sub
Private Sub Text22_GotFocus()
    If Text21.Text = Empty Then
        MsgBox "Niai a Harus Diisi!!!"
        Text21.SetFocus
    End If
End Sub
Private Sub Text23_GotFocus()
    If Text1.Text = Empty Or Text2.Text = Empty Or Text3.Text = Empty Or Text4.Text = Empty Or Text5.Text = Empty Or Text6.Text = Empty Or Text7.Text = Empty Or Text8.Text = Empty Or Text9.Text = Empty Then
            MsgBox "Mohon Lengkapi Data Input"
            Text1.SetFocus
    End If
End Sub
Private Sub Text24_GotFocus()
    If Text23.Text = Empty Then
        MsgBox "Niai a Harus Diisi!!!"
        Text23.SetFocus
    End If
End Sub
Private Sub Text25_GotFocus()
    If Text24.Text = Empty Then
        MsgBox "Niai b Harus Diisi!!!"
        Text24.SetFocus
    End If
End Sub
