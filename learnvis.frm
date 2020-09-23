VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Learning Visual Basic"
   ClientHeight    =   4545
   ClientLeft      =   720
   ClientTop       =   1470
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdcopy 
      Caption         =   "--> textbox"
      Height          =   375
      Left            =   1200
      TabIndex        =   41
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtcode 
      Height          =   1665
      Left            =   2085
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   40
      Top             =   2115
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "next"
      Height          =   375
      Left            =   4920
      TabIndex        =   38
      Top             =   3960
      Width           =   650
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "back"
      Height          =   375
      Left            =   4320
      TabIndex        =   37
      Top             =   3960
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.CommandButton cmdbc9 
      Caption         =   "back"
      Height          =   375
      Left            =   4320
      TabIndex        =   29
      Top             =   3960
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.CommandButton cmdA0 
      Caption         =   "next"
      Height          =   375
      Left            =   4920
      TabIndex        =   28
      Top             =   3960
      Visible         =   0   'False
      Width           =   650
   End
   Begin VB.ComboBox cmbvtyp 
      Height          =   315
      ItemData        =   "LEARNVIS.frx":0000
      Left            =   0
      List            =   "LEARNVIS.frx":0002
      TabIndex        =   27
      Text            =   "VarTypes"
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbansi 
      Height          =   315
      ItemData        =   "LEARNVIS.frx":0004
      Left            =   0
      List            =   "LEARNVIS.frx":0006
      TabIndex        =   26
      Text            =   "ANSI"
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   120
      TabIndex        =   25
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdgoback 
      Caption         =   "Unmore"
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdsimplify 
      Caption         =   "More.."
      Height          =   375
      Left            =   360
      TabIndex        =   23
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtin 
      Height          =   285
      Left            =   480
      TabIndex        =   22
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdcd5 
      Caption         =   "5"
      Height          =   255
      Left            =   6120
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdcd4 
      Caption         =   "4"
      Height          =   255
      Left            =   5880
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdcd3 
      Caption         =   "3"
      Height          =   255
      Left            =   5640
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdcd2 
      Caption         =   "2"
      Height          =   255
      Left            =   5400
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdcd1 
      Caption         =   "1"
      Height          =   255
      Left            =   5160
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdrnd 
      Caption         =   "randm"
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox lstdt 
      Height          =   840
      ItemData        =   "LEARNVIS.frx":0008
      Left            =   5850
      List            =   "LEARNVIS.frx":0027
      TabIndex        =   14
      Top             =   -315
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdboxoffice 
      Caption         =   "calc"
      Height          =   375
      Left            =   840
      TabIndex        =   13
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmd2lv1 
      Caption         =   "+2"
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmd1lv1 
      Caption         =   "+1"
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmd2mv1 
      Caption         =   "+2"
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmd1mv1 
      Caption         =   "+1"
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      Max             =   22
      Min             =   8
      TabIndex        =   8
      Top             =   3720
      Value           =   8
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdrun 
      Caption         =   "&Run"
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox piccode 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   3315
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox picTable 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Height          =   1095
      Left            =   3840
      ScaleHeight     =   1035
      ScaleWidth      =   2475
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblobj 
      Caption         =   "last"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   120
      TabIndex        =   39
      Top             =   4350
      Width           =   615
   End
   Begin VB.Label lblalg 
      Caption         =   "Algebra Reference"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   210
      TabIndex        =   36
      Top             =   165
      Visible         =   0   'False
      Width           =   1825
   End
   Begin VB.Label lblqa2 
      Caption         =   "You just clicked a text label, what are you going to do next?   ""i'm going to Disneyland!"""
      ForeColor       =   &H00404040&
      Height          =   4575
      Left            =   1565
      TabIndex        =   35
      Top             =   500
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label lblqa1 
      Caption         =   $"LEARNVIS.frx":039D
      ForeColor       =   &H00400000&
      Height          =   4575
      Left            =   1305
      TabIndex        =   34
      Top             =   1000
      Visible         =   0   'False
      Width           =   3945
   End
   Begin VB.Label lblqadv 
      Caption         =   "For Advanced Reference"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblqskip 
      Caption         =   "Skip to:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   75
      TabIndex        =   32
      Top             =   3750
      Width           =   735
   End
   Begin VB.Label lblfunct 
      Caption         =   "make a Function"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   110
      TabIndex        =   31
      Top             =   4150
      Width           =   1335
   End
   Begin VB.Label lblbuil 
      Caption         =   "Reference Section"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   90
      TabIndex        =   30
      Top             =   3975
      Width           =   1455
   End
   Begin VB.Label lblqpage 
      Caption         =   "Page"
      Height          =   255
      Left            =   4680
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Line linept2 
      Visible         =   0   'False
      X1              =   3720
      X2              =   3480
      Y1              =   2280
      Y2              =   2640
   End
   Begin VB.Line linept1 
      Visible         =   0   'False
      X1              =   3360
      X2              =   3720
      Y1              =   2400
      Y2              =   2280
   End
   Begin VB.Line lineshaft 
      Visible         =   0   'False
      X1              =   2280
      X2              =   3720
      Y1              =   3360
      Y2              =   2280
   End
   Begin VB.Label lblqprop 
      Alignment       =   1  'Right Justify
      Caption         =   "Properties"
      Height          =   255
      Left            =   5160
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblqcom 
      Caption         =   "comment:"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblcom 
      Height          =   855
      Left            =   1680
      TabIndex        =   3
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label lblmain 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   $"LEARNVIS.frx":04F6
      Height          =   780
      Left            =   1635
      TabIndex        =   1
      Top             =   480
      Width           =   3540
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblchaptr 
      Alignment       =   2  'Center
      Caption         =   "Welcome!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Visual Basic tutorial inspired by the notion that a _
10-year-old could program if they only knew the code _
and of course, had a vision.

Dim caseinpoint As Integer 'see VScroll1_Scroll
Dim mn As String, sn As String 'VScroll1_Scroll
Dim backtrue As Boolean 'specific select case values _
 in contentspre() use this to display different data _
 if the back button was most-recently clicked.
 
Dim prevtext As String 'txtin_Change, txtin_KeyPress
 
'Form resize variables applied to every control
Dim scw As Single, sch As Single, _
 fsr As Single, fs As Single
 
'these get a value so I can Call _
 qscolor("string",blue,1) see Form_Load for vals
Dim blu As Byte, blue As Byte
Dim grn As Byte, green As Byte
Dim mag As Byte, magenta As Byte
Dim bla As Byte, black As Byte
Dim whi As Byte, white As Byte
Dim cyan As Byte
Dim yel As Byte, yellow As Byte
Dim red As Byte

Dim clicktwice(1 To 5) As Byte 'used by piccode_Click

Dim cancelbyte(1) As Byte 'disables click for lblmain _
 after first time.  If I actually disable lblmain, _
 the text becomes light.

Dim aIndex 'added Tue Oct 24, 2000 for cmdA0 and _
 cmdbc9 - only two next and back buttons, beyond _
 the old and kept method of click next, new next _
 becomes Visible, old becomes Invis..
 
Dim pIndex 'pre-index - temporary while I convert _
 previous 'multi-next/back' into another _
 'only one back, only one next' set.
 
Dim q As String 'will represent double quote
Dim qs(1 To 99) As String 'piccode lines of text
Dim qps As String 'qps = "Private Sub "
Dim qhello As String
Dim fn1 As Integer

Dim strM5, strM7 As String 'string for commentary
Dim strVi, strv2, strc3 As String 'more comments
Dim mv1 As Integer 'used for demonstration - see cmdvar1
Dim subj As Byte 'subjects are assigned a value and _
 accessed by 'Page' (cmdcd*) commandbuttons
Dim npage As Byte 'cmdc6 will 'nextpage' rather than _
 'nextsubject'.
Dim com5 As String 'comment variable
Dim firsttime(1 To 5) As Byte 'when firsttime(n) = 1, _
 subroutines do not (disable buttons to guide the user)
Dim clscount As Integer 'Clears pic at so many lines
 
'For graphics section
Dim X(40) As Single, Y(40) As Single, r(40) As Single
Dim incred(40) As Single, incgrn(40) As Single, incblu(40) As Single
Dim ared(40) As Integer, agrn(40) As Integer, ablu(40) As Integer
Dim colorshift As Integer, cs As Integer, sc As Integer, scl As Integer
Dim old6785 As Integer, old4950 As Integer

'Control (object) dimensions set by case Index and aIndex _
(when back or next is clicked) found in contentspre() and _
contents(), and changed in form_resize()
Dim cmdnextwidth As Integer, cmdnextheight As Integer, _
 cmdnextleft As Integer, cmdnexttop As Integer
Dim cmdnextfs As Single 'fs stands for fontsize
Dim cmdbackwidth As Integer, cmdbackheight As Integer, _
 cmdbackleft As Integer, cmdbacktop As Integer
Dim cmdbackfs As Single
Dim cmdbc9width As Integer, cmdbc9height As Integer, _
 cmdbc9left As Integer, cmdcd9top As Integer
Dim cmdbc9fs As Single
Dim cmda0width As Integer, cmda0height As Integer, _
 cmda0left As Integer, cmda0top As Integer
Dim cmda0fs As Single

Dim cd1width As Integer, cd1height As Integer, _
 cd1left As Integer, cd1top As Integer
Dim cd2width As Integer, cd2height As Integer, _
 cd2left As Integer, cd2top As Integer
Dim cd3width As Integer, cd3height As Integer, _
 cd3left As Integer, cd3top As Integer
Dim cd4width As Integer, cd4height As Integer, _
 cd4left As Integer, cd4top As Integer
Dim cd5width As Integer, cd5height As Integer, _
 cd5left As Integer, cd5top As Integer

Dim lblalgwidth As Integer, lblalgheight As Integer, _
 lblalgleft As Integer, lblalgtop As Integer
Dim lblobjwidth As Integer, lblobjheight As Integer, _
 lblobjleft As Integer, lblobjtop As Integer
Dim lblfunctwidth As Integer, lblfunctheight As Integer, _
 lblfunctleft As Integer, lblfuncttop As Integer
Dim lblbuilwidth As Integer, lblbuilheight As Integer, _
 lblbuilleft As Integer, lblbuiltop As Integer

Dim lblqa1width As Integer, lblqa1height As Integer, _
 lblqa1left As Integer, lblqa1top As Integer
Dim lblqa2width As Integer, lblqa2height As Integer, _
 lblqa2left As Integer, lblqa2top As Integer

Dim lblmainwidth As Integer, lblmainheight As Integer, _
 lblmainleft As Integer, lblmaintop As Integer
Dim lblcomwidth As Integer, lblcomheight As Integer, _
 lblcomleft As Integer, lblcomtop As Integer
Dim lblqadvwidth As Integer, lblqadvheight As Integer, _
 lblqadvleft As Integer, lblqadvtop As Integer
Dim lblqcomwidth As Integer, lblqcomheight As Integer, _
 lblqcomleft As Integer, lblqcomtop As Integer
Dim lblqpagewidth As Integer, lblqpageheight As Integer, _
 lblqpageleft As Integer, lblqpagetop As Integer
Dim lblqpropwidth As Integer, lblqpropheight As Integer, _
 lblqoropleft As Integer, lblqproptop As Integer
Dim lblqskipwidth As Integer, lblqskipheight As Integer, _
 lblqskipleft As Integer, lblqskiptop As Integer
Dim lblchaptrwidth As Integer, lblchaptrheight As Integer, _
 lblchaptrleft As Integer, lblchaptrtop As Integer
Dim cmdcopywidth As Integer, cmdcopyheight As Integer, _
 cmdcopyleft As Integer, cmdcopytop As Integer
 
Dim cmd1lv1width As Integer, cmd1lv1height As Integer, _
 cmd1lv1left As Integer, cmd1lv1top As Integer
Dim cmd2lv1width As Integer, cmd2lv1height As Integer, _
 cmd2lv1left As Integer, cmd2lv1top As Integer
Dim cmd1mv1width As Integer, cmd1mv1height As Integer, _
 cmd1mv1left As Integer, cmd1mv1top As Integer
Dim cmd2mv1width As Integer, cmd2mv1height As Integer, _
 cmd2mv1left As Integer, cmd2mv1top As Integer
 

Dim cmdgobackwidth As Integer, cmdgobackheight As Integer, _
 cmdgobackleft As Integer, cmdgobacktop As Integer
Dim cmdsimplifywidth As Integer, cmdsimplifyheight As Integer, _
 cmdsimplifyleft As Integer, cmdsimplifytop As Integer
 
Dim cmdrunwidth As Integer, cmdrunheight As Integer, _
 cmdrunleft As Integer, cmdruntop As Integer
Dim cmdrndwidth As Integer, cmdrndheight As Integer, _
 cmdrndleft As Integer, cmdrndtop As Integer
Dim cmdboxofficewidth As Integer, cmdboxofficeheight As Integer, _
 cmdboxofficeleft As Integer, cmdboxofficetop As Integer

Dim cmbvtypleft As Integer, cmbvtypwidth As Integer, _
 cmbvtypheight As Integer, cmbvtyptop As Integer
 
Dim cmbansileft As Integer, cmbansiwidth As Integer, _
 cmbansiheight As Integer, cmbansitop As Integer

Dim hscroll1width As Integer, hscroll1height As Integer, _
 hscroll1left As Integer, hscroll1top As Integer

Dim piccodewidth As Integer, piccodeheight As Integer, _
 piccodeleft As Integer, piccodetop As Integer
Dim pictablewidth As Integer, pictableheight As Integer, _
 pictableleft As Integer, pictabletop As Integer
Dim txtcodewidth As Integer, txtcodeheight As Integer, _
 txtcodeleft As Integer, txtcodetop As Integer
Dim txtinwidth As Integer, txtinheight As Integer, _
 txtinleft As Integer, txtintop As Integer

Dim vscroll1width As Integer, vscroll1height As Integer, _
 vscroll1left As Integer, vscroll1top As Integer
 
Dim lstdtwidth As Integer, lstdtheight As Integer, _
 lstdtleft As Integer, lstdttop As Integer

Private Sub form_resize()

Call resizeall
If cmdback.Visible = True Then
 Call contentspre
End If
Call contents

End Sub
Private Sub form_lostfocus()
'Call resizeall
End Sub
Private Sub form_gotfocus()
'Call resizeall
End Sub
Private Sub form_load()
 q = Chr(34) 'double quote
 qps = "Private Sub " 'often printed to piccode
 
 old6785 = 6785 'default form dimensions
 old4950 = 4950
 
 piccode.ForeColor = vbBlack 'necessary for certain win _
 colorschemes
 picTable.ForeColor = vbBlack
 piccode.BackColor = vbWhite
 picTable.BackColor = vbWhite
 
 bla = 0: black = 0
 blu = 1: blue = 1
 red = 2
 grn = 4: green = 4
 mag = 6: magenta = 6
 whi = 5: white = 5
 yel = 3: yellow = 3
 cyan = 7
 
 Dim ansindx As Integer
 Dim ansilist(0 To 255) As String
 
 For ansindx = 0 To 255
  ansilist(ansindx) = Chr(ansindx) & "   " & ansindx
 Next ansindx
 
 ansilist(0) = "(null)   0"
 ansilist(9) = "(tab)    9"
 ansilist(10) = "(line feed) 10"
 ansilist(13) = "(carriage return)"
 ansilist(32) = "(space)   32"
    
 For ansindx = 0 To 255
  cmbansi.AddItem ansilist(ansindx)
 Next ansindx
 
 cmbvtyp.AddItem "Byte" & "  1 byte" & "  0 to 255"
 cmbvtyp.AddItem "Boolean" & "  2 bytes" & "  True or False"
 cmbvtyp.AddItem "Integer" & "  2 bytes" & " -32,768 to 32,767"
 cmbvtyp.AddItem "Long" & "  4 bytes" & " -2,147,483,648 to 2,147,483,647"
 cmbvtyp.AddItem "Single " & "  4 bytes" & " -3.402823E38 to -1.401298E-45 for - vals"
 cmbvtyp.AddItem " 1.401298E-45 to 3.402823E38 for +"
 cmbvtyp.AddItem "Double" & "  8 bytes" & " -1.79769313486232E308 to -49065645841247E-324"
 cmbvtyp.AddItem " 4.94065645841247E-324 to 1.79769313486232E308"
 cmbvtyp.AddItem "Currency" & "  8 bytes" & " -922,337,203,685,477.5808 to 922,337,203,685,477.5807"
 cmbvtyp.AddItem "Decimal" & "  14 bytes"
 cmbvtyp.AddItem " +/- 79,228,162,514,264,337,593,543,950,335 (no dec point)"
 cmbvtyp.AddItem " +/-7.9228162514264337593543950335 (that's 28 after dec)"
 cmbvtyp.AddItem " smallest non-zero is +/-0.0000000000000000000000000001"
 cmbvtyp.AddItem "Date" & "  8 bytes" & "     Jan 1, 100 to Dec 31, 9999"
 cmbvtyp.AddItem "String" & "  Length of string" & "     1 to approx 65,400"
 cmbvtyp.AddItem "String var-length" & "  10 bytes + str len" & "     0 to approx 2 billion"
 cmbvtyp.AddItem "Variant (nums)" & "  16 bytes" & "  Any numeric value up to a Double's range"
 cmbvtyp.AddItem "Variant (char)" & "  22 bytes + str len" & "     0 to approx 2 billion"
 cmbvtyp.AddItem "User-defined" & "  # req. by elements" & "   The range of each elem = data type range"
 
 lblmaintop = 480
 lblcomtop = 3480
 lblalgtop = 165
 lblqa1top = 1000
 lblqa2top = 500
 lblfuncttop = 4150
 lblqadvtop = 300
 lblqcomtop = 3480
 lblqpagetop = 120
 lblqproptop = 1200
 lblqskiptop = 3750
 lblchaptrtop = 120
 lblbuiltop = 3975
 cmdcd1top = 120
 cmdcd2top = 120
 cmdcd3top = 120
 cmdcd4top = 120
 cmdcd5top = 120
 cmd1lv1top = 3840
 cmd2lv1top = 3840
 cmd1mv1top = 3840
 cmd2mv1top = 3840
 cmda0top = 3960
 cmdnexttop = 3960
 cmdbc9top = 3960
 cmdbacktop = 3960
 cmdboxofficetop = 1055
 cmdrndtop = 750
 txtintop = 3430
 
 lblmainheight = 780
 lblcomheight = 855
 lblalgheight = 165
 lblqa1height = 4575
 lblqa2height = 4575
 lblfunctheight = 165
 lblqadvheight = 165
 lblqcomheight = 255
 lblqpageheight = 255
 lblqpropheight = 255
 lblqskipheight = 195
 lblchaptrheight = 375
 lblbuilheight = 165
 cmdcd1height = 255
 cmdcd2height = 255
 cmdcd3height = 255
 cmdcd4height = 255
 cmdcd5height = 255
 cmdgobackheight = 375
 cmdsimplifyheight = 375
 cmdrndheight = 375
 cmdrunheight = 375
 cmd1lv1height = 375
 cmd2lv1height = 375
 cmd1mv1height = 375
 cmd2mv1height = 375
 cmda0height = 375
 cmdnextheight = 375
 cmdbc9height = 375
 cmdbackheight = 375
 
 lblchaptrleft = 1920
 lblmainleft = 1635
 lblcomleft = 1680
 lblqcomleft = 840
 lblalgleft = 210
 lblqpropleft = 5160
 lblqpageleft = 4680
 lblqskipleft = 75
 lblbuilleft = 90
 lblfunctleft = 120
 cmda0left = 4920
 cmdnextleft = 4920
 cmdbackleft = 4320
 cmdbc9left = 4320
 cmdcd1left = 5160
 cmdcd2left = cmdcd1.Left + 240
 cmdcd3left = cmdcd2.Left + 240
 cmdcd4left = cmdcd3.Left + 240
 cmdcd5left = cmdcd4.Left + 240
 cmdrunleft = 850
 cmdrndleft = 2300
 txtinleft = 255
 
 lblchaptrwidth = 2775
 lblmainwidth = 3540
 lblcomwidth = 2415
 lblqcomwidth = 735
 lblalgwidth = 1825
 lblqpropwidth = 1215
 lblqpagewidth = 475
 lblqskipwidth = 735
 lblbuilwidth = 1695
 lblfunctwidth = 1695
 cmda0width = 650
 cmdnextwidth = 650
 cmdbackwidth = 600
 cmdbc9width = 600
 cmdrunwidth = 615
 cmdrndwidth = 735
 cmdboxofficewidth = 615
 txtinwidth = 255
 
 pictableheight = 1095
 pictablewidth = 2535
 pictableleft = 3840
 pictabletop = 1440
 
 piccodeheight = 1815
 piccodewidth = 3375
 piccodetop = 1440
 piccodeleft = 120
 
 'cmbansileft = 1160
 cmbansileft = 140
 cmbansiwidth = 735
 cmbansitop = 170
 
 cmbvtypleft = 980
 cmbvtyptop = 170

 hscroll1top = 3720
 hscroll1left = 120
 hscroll1width = 1455
 
 vscroll1top = 2520
 vscroll1left = 120
 vscroll1height = 1215
 
 cmdsimplifytop = 500
 cmdsimplifyleft = 500
 cmdsimplifywidth = 855
 cmdgobacktop = 500
 cmdgobackleft = 500
 cmdgobackwidth = 855
End Sub
Private Sub resizeall()
 'called by piccode_Click, contentspre(), contents(), ..
 
 scw = old6785 / Form1.Width  'ratios based on defaults
 sch = old4950 / Form1.Height
 fsr = 2 / (scw + sch) 'font size ratio is the average _
  of new ratios of Form height and width
 fs = 8
 
 lblmain.FontSize = fs * fsr
 lblcom.FontSize = fs * fsr
 piccode.FontSize = fs * fsr
 picTable.FontSize = fs * fsr
 lblqcom.FontSize = fs * fsr
 'lblalg.FontSize = 7 * fsr
 lblqskip.FontSize = 8.25 * fsr
 'lblqadv.FontSize = 7 * fsr
 'lblbuil.FontSize = 7 * fsr
 'lblfunct.FontSize = 7 * fsr
 lblqpage.FontSize = fs * fsr
 lblqprop.FontSize = fs * fsr
 cmdcd1.FontSize = fs * fsr
 cmdcd2.FontSize = fs * fsr
 cmdcd3.FontSize = fs * fsr
 cmdcd4.FontSize = fs * fsr
 cmdcd5.FontSize = fs * fsr
 cmdnext.FontSize = fs * fsr
 cmdback.FontSize = fs * fsr
 cmdA0.FontSize = fs * fsr
 cmdbc9.FontSize = fs * fsr
 cmdrun.FontSize = fs * fsr
 cmdboxoffice.FontSize = fs * fsr
 cmdrnd.FontSize = fs * fsr
 cmbvtyp.FontSize = fs * fsr
 cmbansi.FontSize = fs * fsr
 txtcode.FontSize = fs * fsr
 
 lblmain.Top = 480 / sch
 lblcom.Top = 3480 / sch
 lblalg.Top = 165 / sch
 lblqa1.Top = 1000 / sch
 lblqa2.Top = 500 / sch
 lblfunct.Top = 4150 / sch
 'lblqadv.Top = 20 / sch
 lblqcom.Top = 3480 / sch
 lblqpage.Top = 120 / sch
 lblqprop.Top = 1200 / sch
 lblqskip.Top = 3750 / sch
 lblchaptr.Top = 120 / sch
 lblbuil.Top = 3975 / sch
 lblobj.Top = 4350 / sch
 cmdcopy.Top = cmdcopytop / sch
 
 cmdcd1.Top = 120 / sch
 cmdcd2.Top = 120 / sch
 cmdcd3.Top = 120 / sch
 cmdcd4.Top = 120 / sch
 cmdcd5.Top = 120 / sch
 cmd1lv1.Top = 3840 / sch
 cmd2lv1.Top = 3840 / sch
 cmd1mv1.Top = 3840 / sch
 cmd2mv1.Top = 3840 / sch
 cmdA0.Top = 3960 / sch
 cmdnext.Top = 3960 / sch
 cmdbc9.Top = 3960 / sch
 cmdback.Top = 3960 / sch
 txtcode.Top = txtcodetop / sch
 
 lblmain.Height = 780 / sch
 lblcom.Height = 855 / sch
 lblalg.Height = 165 / sch
 lblqa1.Height = 4575 / sch
 lblqa2.Height = 4575 / sch
 lblfunct.Height = 165 / sch
 lblqadv.Height = 165 / sch
 lblqcom.Height = 255 / sch
 lblqpage.Height = 255 / sch
 lblqprop.Height = 255 / sch
 lblqskip.Height = 195 / sch
 lblchaptr.Height = 375 / sch
 lblbuil.Height = 165 / sch
 lblobj.Height = 165 / sch
 cmdcopy.Height = 255 / sch
 
 cmdcd1.Height = 255 / sch
 cmdcd2.Height = 255 / sch
 cmdcd3.Height = 255 / sch
 cmdcd4.Height = 255 / sch
 cmdcd5.Height = 255 / sch
 cmdgoback.Height = 375 / sch
 cmdsimplify.Height = 375 / sch
 cmdrnd.Height = 375 / sch
 cmdrun.Height = 375 / sch
 cmdboxoffice.Height = 375 / sch
 cmd1lv1.Height = 375 / sch
 cmd2lv1.Height = 375 / sch
 cmd1mv1.Height = 375 / sch
 cmd2mv1.Height = 375 / sch
 cmdA0.Height = 375 / sch
 cmdnext.Height = 375 / sch
 cmdbc9.Height = 375 / sch
 cmdback.Height = 375 / sch
 txtcode.Height = txtcodeheight / sch
 
 lblchaptr.Left = 1920 / scw
 lblmain.Left = 1635 / scw
 lblqadv.Left = 20 / scw
 lblcom.Left = 1680 / scw
 lblqcom.Left = 840 / scw
 lblalg.Left = 210 / scw
 lblqprop.Left = 5160 / scw
 lblqpage.Left = 4680 / scw
 lblqskip.Left = 75 / scw
 lblbuil.Left = 90 / scw
 lblfunct.Left = 110 / scw
 lblobj.Left = 120 / scw
 cmdcopy.Left = cmdcopyleft / scw
 txtcode.Left = txtcodeleft / scw
 
 cmdA0.Left = 4920 / scw
 cmdnext.Left = 4920 / scw
 cmdback.Left = 4320 / scw
 cmdbc9.Left = 4320 / scw
 cmdcd1.Left = 5160 / scw
 cmdcd2.Left = cmdcd1.Left + 240
 cmdcd3.Left = cmdcd2.Left + 240
 cmdcd4.Left = cmdcd3.Left + 240
 cmdcd5.Left = cmdcd4.Left + 240

 lblchaptr.Width = 2775 / scw
 lblmain.Width = 3540 / scw
 lblcom.Width = 2415 / scw
 lblqcom.Width = 735 / scw
 lblalg.Width = 1825 / scw
 lblqprop.Width = 1215 / scw
 lblqpage.Width = 375 / scw
 lblqskip.Width = 735 / scw
 lblbuil.Width = 1695 / scw
 lblfunct.Width = 1695 / scw
 lblobj.Width = 615 / scw
 cmdcopy.Width = 975 / scw
 
 cmdA0.Width = 650 / scw
 cmdnext.Width = 650 / scw
 cmdback.Width = 600 / scw
 cmdbc9.Width = 600 / scw
 cmdrun.Width = 615 / scw
 cmdrnd.Width = 735 / scw
 cmdboxoffice.Width = 615 / scw
 txtcode.Width = txtcodewidth / scw
 
 piccode.Top = piccodetop / sch
 VScroll1.Top = vscroll1top / sch
 HScroll1.Top = hscroll1top / sch
 cmdrun.Top = cmdruntop / sch
 cmdrnd.Top = cmdrndtop / sch
 cmdsimplify.Top = cmdsimplifytop / sch
 cmdgoback.Top = cmdgobacktop / sch
 cmdboxoffice.Top = cmdboxofficetop / sch
 
 piccode.Height = piccodeheight / sch
 VScroll1.Height = vscroll1height / sch
 
 piccode.Left = piccodeleft / scw
 VScroll1.Left = vscroll1left / scw
 HScroll1.Left = hscroll1left / scw
 cmdrun.Left = cmdrunleft / scw
 cmdrnd.Left = cmdrndleft / scw
 cmdsimplify.Left = cmdsimplifyleft / scw
 cmdgoback.Left = cmdgobackleft / scw
 cmdboxoffice.Left = cmdboxofficeleft / scw
 
 piccode.Width = piccodewidth / scw
 
If aIndex <> 11 Then
 picTable.Top = pictabletop / sch
 picTable.Width = pictablewidth / scw
 picTable.Height = pictableheight / sch
 picTable.Left = pictableleft / scw
End If
 
 HScroll1.Width = hscroll1width / scw
 cmdrun.Width = cmdrunwidth / scw
 cmdrnd.Width = cmdrndwidth / scw
 cmdsimplify.Width = cmdsimplifywidth / scw
 cmdgoback.Width = cmdgobackwidth / scw
 cmdboxoffice.Width = cmdboxofficewidth / scw
 'old6785 = Form1.width
 'old4950 = Form1.height
 
 cmbvtyp.Width = cmbvtypwidth / scw 'set in _dropdown and _Validate
 cmbvtyp.Left = cmbvtypleft / scw
 cmbvtyp.Top = cmbvtyptop / sch
 cmbansi.Width = cmbansiwidth / scw
 cmbansi.Left = cmbansileft / scw
 cmbansi.Top = cmbansitop / sch
 
 txtin.Left = txtinleft / scw
 txtin.Width = txtinwidth / scw
 txtin.Top = txtintop / sch
 'Call contentspre
 
End Sub
Private Sub qscolor(qline As String, ByVal qcolor As Byte, ByVal append As Byte)

Select Case qcolor
Case 0
piccode.ForeColor = vbBlack
Case 1
 piccode.ForeColor = vbBlue
Case 2
 piccode.ForeColor = vbRed
Case 3
 piccode.ForeColor = vbYellow
Case 4
 piccode.ForeColor = vbGreen
Case 5
 piccode.ForeColor = vbWhite
Case 6
 piccode.ForeColor = vbMagenta
Case 7
 piccode.ForeColor = vbCyan

End Select

If append = 0 Then
 piccode.Print qline
Else
 piccode.Print qline;
End If

piccode.ForeColor = vbBlack
End Sub
Private Sub cmbvtyp_DropDown()
'cmbvtyp.Left = 170 / scw
cmbvtyp.Width = 5740 / scw
End Sub
Private Sub cmbvtyp_Validate(Cancel As Boolean)
'cmbvtyp.Left = 900 / scw
cmbvtyp.Width = cmbvtypwidth / scw
cmbvtyp.Text = "VarTypes"
End Sub
Private Sub cmbansi_DropDown()
cmbansi.Width = 1435 / scw
cmbvtyp.Left = 1650 / scw
End Sub
Private Sub cmbansi_Validate(Cancel As Boolean)
cmbansi.Width = 865 / scw
cmbansi.Text = "ANSI"
cmbvtyp.Left = 1090 / scw
End Sub
Private Sub lblmain_Click()
If cancelbyte(1) = 0 Then
lblqa1.Visible = True
cancelbyte(1) = 1
lblfunct.Visible = False
lblbuil.Visible = False
lblobj.Visible = False
lblqskip.Visible = False

lblmain.Visible = False
End If
End Sub
Private Sub lblqa1_Click()
lblqa1.Visible = False
lblqa2.Visible = True
End Sub
Private Sub lblqa2_click()
lblqa2.Visible = False
lblmain.Visible = True
lblmain.Caption = "We're ready to start.  Click the Command Button that says 'next'."

End Sub
Private Sub cmdA0_Click()
npage = 1
aIndex = aIndex + 1
backtrue = False
cmdbc9.Visible = True
Call contents
End Sub
Private Sub enableform()
cmdcd3.Enabled = True
cmdcd4.Enabled = True
cmdcd5.Enabled = True
cmdA0.Enabled = True
cmdbc9.Enabled = True

firsttime(1) = 1
firsttime(2) = 1
firsttime(3) = 1
firsttime(4) = 1
firsttime(5) = 1
End Sub
Private Sub clearform()
cmdrun.Visible = False: cmdcd1.Visible = False
cmdcd2.Visible = False: cmdcd4.Visible = False
HScroll1.Visible = False: cmdcd5.Visible = False
VScroll1.Visible = False: cmdcd3.Visible = False
cmd2mv1.Visible = False: cmd1mv1.Visible = False
cmd2lv1.Visible = False: cmd1lv1.Visible = False
txtin.Visible = False
'lblqpage.Visible = False
'lblqprop.Visible = False
cmdgoback.Visible = False
cmdsimplify.Visible = False
End Sub
Private Sub clearqs() 'this sub is called just before _
 a sub assigns values to qs().
 For fn1 = 1 To 99
  qs(fn1) = ""
 Next fn1
End Sub
Private Sub contents()
Select Case aIndex
Case 1
 Call resizeall
 cmdback.Visible = False
 lblqpage.Caption = ""
 lblchaptr.Caption = "Indexing"
 lblmain.Caption = "'Slowly poofed' to me how to use only one next and one back command button."
 Call clearform
 lblcom.Caption = ""
 piccode.Visible = False: picTable.Visible = False
 
Case 2
 lblmain.Caption = "First I created a subroutine to make a general Form 'wipe'."
 piccodetop = 960: piccodeleft = 400: piccodewidth = 3540: piccodeheight = 2050
 Call resizeall
 piccode.Cls: piccode.Visible = True
 picTable.Visible = False
 piccode.Print "Private Sub clearform()"
 piccode.Print " cmdrun.Visible = False: cmdcd1.Visible = False"
 piccode.Print " cmdcd2.Visible = False: cmdcd3.Visible = False"
 piccode.Print " cmdcd4.Visible = False: cmdcd5.Visible = False"
 piccode.Print " HScroll1.Visible = False"
 piccode.Print " VScroll1.Visible = False"
 piccode.Print " piccode.Visible = False"
 piccode.Print " picTable.Visible = False"
 piccode.Print " txtin.Visible = False"
 piccode.Print "End Sub ";: Call qscolor("'cmdcd* are my page buttons", blue, 0)
 lblcom.Caption = ""

Case 3
 piccode.Visible = False: picTable.Visible = False
 lblmain.Caption = "Then I made the new next and back buttons.  aIndex can be accessed by all subs."
 piccodeleft = 3290: piccodewidth = 2090
 pictablewidth = 2170: pictableleft = 990: piccodetop = 960
 pictabletop = 960: pictableheight = 1600: piccodeheight = 1250
 Call resizeall
 lblcom.Caption = "Private Sub contents() is an entire menu, where aIndex is a menu item."
  piccode.Visible = True: picTable.Visible = True
 piccode.Cls
 picTable.Print "Dim aIndex As Integer"
 piccode.Print "Private Sub cmdnext_Click()"
 piccode.Print " aIndex = aIndex + 1"
 piccode.Print " If aIndex > 34 Then"
 piccode.Print "  aIndex = 34: End If"
 piccode.Print " Call contents"
 piccode.Print "End Sub"
 picTable.Cls: picTable.Visible = True
 picTable.Print "Private Sub cmdback_Click()"
 picTable.Print " aIndex = aIndex - 1"
 picTable.Print " If aIndex < 1 Then"
 'picTable.Print "  cmdback.Visible = False"
 picTable.Print "  aIndex = 1"
 picTable.Print " Else"
 picTable.Print "  Call contents"
 picTable.Print " End If"
 picTable.Print "End Sub"
  
Case 4
 lblchaptr.Caption = "Indexing"
 cmdbc9.Caption = "back": cmdA0.Caption = "next"
 lblmain.Caption = "Here is the meat and potatoes.": picTable.Visible = False
 piccode.Cls: piccodeheight = 2650: piccodeleft = 95
 piccodewidth = 6260: piccodetop = 800
 cmdcopy.Visible = False: txtcode.Visible = False
 Call resizeall
 piccode.Visible = True
 piccode.Print "Private Sub contents()"
 piccode.Print " Select Case aIndex"
 piccode.Print " Case 1"
 piccode.Print "  lblchaptr.Caption = " & q & "Indexing" & q
 piccode.Print "  Call clearform"
 piccode.Print "  lblmain.Caption = " & q & "'Slowly poofed' to me how to use only one next and one back command button." & q
 piccode.Print "  lblcom.Caption = " & q & "I will keep my previous buttons." & q
 piccode.Print
 piccode.Print " Case 2"
 piccode.Print "  lblmain.Caption =" & q & "First I created a subroutine to make a general Form 'wipe'." & q
 piccode.Print "  piccodetop = 960: piccodeleft = 400: piccodewidth = 3500: piccodeheight = 2000"
 piccode.Print "  piccode.Cls: piccode.Visible = True"
 piccode.Print "  .."
 lblcom.Caption = "cmdback drops the value of aIndex, and cmdnext increases the value."
  
Case 5
 Dim lintx As Byte  'For-Next counter
 lblchaptr.Caption = "File Access"
 lblmain.Caption = "A simple matter of knowing the code.  I'm using more regular programming conventions now, with 'l(local)str(string)quote'."
 piccodetop = 1115: piccodeleft = 40
 piccodewidth = 4150: piccodeheight = 3400
 pictabletop = 1055: pictableleft = 3490
 pictableheight = 1850: pictablewidth = 3430
 cmdcopyleft = 300: cmdcopytop = 750
 txtcodeleft = 400: txtcodewidth = 5700
 txtcodetop = 1100: txtcodeheight = 2650
 Call resizeall
 If txtcode.Visible = False Then
  cmdcopy.Visible = True
  picTable.Cls
  piccode.Cls: picTable.Visible = False: piccode.Visible = True
  piccode.Print qps & "LoadFavoriteQuotes()"
  piccode.Print " Dim lstrquote(1 To 100) As String"
  piccode.Print " Dim lstrauthor(1 To 100) As String"
  piccode.Print " Dim lbytindx As Byte: lbytindx = 1"
  
  piccode.Print " Open " & q & "c:\favquots.txt" & q & " For Input As #1"
  Call qscolor(" 'Input into 'Buffer 1'", red, 0)
  Call qscolor(" 'EOF(buffer 1)(line below) is a handy 'End Of File' function", red, 0)
  Call qscolor(" Do While ", mag, 1): piccode.Print "Not EOF(1)"
  piccode.Print "  Input #1, lstrquote(lbytindx), lstrauthor(lbytindx)"
  piccode.Print "  Call randomjumble(lstrquote(lbytindx), lstrauthor(lbytindx))"
  Call qscolor("  If ", blue, 1): piccode.Print "LCase(lstrauthor(lbytindx)) = " & q & "einstein" & q & " Then"
  piccode.Print "   Exit Do  ";: Call qscolor("'A quote from Einstein is a good place to stop.", red, 0)
  Call qscolor("  End If", blue, 0)
  piccode.Print "  lbytindx = lbytindx + 1";: Call qscolor("  'increment to next array slot", red, 0)
  Call qscolor(" Loop", mag, 0)
  piccode.Print " Close #1"
  piccode.Print "End Sub"
  lblcom.Caption = "Exit Do is a nice option."
 Else
  Call cmdcopy_Click
 End If
' For lintx = 1 To 6
'  picTable.Print q & qs(lintx) & q & ", " & q & qs(lintx + 6)
' Next lintx
 
Case 6
 lblmain.Caption = ""
 cmdrnd.Visible = False: picTable.Visible = False
 piccodetop = 685: piccodeleft = 130
 piccodewidth = 6190: piccodeheight = 2800
 cmdcopytop = 400
 txtcodeleft = 400: txtcodewidth = 5700
 txtcodetop = 1100: txtcodeheight = 2650
 txtcode.Visible = False
 Call resizeall
 If txtcode.Visible = False Then
  piccode.Cls
  piccode.Visible = True
  piccode.Print qps & "randomjumble(ByRef quote As String, ByRef author As String)"
  piccode.Print " Dim lsngrandom As Single, lstrjumble As String"
  piccode.Print " lstrjumble = " & q & q
  piccode.Print " Dim lintcount As Integer"
  piccode.Print " For lintcount = 1 To Len(author + quote)"
  piccode.Print "  lsngrandom = Rnd * 10  ";: Call qscolor("'deciding variable for lower or uppercase", blue, 0)
  piccode.Print "  If lsngrandom < 5 Then  ";: Call qscolor("'half of the time, we're making lowercase", blue, 0)
  piccode.Print "   lstrjumble = lstrjumble + LCase(Mid(author + quote, lintcount, 1))"
  piccode.Print "  Else"
  piccode.Print "   lstrjumble = lstrjumble + UCase(Mid(author + quote, lintcount, 1))"
  piccode.Print "  End If";: Call qscolor("    'Adds 1 character from author+quote string, randomly upper or lower-cased", blue, 0)
  piccode.Print " Next lintcount";: Call qscolor("   'each time.", blue, 0)
  piccode.Print " picTable.Print q & Right(lstrjumble, Len(quote) & q & " & q & " - " & q & " & Left(lstrjumble, Len(author))"
  piccode.Print "End Sub"
  lblcom.Caption = ""
 Else
  Call cmdcopy_Click
 End If
 
Case 7
 cmdcopy.Visible = False
 lblchaptr.Caption = "File Access"
 lblcom.Caption = ""
 lblmain.Caption = "The strings here are stored in an array in this tutorial."
 cmdcopy.Visible = False
 txtcode.Visible = False
 qs(1) = "Healthy food is tasty."
 qs(2) = "Today is watering day."
 qs(3) = "Did you see that?"
 qs(4) = "The rainbow softened the meadow."
 qs(5) = "Cool."
 qs(6) = "Hey there, chalupa."
 qs(7) = "Captain Veggie"
 qs(8) = "neighbor"
 qs(9) = "the cat"
 qs(10) = "somebody"
 qs(11) = "Einstein"
 qs(12) = "me"
 
 cmdrndtop = 870
 cmdrnd.Visible = True
 piccode.Visible = False: picTable.Visible = False
 picTable.Cls
 piccodeleft = 120: pictableleft = 2000
 pictabletop = 1260: pictableheight = 1050
 pictablewidth = 4100: picTable.Visible = True
 piccodetop = 2400: piccodewidth = 3740
 piccodeheight = 1640: piccode.Visible = True
 Call resizeall
 picTable.Print "Took me 2 hours to get this to work"
 piccode.Cls
 piccode.Print "   contents of " & q & "C:\favquots.txt" & q
 piccode.Print
 For fn1 = 1 To 6
  piccode.Print q & qs(fn1) & q & ", " & q & qs(fn1 + 6) & q
 Next fn1

Case 8
 cmdrnd.Visible = False
 'Form1.Scale (0, 0)-(6675, 4920)
 picTable.Scale (0, 0)-(100, 100)
 picTable.Visible = False: piccode.Visible = False
 'I do this sometimes because certain instances of _
  resizing look glitchy

 lblchaptr.Caption = "Arrays 2"
 lblmain.Caption = "A 2-dimensional Array, like a calendar grid .. we see Varname(First col To Last col, First row To Last row) .. Arrays can have more than 2 dimensions."
 piccodetop = 1355: piccodeleft = 700: piccodewidth = 3200: piccodeheight = 2000
 pictabletop = 2765: pictableleft = 3990: picTable.Cls: pictableheight = 450
 pictablewidth = 1530
 Call resizeall
 piccode.Cls: picTable.Visible = True: piccode.Visible = True
 piccode.Print " Dim i As Integer, j As Integer"
 piccode.Print " Dim arraymult(";: Call qscolor("2 To 3", mag, 1): piccode.Print ", ";: Call qscolor("1 To 5", blue, 1): piccode.Print ") As Integer"
 piccode.Print " 'each 'rectangle' holds an integer"
 piccode.Print
 Call qscolor(" For i = 2 To 3", mag, 0)
 Call qscolor("  For j = 1 To 5", blue, 0)
 piccode.Print "   arraymult(";: Call qscolor("i", mag, 1): piccode.Print ", ";: Call qscolor("j", blue, 1): piccode.Print ") = j * i"
 piccode.Print "   picTable.Print arraymult(";: Call qscolor("i", mag, 1): piccode.Print ", ";: Call qscolor("j", blue, 1): piccode.Print "); " & q & "  " & q & ";"
 Call qscolor("  Next j", blue, 1): piccode.Print ": picTable.Print"
 Call qscolor(" Next i", mag, 0)
 Dim i As Integer, j As Integer
 Dim arraymult(2 To 3, 1 To 5) As Integer
 For i = 2 To 3: For j = 1 To 5: arraymult(i, j) = j * i
  picTable.Print arraymult(i, j); "  ";: Next j: picTable.Print: Next i
 lblcom.Caption = "I am simply loading values quickly into each 'compartment' or 'rectangle'."
 VScroll1.Visible = False

Case 9
 piccodetop = 860: piccodeleft = 120
 piccodeheight = 3500: piccodewidth = 4100
 Call resizeall
 lblchaptr.Caption = "Boolean"
 lblmain.Caption = "Boolean Variables"
 piccode.Cls
 piccode.Print "Dim tall As Boolean"
 piccode.Print "Dim nplayers As Integer"
 piccode.Print " .."
 piccode.Print qps + "heighttest(feet As Integer, inches As Single)"
 piccode.Print " Dim totalinches As Single"
 'piccode.Print " standtall = False  ";: Call qscolor("'Setting this as Default", blue, 0)
 piccode.Print " totalinches = inches + feet * 12"
 piccode.Print " If totalinches > 73 Then"
 piccode.Print "  ";: Call qscolor("tall", blue, 1): piccode.Print " = True"
 piccode.Print "  nplayers = nplayers + 1"
 piccode.Print " Else"
 piccode.Print "  ";: Call qscolor("tall", blue, 1): piccode.Print " = False"
 piccode.Print " End If"
 piccode.Print "End Sub"
 piccode.Print " .."
 piccode.Print qps + "uniform(tall, team)"
 piccode.Print " If ";: Call qscolor("tall", blue, 1): piccode.Print " Then"
 piccode.Print "  lblusize = " & q & "XLarge" & q
 piccode.Print " End If"
 piccode.Print " .."

 picTable.BackColor = vbWhite: picTable.ForeColor = vbBlack
 picTable.Visible = False
 lblcom.Caption = ""

Case 10
 lblchaptr.Caption = "Graphics"
 picTable.BackColor = vbBlack
 pictabletop = 1160: pictableheight = 1400
 pictablewidth = 1800: pictableleft = 4320
 piccodetop = 1160
 piccodeheight = 3000
 piccodeleft = 220
 piccodewidth = 3920
 Call resizeall
 lblmain.Caption = q & "Wohoo!" & q & "  This is what I always look forward to learning in any programming language.  There are at least 3 methods for changing color."
 piccode.Cls
 piccode.Print " .."
 piccode.Print " Dim pi As Single: pi = 3.14159"
 piccode.Print " Dim x As Single"
 piccode.Print " picTable.backcolor = RGB(100, 95, 120)"
 piccode.Print " picTable.Cls"
 piccode.Print " picTable.Scale (-20, 12)-(120, -2)"
 piccode.Print
 piccode.Print " For x = 0 to 100 Step 0.2"
 Call qscolor(" 'plot about 500 points per function", blue, 0)
 piccode.Print "  picTable.PSet (5*x, sin(x/3)), RGB(199, 0, Int(255 - x*2))"
 piccode.Print "  picTable.Pset (x, Sin(x) + 3), vbGreen"
 piccode.Print "  picTable.Pset (x, Cos(x) + 2), vbYellow"
 piccode.Print " Next x"
 
 Dim x1 As Single
 picTable.BackColor = RGB(100, 95, 120)
 picTable.Cls
 picTable.Scale (-20, 12)-(120, -2)
 For x1 = 0 To 100 Step 0.2
 picTable.PSet (5 * x1, Sin(x1 / 3)), RGB(199, 0, Int(255 - x1 * 2))
 picTable.PSet (x1, Sin(x1) + 3), vbGreen
 picTable.PSet (x1, Cos(x1) + 2), vbYellow
 Next x1
 picTable.Visible = True
 piccode.Visible = True
 cmdrun.Visible = False

Case 11
 picTable.FillColor = vbBlack
 picTable.BackColor = vbWhite: picTable.ForeColor = vbBlack
 VScroll1.Visible = False: cmdrun.Visible = False
 piccode.Visible = False: piccode.Cls
 picTable.FontSize = 8
 picTable.Top = 1160 / sch: picTable.Left = 400 / scw
 picTable.Width = 5700: picTable.Height = 2600
 picTable.Scale (0, 0)-(70, 28)
'piccodetop = 1160: piccodeleft = 200
'piccodewidth = 2600: piccodeheight = 1400
 'Call resizeall
 picTable.Cls
 lblmain.Caption = "Bar, Bar Filled..  To use any of the fillstyles 1 to 7 you must use , , B"
 picTable.CurrentX = 0
 picTable.CurrentY = 0.5
 picTable.Print "(x1, y1)"; Tab(39); "(x1,y1)"
 picTable.Print: picTable.Print
 picTable.CurrentY = 5.8
 picTable.Print Tab(26); "(x2, y2)"; Tab(62); "(x2, y2)"
 picTable.Line (3, 3)-(26, 5.5), , B
 picTable.Line (38.3, 3)-(59.4, 5.5), , BF
 picTable.Print: picTable.Print
 picTable.Print "picTable.Line (x1, y1) - (x2, y2), , B"; Tab(39); "picTable.Line (x1, y1) - (x2, y2), , BF"
 picTable.Print: picTable.CurrentX = 38.8
 picTable.CurrentY = 15
 picTable.Print "picTable.DrawStyle = "
 Dim CY As Single
 For fn1 = 0 To 4
 CY = (fn1 + 1) * 2.1 + 16
 picTable.CurrentX = 36: picTable.CurrentY = CY - 1
 picTable.Print fn1
 picTable.DrawStyle = fn1
 picTable.Line (39, CY)-(60, CY)
 Next fn1
 picTable.DrawStyle = 0

 picTable.CurrentX = 0
 picTable.CurrentY = 13.5: picTable.Print "picTable.Fillstyle ="

 picTable.Print
 For fn1 = 1 To 4
 CY = (fn1 + 1) * 2.9 + 10.8
 picTable.FillStyle = fn1 - 1
 picTable.CurrentX = 0: picTable.CurrentY = CY + 0.2
 picTable.Print fn1 - 1
 picTable.Line (3, CY)-(13, CY + 2.4), , B
 picTable.CurrentX = 16: picTable.CurrentY = CY + 0.2
 picTable.Print fn1 + 3
 picTable.FillStyle = fn1 + 3
 picTable.Line (19, CY)-(31, CY + 2.4), , B
 Next fn1
 picTable.FillStyle = 1

Case 12
 lblmain.Caption = "Positioning Text"
 picTable.Cls: piccode.Cls: cmdruntop = 870
 cmdrun.Visible = True
 piccodeleft = 440: piccodetop = 1300
 piccodeheight = 2200: piccodewidth = 3960
 piccode.Visible = True
 pictableleft = 4510: pictabletop = 1300
 pictablewidth = 1830: pictableheight = 1830
 Call resizeall
 Call clearqs: VScroll1.Min = 1: VScroll1.Max = 14
 VScroll1.Value = 1
 VScroll1.Visible = True

 qs(1) = "Private Sub cmdrun_Click"
 qs(2) = " Select Case aIndex"
 qs(3) = " .."
 qs(4) = " Case 12"
 qs(5) = " Dim qx As Single, qy As Single"
 qs(6) = " picTable.Cls"
 qs(7) = " picTable.Scale (-1.8, 11) - (11, -1.8)"
 qs(8) = " picTable.Line (0, 0) - (0, 10)"
 qs(9) = " picTable.Line (0, 0) - (10, 0)"
'qs(10) = " picTable.currentY = 3 - picTable.TextHeight(" & q & "y=3" & q & ") / 2"
 qs(11) = " For fn1 = 0 To 10  'modular all-purpose Integer"
 qs(12) = "  picTable.Line (n1, .2)-(n1, -0.2), vbBlue"
 qs(13) = "  picTable.CurrentX = n1 - picTable.TextWidth(n1)"
 qs(14) = "  picTable.CurrentY = -.3"
 qs(15) = "  picTable.Print n1"
 qs(16) = "  qy = Int(2.8 * n1 - 10)"
 qs(17) = "  If qy >= 0 and qy <= 10 Then"
 qs(18) = "   picTable.Line (-.3, qy) - (.2, qy), vbMagenta"
 qs(19) = "   picTable.CurrentX = -1.4"
 qs(20) = "   picTable.CurrentY = qy - picTable.TextHeight(qy) / 2"
 qs(21) = "   picTable.Print qy"
 qs(22) = "   End If: Next n1"
 qs(23) = "End Sub"
 Call VScommon

Case 13 'cmdrun_Click sel case aIndex
 'lblmain.Caption = "Arc-circle option and the FillStyle method"
 lblmain.Caption = ""
 picTable.Cls: piccode.Cls: cmdruntop = 760
 piccodeleft = 220: piccodetop = 1160
 piccodeheight = 2800: piccodewidth = 3560
 pictableleft = 4100: pictabletop = 1160
 pictablewidth = 1800: pictableheight = 1400
 Call resizeall
 piccode.Print "Private Sub cmdrun_Click()"
 piccode.Print " Dim c As Single, a As Single, b As Single"
 piccode.Print " picTable.cls"
 piccode.Print " picTableleft = 4100: picTabletop = 1160"
 piccode.Print " picTablewidth = 1800: picTableheight = 1400"
 piccode.Print " picTable.Scale (1, 1)-(60, 60)"
 piccode.Print " c = 2 * 3.14159"
 piccode.Print " a = .0000001"
 piccode.Print " b = .4  ";: Call qscolor("'40%", blue, 0)
 piccode.Print " picTable.Scale (1, 1)-(60, 60)"
 piccode.Print " picTable.FillStyle = 5"
 piccode.Print " picTable.FillColor = vbBlue"
 piccode.Print " picTable.Circle (30, 30), 20, , -a * c, -b * c"
 piccode.Print " picTable.FillStyle = 1";: Call qscolor("  'reset", blue, 0)
 piccode.Print "End Sub"
 VScroll1.Visible = False
 piccode.Visible = True: picTable.Visible = True
 cmdrun.Visible = True
 lblcom.Caption = ""

Case 14 'if change, remembr picTable_Click and VScommon Sel Case aIndex
 txtcode.Visible = False
 lblmain.Caption = "This will give you more of an idea about the RGB function and PSET method"
 VScroll1.Value = 2 'vscroll1.value = 1 later to _
 make text appear in piccode as VScroll1 changes.
 lblchaptr.Caption = "Graphics"
 piccodetop = 960: piccodeleft = 460
 piccodeheight = 2000: piccodewidth = 5200
 picTable.Scale (0, 0)-(100, 100)
 picTable.Cls
 pictabletop = 3110: pictableheight = 1200: pictableleft = 1800
 pictablewidth = 1600
 
 Call resizeall
 txtin.Visible = False
 Call clearqs
 lblcom.Caption = ""
 qs(1) = "Dim x(40) As Single, y(40) As Single, r(40) As Single"
 qs(2) = "Dim incred(40) As Single, incgrn(40) As Single, incblu(40) As Single"
 qs(3) = "Dim ared(40) As Integer, agrn(40) As Integer, ablu(40) As Integer"
 qs(4) = "Dim colorshift As Integer, csi As Integer, satur As Integer, scl As Integer"
 qs(6) = ""
 qs(5) = "Dim n As Byte"
 qs(7) = "Private Sub picTable_Click()"
 qs(8) = "Select Case aIndex"
 qs(9) = "Case 7: Dim z, az, angle, ar, pi, wx As Single"
 qs(10) = ""
 qs(11) = " pi = 3.141593: picTable.Cls"
 qs(12) = " picTable.Scale (-200, 200)-(200, -200)"
 qs(13) = " satur = 90 'saturation"
 qs(14) = ""
 qs(15) = " For n = 1 To 40"
 qs(16) = ""
 qs(17) = "  ared(n) = 255 * Rnd"
 qs(18) = "  agrn(n) = ared(n) + Rnd * satur - satur / 2.9"
 qs(19) = "  ablu(n) = agrn(n) + Rnd * satur - satur / 2.9"
 qs(20) = ""
 qs(21) = "  scl = 230 'larger value = more sparse cylinders"
 qs(22) = "  x(n) = Rnd * scl - scl / 2"
 qs(23) = "  y(n) = Rnd * scl - scl / 2"
 qs(24) = ""
 qs(25) = "  r(n) = 7 * Rnd + 10 'varying radii"
 qs(26) = ""
 qs(27) = "  csi = 2 'colorshift intensity"
 qs(28) = ""
 qs(29) = "  colorshift = Rnd * satur - satur / 2"
 qs(30) = "  incred(n) = colorshift"
 qs(31) = "  incgrn(n) = colorshift"
 qs(32) = "  incblu(n) = colorshift"
 qs(33) = " Next n"
 qs(34) = ""
 qs(35) = " For z = -70 To 400 Step 25"
 qs(36) = "  ar = 1250 / (1250 - z)"
 qs(37) = "  For n = 1 To 40"
 qs(38) = ""
 qs(39) = "   Call straightencolor(ared(n))"
 qs(40) = "   Call straightencolor(agrn(n))"
 qs(41) = "   Call straightencolor(ablu(n))"
 qs(42) = ""
 qs(43) = "   For a = 0 To 2 * pi Step 2 * pi / 9"
 qs(44) = "    picTable.PSet (ar * (r(n) * Sin(a) + x(n)), _"
 qs(45) = "    ar * (r(n) * Cos(a) + y(n))), _"
 qs(46) = "    RGB(ared(n), agrn(n), ablu(n))"
 qs(47) = "   Next a"
 qs(48) = ""
 qs(49) = "   ared(n) = ared(n) - incred(n)"
 qs(51) = "   agrn(n) = agrn(n) - incgrn(n)"
 qs(51) = "   ablu(n) = ablu(n) - incblu(n)"
 qs(52) = ""
 qs(53) = "  Next n"
 qs(54) = " Next z"
 qs(55) = "End Select"
 qs(56) = "End Sub"
 qs(57) = "Private Sub straightencolor(ByRef color As Integer)"
 qs(58) = " If color < 0 Then"
 qs(59) = "  color = color + 255: End If"
 qs(60) = " If color > 255 Then"
 qs(61) = "  color = color - 255: End If"
 qs(62) = "End Sub"

 VScroll1.Min = 1: VScroll1.Value = 1: VScroll1.Max = 57
 VScroll1.Visible = True: piccode.Visible = True
 picTable.Visible = True: picTable.Print "picTable"
 cmdrun.Visible = False

Case 15
 txtcode.Visible = False
 lblchaptr.Caption = "Text Filtering"
 subj = 7: npage = 3  'txtin_Change and txtin_KeyPress
 lblmain.Caption = "Here is how I axe non-numeric keystrokes and the Fifty thou": Call clearqs
 piccodeleft = 410: piccodewidth = 3700
 piccodetop = 1070: piccodeheight = 2300
 piccode.Visible = True: picTable.Visible = False
 VScroll1.Value = 1: VScroll1.Min = 1: VScroll1.Max = 17
 VScroll1.Visible = True: txtinwidth = 640
 txtinleft = 500: txtintop = 3380
 Call resizeall
 qs(1) = "Dim prevtext As String"
 qs(2) = "Private Sub txtin_change()"
 qs(3) = "Dim asc1 As String"
 qs(4) = "Dim lentxtin As Byte"
 qs(5) = " lentxtin = Len(txtin.Text)"
 qs(6) = " For fn1 = 1 To lentxtin  'no non-numeric keystrokes"
 qs(7) = "  asc1 = Asc(Mid(txtin.Text, fn1, 1))"
 qs(8) = "  If asc1 > 57 Or asc1 < 48 Then"
 qs(9) = "   txtin.Text = prevtext: Exit For: End If"
 qs(10) = " Next fn1"
 qs(11) = " If Val(txtin.Text) > 50000 Then"
 qs(12) = "  txtin.Text = 50000: End If"
 qs(13) = " If Val(txtin.Text) > 49 Then"
 qs(14) = "  'cmdrun would calculate for each keystroke"
 qs(15) = "  Call cmdrun_Click: End If"
 qs(16) = "End Sub"
 qs(17) = "Private Sub txtin_KeyPress(k As Integer)"
 qs(18) = " 'txtin_Change executes before _KeyPress."
 qs(19) = " 'txtin_Change uses prevtext if keystroke"
 qs(20) = " 'is non-numeric"
 qs(21) = " prevtext = txtin.Text"
 qs(22) = "End Sub"
 qs(23) = "Private Sub cmdrun_Click()"
 qs(24) = "Dim balance As Single, numYears As Integer"
 qs(25) = " If txtin.Text = "" Or txtin.Text < 50 Then"
 qs(26) = "  txtin.Text = 50: End If"
 qs(27) = " .."
 'qs(28) = " balance = Val(txtin.Text)"
 'qs(29) = " numYears = 0"
 'qs(30) = " Do While balance < 1000000"
 'qs(31) = "  balance = balance + 0.07 * balance"
 'qs(32) = "  numYears = numYears + 1: Loop"
 'qs(33) = ""
 'qs(34) = " picTable.Cls"
 'qs(35) = " picTable.Print " & q & "About " & q & "; numYears; " & q & "years." & q
 'qs(36) = "End Sub"
 Call VScommon
 
Case 16
 lblmain.Caption = "Ya can copy-paste this ya know.": lblchaptr.Caption = "Text Filtering"
 lblcom.Caption = "I'll probably add a button to several previous sections to let u copy-paste"
 txtcodeleft = 1020: txtcodewidth = 3800
 txtcodetop = 900: txtcodeheight = 2500
 Call resizeall
 txtcode.Visible = True: txtin.Visible = False
 VScroll1.Visible = False
 piccode.Visible = False
 picTable.Visible = False
 qs(1) = "Dim prevtext As String"
 qs(2) = "Private Sub txtin_change()"
 qs(3) = "Dim asc1 As String"
 qs(4) = "Dim lentxtin As Byte"
 qs(5) = " lentxtin = Len(txtin.Text)"
 qs(6) = " For fn1 = 1 To lentxtin"
 qs(7) = "  asc1 = Asc(Mid(txtin.Text, fn1, 1))"
 qs(8) = "  If asc1 > 57 Or asc1 < 48 Then"
 qs(9) = "   txtin.Text = prevtext: Exit For: End If"
 qs(10) = " Next fn1"
 qs(11) = " If Val(txtin.Text) > 50000 Then"
 qs(12) = "  txtin.Text = 50000: End If"
 qs(13) = " If Val(txtin.Text) > 49 Then"
 qs(14) = "  Call cmdrun_Click: End If"
 qs(15) = "End Sub"
 qs(16) = "Private Sub txtin_KeyPress(k As Integer)"
 qs(17) = " prevtext = txtin.Text"
 qs(18) = "End Sub"

 txtcode.Text = ""
 For mv1 = 1 To 18
  txtcode.Text = txtcode.Text & qs(mv1) & vbCrLf
 Next mv1
 
Case 17
 lblmain.Caption = "I had a lot of fun with this, but man am I slow.": lblchaptr.Caption = "How I Resize"
 lblcom.Caption = ""
 txtcodeleft = 1220: txtcodewidth = 4200
 txtcodetop = 1160: txtcodeheight = 2500
 Call resizeall
 txtcode.Visible = True
 'VScroll1.Visible = False
 'piccode.Visible = False
 'picTable.Visible = False
 
 qs(1) = "Dim defaultw As Integer, defaulth As Integer"
 qs(2) = ""
 qs(3) = "'width & height ratios, fontsize ratio, fontsize"
 qs(4) = "Dim scw As Single, sch As Single, _"
 qs(5) = " fsr As Single, fs As Single"
 qs(6) = ""
 qs(7) = "Dim obj1width As Integer, obj1left As Integer, _"
 qs(8) = " obj1height As Integer, obj1top As Integer"
 qs(9) = "Dim obj2width As Integer, obj2left As Integer, _"
 qs(10) = " obj2height As Integer, obj2top As Integer"
 qs(11) = ""
 qs(12) = "Private Sub Form_Load()"
 qs(13) = " defaultw = 6785 'default form dimensions"
 qs(14) = " defaulth = 4950 'meant only to be read"
 qs(15) = " fs = 8 'all-purpose default font size"
 qs(16) = "End Sub"
 qs(17) = ""
 qs(18) = "Private Sub form_resize()"
 qs(19) = " Call resizeall"
 qs(20) = " Call contents"
 'qs(21) = " 'I actually say Call ResizeAll because my code is _"
 'qs(22) = "  so screwed up, if I call contents, pages reset, _"
 'qs(23) = "  it gets confusing because I'm a poopy programmer. _"
 'qs(24) = "  Call ResizeAll is effective but sometimes there's _"
 'qs(25) = "  picbox content clipping.  I may figure it all out _"
 'qs(26) = "  in a future release."
 qs(27) = "End Sub"
 qs(28) = ""
 qs(29) = "Private Sub ResizeAll()"
 qs(30) = " scw = defaultw / Form1.Width"
 qs(31) = " sch = defaulth / Form1.Height"
 qs(32) = " fsr = 2 / (scw + sch) 'fontsize ratio"
 qs(33) = " obj1.FontSize = fs * fsr"
 qs(34) = " obj2.FontSize = fs * fsr"
 qs(35) = " .."
 qs(36) = " obj1.width = obj1width / scw"
 qs(37) = " obj1.left = obj1left / scw"
 qs(38) = " obj1.height = obj1height / sch"
 qs(39) = " obj1.top = obj1top / sch"
 qs(40) = " .."
 qs(41) = "End Sub"
 qs(42) = ""
 qs(43) = "Private Sub contents()"
 qs(44) = " Select Case aIndex"
 qs(45) = " Case 1"
 qs(46) = "  obj1top = .."
 qs(47) = "  obj1height = .."
 qs(48) = "  obj2top .."
 qs(49) = "  Call ResizeAll"
 qs(50) = ""
 qs(51) = "  obj1.Caption = .. 'labels Caption"
 qs(52) = "  obj2.Print .. 'picboxes Print"
 qs(53) = ""
 qs(54) = " Case 2"
 qs(55) = "  obj1top = .."
 qs(56) = "  obj2top = .."
 qs(57) = ""
 qs(58) = "  Call ResizeAll"
 qs(59) = "  obj1.Caption = .."
 qs(60) = "  obj2.Print"
 qs(61) = " End Select"
 qs(62) = "End Sub"
 txtcode.Text = ""
 For mv1 = 1 To 20
  txtcode.Text = txtcode.Text & qs(mv1) & vbCrLf
 Next mv1
 For mv1 = 27 To 62
  txtcode.Text = txtcode.Text & qs(mv1) & vbCrLf
 Next mv1

Case Is > 17
txtcode.Visible = False
txtin.Visible = False
Call clearform
picTable.Scale (0, 0)-(100, 100)
aIndex = 18
Call resizeall
lblmain.Caption = " .. more to come .. if you have comments, email fluoats@hotmail.com"
lblcom.Caption = ""
picTable.Visible = False
piccode.Visible = False

Case 21
txtcodeleft = 220: txtcodewidth = 4800
txtcodetop = 960: txtcodeheight = 2500
Call resizeall
txtcode.Visible = True
qs(1) = "Dim correct As Byte"
qs(2) = "Dim tries As Byte"
qs(3) = "Dim stat As Byte"
qs(4) = ".."
qs(5) = "Private Sub form_load()"
qs(6) = " For fn1 = 0 To 5"
qs(7) = "  imgdrag(fn1).DragIcon = imgdrag(fn1).Picture"
qs(8) = " Next fn1"
qs(9) = "End Sub"
qs(10) = ""
qs(11) = "Private Sub picdrop_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)"
qs(12) = " tries = tries + 1"
qs(13) = " If tries = 1 Then"
qs(14) = "  Form1.Cls"
qs(15) = "  lblcom.Caption = "": End If"
qs(16) = " If Index = stat Then"
qs(17) = "  correct = correct + 1"
qs(18) = "  picdrop(Index).Picture = Source.Picture"
qs(19) = "  If correct = 5 And tries = 5 Then"
qs(20) = "   lblcom.Caption = " & q & "Ya think?" & q & ": End If"
qs(21) = "  Call printresults"
qs(22) = " Else"
qs(23) = "  Source.Visible = True: Call printresults: End If"
qs(24) = " If tries > 5 Then"
qs(25) = "  Print " & q & "Game Over  " & q & ";: Call shuffle"
qs(26) = "  lblcom.Caption = " & q & "Starting ova" & q
qs(27) = "  If correct = 6 Then"
qs(28) = "   Print " & q & "You're the junk" & q & ": End If"
qs(29) = "  Call printresults: tries = 0: correct = 0"
qs(30) = " End If"
qs(31) = " lbldrop.Caption = Left(lbldrop.Caption, 11) & Index"
qs(32) = "End Sub"
qs(33) = "Private Sub cmdstartover_Click()"
qs(34) = " Call shuffle"
qs(35) = " Call reset"
qs(36) = " lbldrop.Caption = " & q & "drop index 0" & q
qs(37) = "End Sub"
qs(38) = "Private Sub reset()"
qs(39) = " correct = 0"
qs(40) = " tries = 0"
qs(41) = " stat = 0"
qs(42) = " Index = 0"
qs(43) = " Call printresults"
qs(44) = "End Sub"
qs(45) = "Private Sub printresults()"
qs(46) = " lblcorr.Caption = Left(lblcorr.Caption, 8) & correct"
qs(47) = " lbldrag.Caption = Left(lbldrag.Caption, 11) & stat"
qs(48) = " lbltries.Caption = Left(lbltries.Caption, 6) & tries"
qs(49) = "End Sub"
qs(50) = "Private Sub imgdrag_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)"
qs(51) = " Source.Visible = False"
qs(52) = " stat = Index"
qs(53) = "End Sub"
qs(54) = "Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)"
qs(55) = " Source.Visible = True"
qs(56) = "End Sub"
qs(57) = "Private Sub shuffle()"
qs(58) = " Randomize"
qs(59) = " Dim swap As Integer"
qs(60) = " Dim rand As Byte"
qs(61) = " For fn1 = 0 To 5"
qs(62) = "  rand = Int(Rnd * 6)"
qs(63) = "  swap = imgdrag(fn1)left"
qs(64) = "  imgdrag(fn1)left = imgdrag(rand)left"
qs(65) = "  imgdrag(rand)left = swap"
qs(66) = "  imgdrag(fn1).Visible = True"
qs(67) = "  picdrop(fn1).Picture = picswap.Picture: Next fn1"
qs(68) = "End Sub"
qs(69) = "Private Sub lblsign_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)"
qs(70) = " Source.Visible = True"
qs(71) = "End Sub"
qs(72) = "Private Sub imgdrag_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)"
qs(73) = " Source.Visible = True"
qs(74) = "End Sub"
qs(75) = "Private Sub lblcorr_DragDrop(Source As Control, X As Single, Y As Single)"
qs(76) = " Source.Visible = True"
qs(77) = "End Sub"
qs(78) = "Private Sub lbldrop_DragDrop(Source As Control, X As Single, Y As Single)"
qs(79) = " Source.Visible = True"
qs(80) = "End Sub"
qs(81) = "Private Sub lbldrag_DragDrop(Source As Control, X As Single, Y As Single)"
qs(82) = " Source.Visible = True"
qs(83) = "End Sub"
qs(84) = "Private Sub lblcom_DragDrop(Source As Control, X As Single, Y As Single)"
qs(85) = " Source.Visible = True"
qs(86) = "End Sub"
qs(87) = "Private Sub lbltries_DragDrop(Source As Control, X As Single, Y As Single)"
qs(88) = " Source.Visible = True"
qs(89) = "End Sub"
qs(90) = "Private Sub cmdstartover_DragDrop(Source As Control, X As Single, Y As Single)"
qs(91) = " Source.Visible = True"
qs(92) = "End Sub"

txtcode.Text = qs(1)
For mv1 = 2 To 92 'stealing a currently-unused modular variable
 txtcode.Text = txtcode.Text & qs(mv1) & vbCrLf
Next mv1
VScroll1.Visible = False
piccode.Visible = False
picTable.Visible = False
End Select

End Sub
Private Sub cmdbc9_Click()
npage = 1
aIndex = aIndex - 1
backtrue = True
If aIndex < 1 Then
 cmdbc9.Visible = False
 subj = 7
 npage = 1
 pIndex = 17  'contentspre select case
 cmdbc9.Visible = False
 cmdback.Visible = True
 cmdA0.Visible = True
 cmdcd1.Visible = True
 cmdcd2.Visible = True
 cmdcd3.Visible = True
 cmdcd4.Visible = True
 cmdcd5.Visible = False
 Call contentspre
Else
 Call contents
End If
End Sub
Private Sub cmdnext_Click()
npage = 1
If pIndex < 2 Then
 lblqa1.Visible = False
 lblqa2.Visible = False
 lblmain.Visible = True
End If
pIndex = pIndex + 1
backtrue = False
Call contentspre
End Sub
Private Sub cmdback_click()
npage = 1
pIndex = pIndex - 1
If pIndex < 1 Then
 pIndex = 0: End If
backtrue = True
Call contentspre
End Sub
Private Sub contentspre()
Select Case pIndex
Case 0
lblfunct.Visible = True
lblbuil.Visible = True
lblobj.Visible = True
lblqskip.Visible = True
cmdback.Visible = False
picTable.Visible = False
lblqcom.Visible = False
lblqprop.Visible = False
lblmain.Caption = "Hi there!"
lblcom.Caption = ""
Case 1
cmdback.Visible = True
cancelbyte(1) = 1
lblfunct.Visible = False
lblbuil.Visible = False
lblobj.Visible = False
lblqskip.Visible = False
cmbansi.Visible = False
cmbvtyp.Visible = False
lblqprop.Visible = True
lblqcom.Visible = True
picTable.Visible = True
lblqcom.Visible = True
lblchaptr.Caption = "Intro"
lblmain.Caption = "Visual Basic uses 'objects'.  Objects have 'properties'.  In the white rectangle (known as a 'picture box'), I list a few properties of a text 'label' on this Form."
lblcom.Caption = "A label has 37 properties!  As of VB ver. 6 anyway,  BackColor and ForeColor are also among the properties."
'lblcom.Visible = True
picTable.Cls
picTable.Print "(name)"; Tab(12); "lblchaptr"
picTable.Print "Caption"; Tab(12); "Intro"
picTable.Print "Visible"; Tab(12); "True"
picTable.Print "."
picTable.Print "."

Case 2
lblmain.Caption = "It is really really helpful to give object names meaningful prefixes to keep track of potentially a TON of different code pieces."
picTable.Cls
picTable.Print "(name)"; Tab(12); "lblmain"
picTable.Print "Caption"; Tab(12); lblmain.Caption
picTable.Print "Visible"; Tab(12); "True"
picTable.Print "."
picTable.Print "."
lblcom.Caption = "I selected a 'picture box' to list those properties because I like how it looks."
lblqprop.Visible = True
lblqprop.Caption = "Properties"
lblcom.Caption = "Is there such a thing as a ton of code?"
picTable.Visible = True

Case 3
lblqprop.Visible = False
picTable.Visible = False
lblmain.Caption = "The commandbutton with 'next' as its caption has a 'click event' that activates code!"
lblcom.Caption = "Yaay."
lblchaptr.Caption = "Intro"
piccode.Visible = False
cmdnext.Caption = "next"

Case 4
cmdnext.Caption = "hello"
lblqprop.Visible = True
lblchaptr.Caption = "Code"
picTable.Visible = True: piccode.Visible = True
picTable.Visible = True: cmdrun.Visible = False
picTable.Cls
lblcom.Caption = "Remember 'If this is your first time, click HERE'?  You can assign click events to controls other than command buttons."

If Not backtrue Then
 lblmain.Caption = "cmdnext is the name I gave to the command button you just clicked.  The _Click means .. ah .. do I really need to tell you?"
 piccode.Cls
 piccode.Print "Private Sub cmdnext_Click()"
 piccode.Print " lblchaptr.Caption = " & q & "Code" & q
 piccode.Print " lblcom.Caption = " & q & lblcom.Caption
 piccode.Print " cmdnext.Caption = " & q & "hello" & q
 piccode.Print " lblmain.Caption = " & q & lblmain.Caption
 piccode.Print " ..."
 piccode.Print "End Sub"
 picTable.Print "(name)"; Tab(12); "cmdnext "
 picTable.Print "Caption"; Tab(12); "next  (my default)"
Else
 lblmain.Caption = "cmdback is the name I gave to the command button you just clicked."
 piccode.Cls
 piccode.Print "Private Sub cmdback_Click()"
 piccode.Print " lblchaptr.Caption = " & q & "Code" & q
 piccode.Print " lblcom.Caption = " & q & lblcom.Caption
 piccode.Print " cmdback.Caption = " & q & "hello" & q
 piccode.Print " lblmain.Caption = " & q & lblmain.Caption
 piccode.Print " ..."
 piccode.Print "End Sub"
 picTable.Print "(name)"; Tab(12); "cmdback"
 picTable.Print "Caption"; Tab(12); "back"
End If

Case 5
piccode.Enabled = False
If Not backtrue Then
 cmdnext.Caption = "next"
 lblmain.Caption = "My command buttons, combined with other code (that I'm hiding), allow me to select content for each 'page' in this tutorial."
 lblcom.Caption = "Next page a scrollbar i named HScroll1 becomes visible using this:  HScroll1.Visible = True"
 piccode.Visible = True
 piccode.Cls
 piccode.Print "Private Sub cmdnext_Click()"
 piccode.Print "  ..": piccode.Print "  ..": piccode.Print "  .."
 piccode.Print "  lblmain.Caption = " & q & lblmain.Caption
 piccode.Print "  lblcom.Caption = " & q & lblcom.Caption
 piccode.Print " .."
 piccode.Print "End Sub"
 picTable.Cls
 picTable.Print "(name)"; Tab(12); "cmdnext"
 picTable.Print "Caption"; Tab(12); "next"
 picTable.Visible = True

Else
 HScroll1.Visible = False
 lblchaptr.FontSize = 14
 lblmain.Caption = "Ooo you just found an old page .. think I'll keep it for kicks.  This is some of the code I was hiding."
 lblqcom.Caption = "comment:"
 lblcom.Caption = "Shown is code for the back button you just clicked."
 piccodetop = 1440: piccodeheight = 1815
                                     piccodeleft = 120
 piccode.Visible = True
 piccode.Cls
 piccode.Print "Private Sub cmdback_Click()"
 piccode.Print " aIndex = aIndex - 1"
 piccode.Print " Select Case aIndex"
 piccode.Print " Case 5"
 piccode.Print "  HScroll1.Visible = False"
 piccode.Print "  lblmain.Caption = " & q & "Ooo you just found an old page .. think I'll keep "
 piccode.Print "  lblqcom.Caption = " & q & "comment:" & q
 piccode.Print "  lblcom.Caption = " & q & "Shown is code for the back button"
 piccode.Print " End Select: End Sub"
 picTable.Cls
 picTable.Print "(name)"; Tab(12); "cmdback"
 picTable.Print "Caption"; Tab(12); "back"
 picTable.Visible = True
End If

Case 6 'Horizontal Scrollbar intro
 lblchaptr.Caption = "Code"
 subj = 0
 piccode.Enabled = True
 cmdnext.Caption = "next"
 Call clearform
 lblqpage.Visible = False
 piccode.Visible = True
 lblalg.Visible = False
 lblalg.Caption = "Algebra Reference"
 'lblqprop.Caption = "Properties"
 hscroll1left = 120
 hscroll1top = 3720
 hscroll1width = 1455
 lblmain.Caption = "Try the Scrollbar!"
 HScroll1.Min = 8
 HScroll1.Value = 14
 HScroll1.Max = 22
 HScroll1.Visible = True
If clicktwice(1) = 0 Then
 piccodetop = 1440: piccodeheight = 1815
 piccodeleft = 220: piccodewidth = 3375
 pictableleft = 3840: pictablewidth = 2535
 pictableheight = 1095: pictabletop = 1440
 Call resizeall
 
 piccode.Cls
 piccode.Print "Private Sub HScroll1_Change()"
 piccode.Print " lblqcom.Caption = " & q & "Value:" & q
 piccode.Print " lblcom.Caption = HScroll1.Value"
 piccode.Print " lblchaptr.FontSize = HScroll1.Value"
 piccode.Print "End Sub"
 piccode.Print
 piccode.ForeColor = vbBlue
 piccode.Print "'The Scrollbar also requires a  _Scroll() Sub"
 piccode.Print "'for full functionality.  Click in here to"
  Call qscolor("'see that.", blue, 0)
 
 picTable.Visible = True
 picTable.Cls
 picTable.Print "(Name)"; Tab(12); "HScroll1"
 picTable.Print "Max"; Tab(12); "22"
 picTable.Print "Min"; Tab(12); "8"
 picTable.Print "Value"; Tab(12); "14  (default value)"
 lblqcom.Caption = "Value:"
 lblcom.Caption = ""
 piccode.Visible = True
End If

Case 7
If lblalg.Caption <> "back" Then
 lblalg.Caption = "Algebra Reference"
 cmdruntop = 1040: cmdrunleft = 480
 lblchaptr.Caption = "Code": lblchaptr.FontSize = 14
 subj = 0 'case select for cmdrun()
 lblalg.Visible = True
 lblmain.Caption = "Dig.   To the right of equal signs are 'expressions'." & Chr(13) & "click Algebra Reference if you feel adventurous"
 pictableleft = 3840
 pictabletop = 1440
 piccodetop = 1440: piccodeheight = 1815
 piccodewidth = 3375
 pictableheight = 1095
 Call resizeall
 picTable.Cls
 picTable.Print "(Name)"; Tab(12); "cmdrun"
 picTable.Print "Caption"; Tab(12); "&Run"
 lblcom.Caption = Chr(38) & "The &Run Caption niftily assigns alt-r as the hotkey which sets focus to this command button"
 piccode.Cls
 piccode.Print "Private Sub cmdrun_Click()"
 piccode.Print " a = 2: b = 3"
 piccode.Print " str1 = " & q & " yo." & q
 'Call qscolor(" '                       = 'expression'", blue, 0)
 piccode.Print " lblcom.Caption = b ^ (a + 1) & str1"
 piccode.Print " lblqcom.Caption = " & q & "Value:" & q
 piccode.Print "End Sub"
 cmdrun.Visible = True
 lblqcom.Caption = "comment:"
 HScroll1.Visible = False
 piccode.Enabled = False
 lblqprop.Visible = True
 piccode.Visible = True
 picTable.Visible = True

Else: Select Case npage
 Case 1
  Call codecommentary1
 Case 2
  Call codecommentary2
 Case 3
  Call codecommentary3
 Case 4
  Call codecommentary4
 End Select
End If

Case 8
 Call clearform
 lblalg.Caption = "Algebra Reference"
 lblqpage.Visible = False
 piccode.Visible = False
 picTable.Visible = False
 lblalg.Visible = False
 lblchaptr.Caption = "Efficiency"
 lblqcom.Caption = ""
 lblcom.Caption = "Now we're going to look at variables."
 lblmain.Caption = "A Visual Basic editor lets you quickly paste objects all over the place.  When you have them sized and everything, a simple double-click on one will enter the text editor where you type the code."
 lblqprop.Visible = False
 mv1 = 0

Case 9 'variable page 1
 Call clearform
 lblchaptr.Caption = "Modular variables"
 lblmain.Caption = "Modular variables are 'declared' outside all subroutines via the Dim statement, (means 'Dimension this variable') and can be used or modfied by any Subroutine."
 cmd1mv1.Visible = True
 cmd2mv1.Visible = True
 piccode.Visible = True
 piccode.Cls
 piccode.Print "Dim mv1 As Integer"
 piccode.Print
 piccode.Print "Private Sub cmdmodvexample1_Click()"
 piccode.Print " mv1 = mv1 + 1"
 piccode.Print " lblcom.Caption = mv1"
 piccode.Print "End Sub"
 piccode.Print
 piccode.Print "Private Sub cmdmodvexample2_Click()"
 piccode.Print " mv1 = mv1 + 2"
 piccode.Print " lblcom.Caption = mv1"
 piccode.Print "End Sub"

 lblqcom.Caption = ""
 lblcom.Caption = "Howdy"

Case 10 'variable page 2 (locals)
 lblqpage.Visible = False
 cmbansi.Visible = False
 cmbvtyp.Visible = False
 lblqadv.Visible = False

 lblqcom.Caption = ""
 piccode.Visible = True
 lblcom.Caption = 0

 Call clearform

 cmdboxoffice.Visible = False

 piccodeheight = 1815
 piccodetop = 1440
 piccodewidth = 3375
 picTable.Visible = False
 pictableheight = 1095
 pictabletop = 1440
 pictableleft = 3840
 pictablewidth = 2535
 piccodewidth = 3375
 Call resizeall
 
 lblcom.Caption = 0
 cmdrnd.Visible = False
 HScroll1.Visible = False
 txtin.Visible = False
 lblchaptr.Caption = "Local variables"
 lblmain.Caption = "Local variable names from different subroutines can be the same.  A dimensioned local variable name cannot be the same as a modular's."
 picTable.Visible = False
 piccode.Visible = True
 piccode.Cls
 piccode.Print "Private Sub cmd1lv1_Click()"
 piccode.Print " Dim lv1 As Integer"
 piccode.Print " lv1 = lv1 + 1"
 piccode.Print " lblcom.Caption = lv1"
 piccode.Print "End Sub"
 piccode.Print
 piccode.Print "Private Sub cmd2lv1_Click()"
 piccode.Print " Static lv1 As Integer"
 piccode.Print " lv1 = lv1 + 2"
 piccode.Print " lblcom.Caption = lv1"
 piccode.Print "End Sub"

 lblqcom.Caption = "local"

 cmd1lv1.Visible = True
 cmd2lv1.Visible = True

Case 11  'built-in functions
 Call clearform
 cmdcd1.Visible = True
 cmdcd2.Visible = True
 cmdcd3.Visible = True
 cmdcd4.Visible = True
 cmdcd5.Visible = True
 subj = 1
 lblchaptr.Caption = "Built-in functions"
 lstdt.Visible = False
 Select Case npage
 Case 1
  Call codecommentary1
 Case 2
  Call codecommentary2
 Case 3
  Call codecommentary3
 Case 4
  Call codecommentary4
 Case 5
  Call codecommentary5
 
 End Select

 
Case 12
 cmbvtyp.Visible = False 'reference combo boxes off
 cmbansi.Visible = False
 lblqadv.Visible = False

 Call clearform
 cmdboxoffice.Visible = False

 subj = 2
 cmdrnd.Visible = False
 cmdcd1.Visible = True
 cmdcd2.Visible = True
 cmdcd3.Visible = True
 cmdcd4.Visible = True
 cmdcd5.Visible = False

 Call resizeall
 
 lblchaptr.Caption = "Picbox printing"
 Select Case npage
 Case 1
  Call codecommentary1
 Case 2
  Call codecommentary2
 Case 3
  Call codecommentary3
 Case 4
  Call codecommentary4
 End Select
 
  
  
'lblqcom.Caption = "Value:"
'lblcom.Caption = ""
'Call cmdcd1_click  keep this - it works and i'd like to know why here and not elsewhere!!
 'piccode.Visible = False
 'picTable.Visible = False

Case 13
 subj = 3
 Call clearform
 cmdcd1.Visible = True
 cmdcd2.Visible = True
 cmdcd3.Enabled = True
 cmdcd4.Enabled = True
 cmdcd5.Enabled = True
 Call resizeall
 lblchaptr.Caption = "Variable types"
 lblmain.Caption = "I made this as an early reference page."
 lblqcom.Caption = ""
 lblcom.Caption = ""
 piccode.Visible = True
 picTable.Visible = True
 Select Case npage
 Case 1
  Call codecommentary1
 Case 2
  Call codecommentary2
 End Select
 
Case 14
 lblchaptr.Caption = "Sub Procedures"
 subj = 4
 VScroll1.Visible = False
 HScroll1.Visible = False
 cmdcd3.Visible = True
 cmdcd4.Visible = True
 cmdcd5.Visible = True

 Call resizeall
 Select Case npage
 Case 1
  Call codecommentary1
 Case 2
  Call codecommentary2
 Case 3
  Call codecommentary3
 Case 4
  Call codecommentary4
 Case 5
  Call codecommentary5
 End Select

Case 15
 Call resizeall
 lblchaptr.Caption = "Function Procedures"
 subj = 5
 Call enableform
 cmdgobacktop = 840: cmdsimplifytop = 840
 cmdgoback.Visible = False
 cmdsimplify.Visible = False
 cmdrun.Visible = False
 picTable.Visible = False
 cmdrnd.Visible = False
 cmdcd3.Enabled = True: cmdcd4.Enabled = True
 cmdcd5.Enabled = True
 Select Case npage
 Case 1
  Call codecommentary1
 Case 2
  Call codecommentary2
 Case 3
  Call codecommentary3
 Case 4
  Call codecommentary4
 Case 5
  Call codecommentary5
 End Select
 lblqcom.Caption = ""

Case 16
 subj = 6
 cmdcd5.Visible = True
 Call codecommentary1
 txtinwidth = 255
 Call resizeall
 lblchaptr.Caption = "Decisions"
 cmdrun.Visible = False
 cmdA0.Visible = False
 cmdnext.Visible = True
 Select Case npage
 Case 1
  Call codecommentary1
 Case 2
  Call codecommentary2
 Case 3
  Call codecommentary3
 Case 4
  Call codecommentary4
 Case 5
  Call codecommentary5
 End Select

Case 17
 lblchaptr.Caption = "Loops"
 subj = 7

 cmdbc9.Visible = False
 cmdback.Visible = True
 cmdnext.Visible = False

 cmdA0.Visible = True
 cmdcd1.Visible = True
 cmdcd2.Visible = True
 cmdcd3.Visible = True
 cmdcd4.Visible = True
 cmdcd5.Visible = False
 Call resizeall
 Select Case npage
 Case 1
  Call codecommentary1
 Case 2
  Call codecommentary2
 Case 3
  Call codecommentary3
 Case 4
  Call codecommentary4
 End Select
 
End Select
End Sub
Private Sub piccode_Click()
If pIndex = 6 Then
 If clicktwice(1) = 0 Then
  piccodetop = 1150: piccodeheight = 2215
  piccode.Visible = False
  Call resizeall
  qs(1) = qps + "HScroll1_Change()"
  qs(2) = " Call HScommon: End Sub" + Chr(13)
  qs(3) = qps + "HScroll1_Scroll()"
  qs(4) = qs(2)
  qs(5) = qps + "HScommon()"
  qs(6) = " lblqcom.Caption = " & q & "Value:" & q
  qs(7) = " lblcom.Caption = HScroll1.Value"
  qs(8) = " lblchaptr.FontSize = HScroll1.Value"
  qs(9) = "End Sub"
  piccode.Cls
  piccode.Visible = True
  For fn1 = 1 To 9
   piccode.Print qs(fn1)
  Next fn1
  clicktwice(1) = 1 'sets up for click to where we were before
 Else
  clicktwice(1) = 0
  Call contentspre: End If
End If

End Sub
Private Sub lblbuil_Click() 'skip to subject: Built-in functions
pIndex = 11
lblqadv.Visible = True
cancelbyte(1) = 1
npage = 1
subj = 1
Call enableform

lblqskip.Visible = False
lblfunct.Visible = False
lblbuil.Visible = False
lblobj.Visible = False

lblqprop.Visible = False
picTable.Visible = False
piccode.Visible = True
cmdback.Visible = True
Call contentspre
End Sub
Private Sub lblfunct_click()
npage = 1
pIndex = 15

cmbansi.Visible = False
cmbvtyp.Visible = False

Call enableform

cancelbyte(1) = 1
firsttime(1) = 1
lblqpage.Visible = True

lblqskip.Visible = False
lblbuil.Visible = False
lblfunct.Visible = False
lblobj.Visible = False

piccode.Visible = True

cmdcd1.Visible = True
cmdcd2.Visible = True
cmdcd3.Visible = True
cmdcd4.Visible = True
cmdcd5.Visible = True
cmdback.Visible = True
Call contentspre
End Sub
Private Sub lblobj_Click()
aIndex = 16
pIndex = 17
'cmbansi.Visible = False
'cmbvtyp.Visible = False
cmdback.Visible = False: cmdnext.Visible = False
cmdbc9.Visible = True: cmdA0.Visible = True

Call enableform

cancelbyte(1) = 1
firsttime(1) = 1
'lblqpage.Visible = True

lblqskip.Visible = False
lblbuil.Visible = False
lblfunct.Visible = False
lblobj.Visible = False

piccode.Visible = True
'cmdbc6.Visible = True
'cmdcd1.Visible = True
'cmdcd2.Visible = True
'cmdcd3.Visible = True
'cmdcd4.Visible = True
'cmdcd5.Visible = True

Call contents
End Sub
Private Sub pagebeknown()
cmdcd1.FontBold = False
cmdcd2.FontBold = False
cmdcd3.FontBold = False
cmdcd4.FontBold = False
cmdcd5.FontBold = False
Select Case npage
Case 1
cmdcd1.FontBold = True
Case 2
cmdcd2.FontBold = True
Case 3
cmdcd3.FontBold = True
Case 4
cmdcd4.FontBold = True
Case 5
cmdcd5.FontBold = True

End Select
End Sub
Private Sub cmdcd1_click() 'cd series are the page buttons _
 that first appear with 'Built-in functions' builc3page()
npage = 1
Call codecommentary1
End Sub
Private Sub codecommentary1()  'page1 of 'chapter' denoted by _
 the value of subj
 'i now realize that the pIndex value (which changes at cmdnext _
  or cmdback click) could be used
  
Call pagebeknown
Select Case subj
Case 1
cmdgoback.Visible = False
cmdsimplify.Visible = False
cmdrnd.Visible = False
picTable.Visible = False

lblmain.Caption = "Visual Basic has many extremely useful built-in functions like:  FormatCurrency()"
 
'Call clearform
 Dim localn As Byte
 subj = 1
 npage = 1: Call pagebeknown
 cmbansi.Visible = True: lblqadv.Visible = True
 cmbvtyp.Visible = True
 cmbvtypwidth = 1055
 cmbansiwidth = 735
 'cmbansi.Text = "ANSI"

 lblqpage.Visible = True
 lblqcom.Caption = "Total:"
 lblcom.Caption = ""
 picTable.Visible = False
 piccodeheight = 1815
 piccodetop = 1440
 piccodewidth = 3375
 piccodeleft = 120
 picTable.Visible = False
 pictableheight = 1095
 pictabletop = 1440
 pictableleft = 3840
 pictablewidth = 2535
 piccode.Visible = False
 cmdboxofficeleft = 570
 Call resizeall
 piccode.Cls
 piccode.Visible = True
 Call qscolor(qps & "cmdboxoffice_click()", 0, 0)
 piccode.Print " Dim pop As Integer, soda As Integer"
 piccode.Print " pop = 50"
 piccode.Print " soda = 80"
 piccode.Print " lblcom.Caption = FormatCurrency(pop + soda)"
 piccode.Print "End Sub"
 cmdboxoffice.Visible = True
 txtin.Visible = False
 HScroll1.Visible = False

Case 2
 piccode.Visible = False: picTable.Visible = False
 piccode.Cls: picTable.Cls
 piccodetop = 950: pictabletop = 1730
 piccodeleft = 200: piccodewidth = 3300
 pictableleft = 3600: pictablewidth = 2800
 pictableheight = 1660: piccodeheight = 2450
 Call resizeall
 piccode.Visible = True: picTable.Visible = True
 Dim qhello As String
 cr = Chr(13): ltab = Int(HScroll1.Value / 5)
 qhello = "hello"
 lblmain.Caption = "These white rectangles are picture boxes.  They are graphics-capable (we'll get into that)"
 piccode.Print "Private Sub PrintToPicExample()"
 piccode.Print " Dim qhello As String"
 piccode.Print " qhello = " & q & "hello" & q
 piccode.Print " picTable.Cls  ";: Call qscolor("'This clears picTable", blue, 0)
 piccode.Print " picTable.Print qhello & 1; qhello"
 piccode.Print " picTable.Print " & q & "veddy inteddesting." & q & ", qhello"
 piccode.Print " picTable.Print Tab(19); qhello"
 piccode.Print " picTable.Print Tab(11); qhello, qhello;"
 piccode.Print " picTable.Print qhello"
 piccode.Print " picTable.Print 123"
 piccode.Print " picTable.Print -123"
 piccode.Print "End Sub"
 picTable.Print qhello & 1; qhello; ""
 picTable.Print "veddy inteddesting.", qhello
 picTable.Print Tab(19); qhello
 picTable.Print Tab(11); qhello, qhello;
 picTable.Print qhello
 picTable.Print 123
 picTable.Print -123
 lblcom.Caption = " , ; + and ampersand tell picTable to combine printed variable data and regular text a certain way."
 lblqcom.Caption = ""
 HScroll1.Visible = False

Case 3
 HScroll1.Visible = False
 piccode.Visible = False: picTable.Visible = False
 piccodewidth = 2990
 pictablewidth = 3300
 piccodeleft = 130
 pictableleft = 3220
 piccodeheight = 1450
 pictableheight = 1450
 piccodetop = 1150: pictabletop = 1150
 Call resizeall
 piccode.Cls
 picTable.Cls
 piccode.Visible = True: picTable.Visible = True
 piccode.Print "Type"; Tab(22); "Size"
 piccode.Print
 picTable.Print "Content"
 picTable.Print
 piccode.Print "String"; Tab(22); "Length of string"
 picTable.Print "1 to approx 65,400"
 piccode.Print "String var-length"; Tab(22); "10 bytes + str len"
 picTable.Print "0 to approx 2 billion"
 piccode.Print "Variant (nums)"; Tab(22); "16 bytes"
 picTable.Print "Any numeric value up to a Double's range"
 piccode.Print "Variant (char)"; Tab(22); "22 bytes + str len"
 picTable.Print "0 to approx 2 billion"
 piccode.Print "User-defined"; Tab(22); "# req. by elements"
 picTable.Print "The range of each elem = data type range"

Case 4
If firsttime(1) = 0 Then  '0 here means 'true'
 cmdcd3.Enabled = False
 cmdcd4.Enabled = False
 cmdcd5.Enabled = False
End If
 piccode.Visible = False: picTable.Visible = False
 lblmain.Caption = "You are about to learn a most fundamental programming concept.  This is one to write home about.  It is a method that allows a team of programmers to make one huge program."
 cmdrun.Visible = False
 cmdsimplify.Visible = False
 cmdgoback.Visible = False
 lblcom.Caption = ""

Case 5
 piccodeheight = 2200: piccodewidth = 4300
 piccodeleft = 500: piccodetop = 1270: piccode.Cls
 Call resizeall
 piccode.Visible = True
 HScroll1.Visible = False: VScroll1.Visible = False
 lblcom.Caption = ""
 Call printmore
 cmdsimplify.Visible = False

Case 6
 txtin.Visible = False
 cmdsimplify.Visible = False
 cmdgoback.Visible = False
 lblmain.Caption = "If I want to go to Dallas, I will take the next exit.  If gas is running low, I will head for McDonald's." & Chr(13) & q & "No .. what.  Really?" & q & Chr(13) & q & "Oh.  Yeah I meant a food station." & q
 piccode.Visible = False: picTable.Visible = False
 HScroll1.Visible = False: VScroll1.Visible = False
 lblcom.Caption = ""

Case 7
 HScroll1.Visible = False
 cmdsimplify.Visible = False: cmdruntop = 960
 cmdgoback.Visible = False
 VScroll1.Visible = False: txtin.Visible = False
 piccodetop = 1350: pictabletop = 1350
 piccodeleft = 350: piccodewidth = 3850
 piccodeheight = 2400: pictableheight = 2400
 pictableleft = 4360: pictablewidth = 1950
 Call resizeall
 lblmain.Caption = q & "Loops can be fun!  " & q & "  This is a For-Next loop."
 cmdrun.Visible = True
 piccode.Visible = True: picTable.Visible = True
 piccode.Cls: picTable.Cls
 piccode.Print "Private sub cmdrun_click()"
 piccode.Print " Dim lcount As Byte, lrnd As Single"
 piccode.Print " Dim localsum As Single"
 piccode.Print " picTable.Cls"
 piccode.ForeColor = vbBlue
 piccode.Print " For lcount = ";: piccode.ForeColor = vbMagenta: piccode.Print "1 to 3": piccode.ForeColor = vbBlack
 piccode.Print "  lrnd = Rnd * 12"
 piccode.Print "  localsum = localsum + lrnd"
 piccode.Print "  picTable.Print lcount; lrnd": piccode.ForeColor = vbBlue
 piccode.Print " Next lcount": piccode.ForeColor = vbBlack
 piccode.Print " picTable.Print " & q & "Final lcount" & q & ", lcount"
 piccode.Print "picTable.Print " & q & "Rnd Average" & q & ", localsum / (lcount - 1)"
 piccode.Print "End Sub"
 lblqcom.Caption = "": lblcom.Caption = ""

Case 8
piccode.Cls
lblqcom.Caption = "": lblcom.Caption = ""
lblmain.Caption = "Expressions"
piccode.Print "Consider:": piccode.Print
piccode.ForeColor = vbBlue
piccode.Print "To the right of the equal sign is an Expression."
piccode.Print "Expressions are calculated First."
piccode.Print "If x = 56, and we get a new line of code:"
piccode.Print "x = x + 1,  57 would be x's new value."
piccode.Print: piccode.Print
Call qscolor(" click Page 2", mag, 0)

End Select
End Sub
Private Sub cmdcd2_click()
Call codecommentary2
End Sub
Private Sub codecommentary2()  'page2 of 'chapter' denoted by _
 the value of subj
txtin.Visible = False
HScroll1.Visible = False
npage = 2
Call pagebeknown
Select Case subj
Case 1
cmdgoback.Visible = False
cmdsimplify.Visible = False
cmdboxoffice.Visible = False
cmdrnd.Visible = False
piccodetop = 1150
pictabletop = 1150
piccodeheight = 2215
pictableheight = 2215
pictableleft = 2220
pictablewidth = 1200
piccodewidth = 2055
piccodeleft = 120
Call resizeall

picTable.Visible = True
lblmain.Caption = "FormatCurrency() worked with numeric data, but the output was in string format because of the Dollar sign.  These functions output numbers."
piccode.Cls
picTable.Cls
piccode.Print "  Numeric function"
picTable.Print "   Value"
piccode.Print "Sqr(2)"
picTable.Print 1.414214
piccode.Print "Round(2.7)"
picTable.Print 3
piccode.Print "Round(2.317, 2)"
picTable.Print 2.32
piccode.Print "Int(2.7)"
picTable.Print 2
piccode.Print "Int(-2.7)"
picTable.Print -3
piccode.Print "Val(" & q & "$53.04" & q & ")"
picTable.Print Val("$53.04")
piccode.Print "Val(" & q & "53q4" & q & ")"
picTable.Print Val("53q4")

str1 = "hello there, chalupa"
piccode.Print "str1 =" & q & "hello there, chalupa" & q
picTable.Print
piccode.Print "Len(str1)"
picTable.Print Len(str1)
piccode.Print "InStr(str1, " & q & "e," & q & ")"
picTable.Print InStr(str1, "e,")

lblqcom.Caption = "comment:"
lblcom.Caption = "The Val function is common to almost every programming language."

Case 2
piccode.Cls
piccodetop = 1200: pictabletop = 1200
piccodeleft = 240: piccodewidth = 1780
pictableleft = 2110: pictablewidth = 4200
pictableheight = 2050: piccodeheight = 2050
Call resizeall
piccode.Visible = True: picTable.Visible = True
piccode.Visible = True: picTable.Visible = True
qhello = "hello"
HScroll1.Min = 1: HScroll1.Max = 55
cr = Chr(13)
lblmain.Caption = "Put Slider completely left to more easily visualize match-up of code shown in the left box."
piccode.Print "qhello & 1; qhello"
'piccode.Print "qhello & 1, qhello"
'picTable.Print qhello & 1, qhello
piccode.Print "qhello & 1; 1 + 1, qhello"
piccode.Print "-1; 2"
piccode.Print "1,"
piccode.Print "12; Tab(ltab); qhello, 1"
piccode.Print "qhello + qhello"
piccode.Print "qhello + 1"
HScroll1.Visible = True
Call HScommon

Case 3
piccodewidth = 1690
pictablewidth = 4400
pictableleft = 1920
piccodeheight = 2650
pictableheight = 2650
piccodetop = 950: pictabletop = 950
Call resizeall
piccode.Cls
picTable.Cls
';tab(29);
piccode.Print "Byte"; Tab(14); "1 byte"
picTable.Print "0 to 255"
piccode.Print "Boolean"; Tab(14); "2 bytes"
picTable.Print "True or False"
piccode.Print "Integer"; Tab(14); "2 bytes"
picTable.Print "-32,768 to 32,767"
piccode.Print "Long"; Tab(14); "4 bytes"
picTable.Print "-2,147,483,648 to 2,147,483,647"
piccode.Print "Single "; Tab(14); "4 bytes"
picTable.Print "-3.402823E38 to -1.401298E-45 for - vals"
piccode.Print
picTable.Print " 1.401298E-45 to 3.402823E38 for +"
piccode.Print "Double"; Tab(14); "8 bytes"
picTable.Print "-1.79769313486232E308 to -49065645841247E-324"
piccode.Print
picTable.Print " 4.94065645841247E-324 to 1.79769313486232E308"
piccode.Print "Currency"; Tab(14); "8 bytes"
picTable.Print "-922,337,203,685,477.5808 to 922,337,203,685,477.5807"
piccode.Print "Decimal"; Tab(14); "14 bytes"
picTable.Print " +/- 79,228,162,514,264,337,593,543,950,335 (no dec point)"
piccode.Print
picTable.Print " +/-7.9228162514264337593543950335 (that's 28 after dec)"
piccode.Print
picTable.Print "smallest non-zero is +/-0.0000000000000000000000000001"
piccode.Print "Date"; Tab(14); "8 bytes"
picTable.Print "Jan 1, 100 to Dec 31, 9999"

Case 4
lblmain.Caption = "We are 'passing' variables, starting with the Call statement, to another Subroutine.  It is like a relay race."
piccodetop = 1140: piccodewidth = 4900
piccodeheight = 2200: piccodeleft = 1200: piccode.Cls
cmdgobacktop = 1190: cmdgobackleft = 280
cmdsimplifytop = 1190: cmdsimplifyleft = 280
'Call contentspre
Call resizeall
piccode.Visible = True
If firsttime(1) = 0 Then
End If
Call printmore
'cmdgoback.Caption = "More.."
cmdgoback.Visible = True: cmdgoback.SetFocus

Case 5
lblmain.Caption = "In this example, HSCommon says, 'Area, here is some data.  Give me your expert conclusion.'"
piccode.Cls
piccode.Print "Private Sub HScommon()"
piccode.Print " Dim ";: Call qscolor("x As Integer", mag, 1): piccode.Print ",";: Call qscolor(" y As Integer", red, 0)
Call qscolor(" x ", mag, 1): piccode.Print "= HScroll1.Value: ";: Call qscolor("y ", red, 1): piccode.Print "= VScroll1.Value"
piccode.Print " ..Area(";: Call qscolor("x", mag, 1): piccode.Print ",";: Call qscolor(" y", red, 1): piccode.Print "): End Sub"
'piccode.Print "Private Sub VScommon()"
'piccode.Print " Dim x As Integer, y As Integer"
piccode.Print: piccode.Print
piccode.Print " .."
'piccode.Print " x = HScroll1.Value: y = HScroll1.Value"
'piccode.Print " ..Area(x, y): End Sub"
piccode.Print
piccode.Print "Private Function Area(.......................................) As Integer"
piccode.Print " ................. "
piccode.Print "End Function"
HScroll1.Visible = False: VScroll1.Visible = False
lblcom.Caption = ""
Case 6
cmdsimplify.Visible = False
cmdgoback.Visible = False
HScroll1.Visible = False: VScroll1.Visible = False
lblcom.Caption = ""
lblmain.Caption = "Decisions .. If 'expression' Then 'statement(s)'.  In vbBlindmeGreen is a Logical Operator."
piccode.Visible = True: picTable.Visible = False
piccodetop = 950: pictabletop = 1150
piccodeleft = 500: piccodewidth = 4150
pictableleft = 3050: pictableheight = 1900
piccodeheight = 2200: pictablewidth = 2700
Call resizeall
piccode.Cls
'piccode.Print "If x = 3 Or x = 6 Then"
'piccode.Print "If x <= 1 And x > 0 Then"
'piccode.Print "If x <> 3 * y .. (means if x is not equal to 3 times y)"
'piccode.Print "If lastname > " & q & "M" & q & ".. recognizes that 'A' comes before 'B'"
'piccode.Print
piccode.Print "Private Sub HScommon()"
Call qscolor(" If", blue, 1): piccode.Print " HScroll1.Value < 50 ";: Call qscolor("Then", blue, 0)
piccode.Print "  lblcom.Caption = " & q & "   " & q & " & HScroll1.Value"
Call qscolor(" Else", blue, 1)
piccode.Print "  lblcom.Caption = HScroll1.Value"
Call qscolor(" End If", blue, 0)
piccode.Print " lblqcom.Caption = " & q & q
Call qscolor(" If", mag, 1): piccode.Print " HScroll1.Value >= 20 ";: Call qscolor(" And ", grn, 1): piccode.Print "HScroll1.Value <= 30 ";: Call qscolor("Then", mag, 0)
piccode.Print "  lblqcom.Caption = " & q & "Rain" & q
Call qscolor(" End If", mag, 0)
piccode.Print "End Sub"
lblqcom.Caption = "": lblqcom.Visible = True '?
HScroll1.Min = 1: VScroll1.Min = 1
HScroll1.Value = 1: VScroll1.Value = 1
HScroll1.Max = 80: VScroll1.Max = 50
HScroll1.Visible = True

Case 7
lblmain.Caption = "Mixing things up .. This subscripted variable is called an Array."
txtintop = 3500: txtin.Visible = True: txtin.SetFocus
txtinwidth = 255: txtinleft = 900: lblcom.Caption = ""

cmdrun.Visible = False: cmdsimplify.Visible = False
piccodeleft = 200: VScroll1.Visible = False
lblcom.Visible = True
piccodewidth = 6300: picTable.Visible = False
piccodeheight = 2450: piccodetop = 1000: piccode.Cls
Call resizeall
piccode.Print "Private Sub txtin_Keypress(K As Integer)";: Call qscolor(" 'K gets ANSI value.  If '1' is pressed, K = 49.", blue, 0)
piccode.Print " Dim teamName";: Call qscolor("(1 To 5)", mag, 1): piccode.Print " As String, tn As Integer"
piccode.Print " teamName";: Call qscolor("(1)", mag, 1): piccode.Print " = " & q & "Red Sox" & q
piccode.Print " teamName";: Call qscolor("(2)", mag, 1): piccode.Print " = " & q & "Giants" & q
piccode.Print " teamName";: Call qscolor("(3)", mag, 1): piccode.Print " = " & q & "White Sox" & q
piccode.Print " teamName";: Call qscolor("(4)", mag, 1): piccode.Print " = " & q & "Cubs" & q & ": teamName";: Call qscolor("(5)", mag, 1): piccode.Print " = " & q & "Cubs" & q
piccode.Print " tn = K - 48  ";: Call qscolor("'i.e. '1' is pressed, K = 49, tn = 1.", blue, 0)
piccode.Print " If tn < 1 Or tn > 5 Then"
piccode.Print "  lblcom.Caption = " & q & q & ": Else"
piccode.Print "  lblcom.Caption = " & q & "The " & q & " & teamName";: Call qscolor("(tn)", mag, 1): piccode.Print " & " & q & " won World Series number " & q & " & tn: End If"
piccode.Print " txtin.SetFocus"
piccode.Print "End Sub"
txtin.Text = ""
Case 8
lblmain.Caption = "Order of Operations is the same for all algebraic problem-solving."
piccode.Cls
Call qscolor("The equation", blue, 0)
piccode.Print "x = 2 ^ (4 + 3 * 2) / 5";: Call qscolor(" is solved by:", blue, 0)
piccode.Print "1. Outermost Parentheses";: Call qscolor(" 4 + 3 * 2", blue, 0)
piccode.Print
'piccode.Print "2. Power, or Exponent ";: Call qscolor("4 + 3 * 2 ", blue, 0)
piccode.Print "3. Multiply, Divide ";: Call qscolor("4 + ", blue, 1): Call qscolor("6", mag, 0)
piccode.Print "4. Add, Subtract ";: Call qscolor("10", blue, 0)
piccode.Print: piccode.Print
Call qscolor("psst, i'm a loser", mag, 0)
End Select
End Sub
Private Sub cmdcd3_click()
Call codecommentary3
End Sub
Private Sub codecommentary3()  'page3 of 'chapter' denoted by _
 the value of subj
npage = 3
Call pagebeknown
Select Case subj
Case 1
txtin.Visible = False
HScroll1.Visible = False
cmdgoback.Visible = False
cmdsimplify.Visible = False
cmdboxoffice.Visible = False
cmdrnd.Visible = False
lblmain.Caption = "These are String functions like FormatCurrency().."
lblcom.Caption = ""
picTable.Visible = True
pictableleft = 3470
pictablewidth = 2600
pictabletop = 1000
piccodeleft = 400
piccodetop = 1000
piccodewidth = 3015
piccodeheight = 2430
pictableheight = 2430
Call resizeall

piccode.Cls
picTable.Cls
'piccode.Print "Trim(" & q & " hello " & q & ")"
'picTable.Print q & Trim(" hello ") & q
piccode.Print "UCase(" & q & "hello" & q & ")"
picTable.Print q & UCase("hello") & q
piccode.Print "LCase(" & q & "heLLO" & q & ")"
picTable.Print q & LCase("HELLO") & q
piccode.Print "Left(" & q & "hello" & q & ", 3)"
picTable.Print q & Left("hello", 3) & q
piccode.Print "Right(" & q & "hello" & q & ", 4)"
picTable.Print q & Right("hello", 4) & q
piccode.Print "Mid(" & q & "hola there" & q & ", 4, 3)"
picTable.Print q & Mid("hola there", 4, 3) & q
piccode.Print "FormatNumber(1000 + Sqr(2), 3)"
picTable.Print q & FormatNumber(1000 + Sqr(2), 3) & q
piccode.Print "FormatPercent( .185, 2)"
picTable.Print q & FormatPercent(0.185, 2) & q
piccode.Print "FormatDateTime(9-15-99)"
picTable.Print q & FormatDateTime(9 - 15 - 99) & q
piccode.Print "FormatDateTime(" & q & "9-15-99" & q & ")"
picTable.Print q & FormatDateTime("9-15-99") & q
piccode.Print "FormatDateTime(9-15-99, vbLongDate)"
picTable.Print q & FormatDateTime(9 - 15 - 99, vbLongDate) & q
piccode.Print "FormatDateTime(" & q & "9-15-99" & q & ", vbLongDate)"
picTable.Print q & FormatDateTime("9-15-99", vbLongDate) & q
piccode.Print "Chr(64)"
picTable.Print q & Chr(64) & q
lblqcom.Caption = "comment:"
lblcom.Caption = "The Chr() function will print 1 of 255 characters from the ANSI Character set.  'A' has a value of 65."
lblchaptr.Caption = "Built-in functions"

Case 2
lblmain.Caption = "Simple color text output to picturebox"
HScroll1.Visible = False: lblqcom.Caption = ""
piccode.Cls: picTable.Cls
piccodetop = 1150: pictabletop = 1150
piccodeleft = 140: piccodewidth = 2300
pictableleft = 2500: pictablewidth = 3400
pictableheight = 2250: piccodeheight = 2250
Call resizeall
Dim qhello As String
qhello = "hello"
lblcom.Caption = "vbRed, vbYellow, vbGreen, vbBlue, vbBlack, vbWhite, vbMagenta, vbCyan"
piccode.Print ".."
piccode.Print "picTable.ForeColor = vbRed": picTable.ForeColor = vbRed
piccode.Print "picTable.Print qhello": picTable.Print qhello
piccode.Print "picTable.ForeColor = vbBlack": picTable.ForeColor = vbBlack
piccode.Print "picTable.Print qhello": picTable.Print qhello

Case 4
piccode.Visible = False
picTable.Visible = False
lblmain.Caption = "On page 2, we see a 'pass By Reference', where ndimes and dimez share a memory location.  Think of the memory location as a person, and the names as alias."
lblcom.Caption = ""
cmdrun.Visible = False
cmdsimplify.Visible = False
cmdgoback.Visible = False

Case 5
lblmain.Caption = "Function says 'okay, I got your data, your types match my types so now I'm computing.."
piccode.Cls
lblcom.Caption = "By default, x and y pass By Reference to width and height respectively.  width is an alias for x, height is an alias for y."
HScroll1.Visible = False: VScroll1.Visible = False
piccode.Print "Private Sub HScommon()"
piccode.Print " Dim x As Integer, y As Integer"
piccode.Print " x = HScroll1.Value: y = VScroll1.Value"
piccode.Print " lblcom.Caption = .. Area(";: Call qscolor("x", mag, 1): piccode.Print ", ";: Call qscolor("y", red, 1): piccode.Print "): End Sub"
'piccode.Print "Private Sub VScommon()"
'piccode.Print " Dim x As Single, y As Single"
piccode.Print: piccode.Print
piccode.Print " .."
'piccode.Print " x = HScroll1.Value: y = HScroll1.Value"
'piccode.Print " ..Area(x, y): End Sub"
piccode.Print
piccode.Print "Private Function Area(";: Call qscolor("width", mag, 1): piccode.Print " As Integer, ";: Call qscolor("height", red, 1): piccode.Print " As Integer) .."
piccode.Print " Area = ";: Call qscolor("width", mag, 1): piccode.Print " * ";: Call qscolor("height", red, 0)
piccode.Print "End Function"

Case 6
cmdsimplify.Visible = False
cmdgoback.Visible = False
lblmain.Caption = "The If statement, continued." & Chr(13) & Chr(13) & q & "Eww, logic" & q
HScroll1.Visible = False: VScroll1.Visible = False
lblcom.Caption = ""
piccodetop = 1200: pictabletop = 1200
piccodeleft = 50: piccodewidth = 2900
pictableleft = 3000: pictablewidth = 3700
piccodeheight = 2400: pictableheight = 1600
Call resizeall

piccode.Cls: picTable.Cls
piccode.Visible = True: picTable.Visible = True
piccode.Print "Logical Operators"
piccode.Print
piccode.Print "Xor", "1 and only 1 must be true"
piccode.Print "Or", "at least one must be true"
piccode.Print "And", "both must be true"
piccode.Print "Not", "opposite"
piccode.Print
piccode.Print "Relational Operators"
piccode.Print
piccode.Print "=    <>", "equal to , not equal to"
piccode.Print "<=   <", "lessthan or equal , less"
piccode.Print ">=   >", "greater or equal , greater"
picTable.Print "Example": picTable.Print
picTable.Print "road A Xor road B"
picTable.Print "I have $5.  Should I buy this $1 Or this $2 item?"
picTable.Print "Find me [first name] [last name]'s phone number."
picTable.Print "If Not WalmartPrice1 = TargetPrice1 Then"
picTable.Print " lstrfam = " & q & "Honey, let's buy a hovercraft instead." & q
lblqcom.Caption = ""
Case 7
txtin.Text = ""
lblmain.Caption = "The Do-While Loop .. Type a value, click Run, .. we're assuming a interest rate of 7%"
txtintop = 3360: txtinwidth = 600: txtin.Visible = True
cmdruntop = 570: cmdrun.Visible = True: picTable.Cls
piccodetop = 950: pictabletop = 950
piccodeleft = 300: piccodewidth = 3600
pictableleft = 4200: pictablewidth = 2000
piccodeheight = 2300: pictableheight = 1000
Call resizeall

piccode.Cls: picTable.Visible = True
piccode.Print "Private sub cmdrun_Click()"
piccode.Print " Dim balance As Single, numYears As Integer"
piccode.Print " balance = Val(txtin.Text)"
piccode.Print " numYears = 0"
piccode.Print " Do While balance < 1000000"
piccode.Print "  balance = balance + .07 * balance"
piccode.Print "  numYears = numYears + 1"
piccode.Print " Loop"
piccode.Print " picTable.Cls"
piccode.Print " picTable.Print " & q & "About " & q & "; " & "numYears; " & q & "years." & q
piccode.Print "End Sub"
lblcom.Caption = "Enter a one-time deposit between $50 and $50,000."
txtin.SetFocus

Case 8
piccode.Cls: piccode.Print
piccode.Print "x = 2 ^ (     10    ) / 5"
piccode.Print
piccode.Print "2. Power, or Exponent ";: Call qscolor("2^10", blue, 1): piccode.Print " / 5"
piccode.Print "3. Multiply, Divide ";: Call qscolor("1024 / 5", blue, 0)
piccode.Print: piccode.Print: piccode.Print
Call qscolor("i want to make a game ", mag, 0)
End Select
End Sub
Private Sub cmdcd4_click()
Call codecommentary4
End Sub
Private Sub codecommentary4()  'page4 of 'chapter' denoted by _
 the value of subj
txtin.Visible = False
HScroll1.Visible = False
cmdgoback.Visible = False
cmdsimplify.Visible = False
npage = 4
Call pagebeknown
Select Case subj
Case 1
cmdrndtop = 900
cmdrndleft = 800
cmdboxoffice.Visible = False
picTable.Visible = True
lblqcom.Caption = "comment:"
lblcom.Caption = "Page 5 get ready .. try to remain 'underwhelmed' .. :)  Pay attention to my label comments."
piccodetop = 1290
pictabletop = 1290
pictableleft = 3450
pictablewidth = 2600
piccodewidth = 3105
piccodeheight = 2015
pictableheight = 2015
piccodeleft = 320
Call resizeall

lblmain.Caption = "The Rnd function generates random numbers.."
piccode.Cls
picTable.Cls
piccode.Print "Private Sub cmdrnd_Click()"
piccode.Print " picTable.Cls  'remark: this clears picTable"
piccode.Print " picTable.Print"
picTable.Print
piccode.Print " picTable.Print Int(10 * Rnd);"
piccode.Print " picTable.Print 10 * Rnd"
picTable.Print Int(10 * Rnd);
picTable.Print 10 * Rnd
piccode.Print "End Sub"
cmdrnd.Visible = True

Case 2
lblmain.Caption = "At 1440 units per inch, (depends on screen) picTabletop = 0 would hit just below the 'Learning Visual Basic' bar."
HScroll1.Visible = False: lblqcom.Caption = ""
piccode.Cls: picTable.Cls
piccodetop = 1150: pictabletop = 1150
piccodeleft = 800: piccodewidth = 2300
pictableleft = 3180: pictablewidth = 2500
pictableheight = 2250: piccodeheight = 2250
Call resizeall
Dim qhello As String
qhello = "hello"
lblcom.Caption = "vbRed, vbYellow, vbGreen, vbBlue, vbBlack, vbWhite, vbMagenta, vbCyan"
piccode.Print "piccodetop = 1150"
piccode.Print "piccodeheight = 2250"
piccode.Print "piccodeleft = 800"
piccode.Print "piccodewidth = 2300"
picTable.Print "picTabletop = 1150"
picTable.Print "picTableheight = 2250"
picTable.Print "picTableleft = 3160"
picTable.Print "picTablewidth = 2500"
lblcom.Caption = "And of course there is piccode.Visible = False .."


Case 4
cmdrun.Visible = False
piccode.Visible = True
piccodetop = 1290: piccodewidth = 5050: piccodeheight = 2050: piccodeleft = 1240: piccode.Cls
cmdgobacktop = 1370: cmdgobackleft = 310
cmdsimplifytop = 1370: cmdsimplifyleft = 310
Call resizeall
lblmain.Caption = "Passing By Value, however, tells the called procedure to leave the original value intact by creating a new memory location."
lblcom.Caption = ""
Call printmore
cmdgoback.Visible = True

Case 5
lblcom.Caption = ""
lblmain.Caption = q & "Okay buddy here's your Integer" & q
piccode.Cls: HScroll1.Visible = False: VScroll1.Visible = False
piccode.Print "Private Sub HScommon()"
piccode.Print " .."
piccode.Print " .."
piccode.Print " lblcom.Caption = ............................ ";: Call qscolor("Area(", blue, 1): piccode.Print "x, y";: Call qscolor(")", blue, 1): piccode.Print " End Sub"
'piccode.Print " lblcom.Caption = x & " & q & "x" & q & " & y & " & q & "=" & q & " & Area(x, y): End Sub"
'piccode.Print "Private Sub VScommon()"
'piccode.Print " Dim x As Single, y As Single"
piccode.Print: piccode.Print
piccode.Print " .."
'piccode.Print " x = HScroll1.Value: y = VScroll1.Value"
'piccode.Print " ..Area(x, y): End Sub"
piccode.Print
piccode.Print "Private Function ";: Call qscolor("Area(", blue, 1): piccode.Print ".......................................";: Call qscolor(") As Integer", blue, 0)
Call qscolor(" Area ", blue, 1): piccode.Print "= width * height"
piccode.Print "End Function"
VScroll1.Visible = False

Case 6
lblmain.Caption = "Highlited in vbMagenta are expressions."
cmdsimplify.Visible = False
cmdgoback.Visible = False

pictabletop = 960: piccodetop = 960
piccodeheight = 2400: pictableheight = 2400
piccodewidth = 2490: pictableleft = 2690
pictablewidth = 3700: piccodeleft = 120

Call resizeall

piccode.Visible = True: picTable.Visible = True
HScroll1.Visible = True: VScroll1.Visible = False
piccode.Cls: picTable.Cls
piccode.Print "Dim clscount As Integer"
piccode.Print "Private Sub HScommon()"
piccode.Print " Dim t As Single"
piccode.Print " If ";: piccode.ForeColor = vbMagenta: piccode.Print "clscount";: piccode.ForeColor = vbBlue: piccode.Print " > ";: piccode.ForeColor = vbMagenta: piccode.Print "14 ";: piccode.ForeColor = vbBlack: piccode.Print "Then"
piccode.Print " picTable.Cls: clscount = 0: End If"
piccode.Print " t = HScroll1.Value"
piccode.Print " If";: Call qscolor(" t / 5 ", mag, 1): Call qscolor("=", blue, 1): Call qscolor(" Int(t / 5) ", mag, 1): Call qscolor("Or ", grn, 1): piccode.Print "_": piccode.ForeColor = vbMagenta
piccode.Print "  t / 4 ";: Call qscolor("=", blue, 1): Call qscolor(" Int(t / 4)", mag, 1): piccode.Print " Then"
piccode.Print " clscount = clscount + 1"
piccode.Print " picTable.Print Tab(t); t;"
piccode.Print " End If"
piccode.Print "End Sub"
lblqcom.Caption = ""
lblcom.Caption = ""
HScroll1.Min = 1: HScroll1.Max = 40

Case 7
cmdrun.Visible = False: VScroll1.Visible = False
HScroll1.Visible = False: piccode.Visible = True: picTable.Visible = False
piccodeheight = 1250: piccodewidth = 2200: piccodeleft = 2000
Call resizeall
piccode.Cls: lblmain.Caption = "And one last For-Next Loops tidbit:"
piccode.Print "For n = 1 To 40 Step .1"
piccode.Print " ..: Next n"
piccode.Print "For n = 1 To 40 Step x / 10"
piccode.Print " ..: Next n"
piccode.Print "For n = 8 to 2 Step -.03"
piccode.Print " ..: Next n"
lblcom.Caption = "Loops section q's/comments to fluoats@hotmail.com"
cmdA0.Enabled = True

Case 8
piccode.Cls: piccode.Print
piccode.Print "x = 1024 / 5"
piccode.Print
'piccode.Print "2. Power, or Exponent ";: Call qscolor(" 2^2", blue, 0)
piccode.Print
piccode.Print "x = " & 1024 / 5
piccode.Print: piccode.Print: piccode.Print
Call qscolor("i'm making this tute cuz i learn in the process ", mag, 0)

End Select
End Sub
Private Sub cmdcd5_click()

Call codecommentary5
End Sub
Private Sub codecommentary5()  'page5 of 'chapter' denoted by _
 the value of subj
cmdboxoffice.Visible = False
cmdrnd.Visible = False
npage = 5
Call pagebeknown
Select Case subj
Case 1

cmdgobackleft = 700
cmdgobacktop = 900
cmdsimplifyleft = 700
cmdsimplifytop = 900
cmdgoback.Caption = "More .."
cmdgoback.Visible = True
cmdsimplify.Caption = "Unmore"
lblqcom.Caption = "Value:"
lblcom.Caption = ""
piccodetop = 1300
piccodewidth = 6350
piccodeheight = 2015
piccodeleft = 120
picTable.Visible = False
HScroll1.Min = 32
HScroll1.Value = 32
HScroll1.Max = 255
HScroll1.Visible = True
txtintop = 3390
Call resizeall
txtin.Visible = True
Call printmore   'sibling to cmdgoback().  Switches _
 between simple/in-depth code samples.
txtin.SetFocus

Case 4
lblcom.Caption = ""
lblmain.Caption = "Sub Procedure, servant to Event Procedures.  This architecture is referred to as Top-down programming."
cmdrun.Visible = False
cmdgoback.Visible = False
cmdsimplify.Visible = False
piccode.Cls
piccodewidth = 4520
piccodetop = 1150
piccodeheight = 2200
piccodeleft = 500
Call resizeall
piccode.Visible = True
piccode.Print "Private sub cmdnext_click()"
piccode.Print "Call message3"
'piccode.Print ".."
piccode.Print "End Sub"
piccode.Print
piccode.Print "Private sub cmdback_Click()" 'cmdbm3 only visible on 'message4'"
piccode.Print "Call message3"
'piccode.Print ".."
piccode.Print "End Sub"
piccode.Print
piccode.Print "Private sub message3()"
piccode.Print "lblmain.Caption = " & q & "The commandbutton with 'next' as its caption has a 'click event' that activates code!  Next I'm going to show you what that looks like." & q
'piccode.Print ".."
piccode.Print "End Sub"

Case 5
 lblmain.Caption = "And I couldn't resist throwing in another scrollbar."
 piccode.Cls
 piccode.Print "Private Sub HScommon()"
 piccode.Print " Dim x As Integer, y As Integer"
 piccode.Print " x = HScroll1.Value: y = VScroll1.Value"
 piccode.Print " lblcom.Caption = x & " & q & "*" & q & " & y & " & q & "=" & q & " & ";
 Call qscolor("Area(", blue, 1): Call qscolor("x, y", black, 1): Call qscolor(")", blue, 1): piccode.Print ": End Sub"
 piccode.Print "Private Sub VScommon()"
 piccode.Print " Dim x As Integer, y As Integer"
 piccode.Print " x = HScroll1.Value: y = VScroll1.Value"
 piccode.Print " lblcom.Caption = x & " & q & "*" & q & " & y & " & q & "=" & q & " &";: Call qscolor(" Area(", blue, 1): Call qscolor("x, y", black, 1): Call qscolor(")", blue, 1): piccode.Print " End Sub"
 Call qscolor("Private Function ", black, 1): Call qscolor("Area(", blue, 1): piccode.Print "width As Integer, height As Integer";: Call qscolor(") As Integer", blue, 0)
 Call qscolor(" Area", blue, 1): piccode.Print " = width * height"
 piccode.Print "End Function"
 HScroll1.Min = 20: HScroll1.Max = 50
 VScroll1.Min = 20: VScroll1.Max = 30
 HScroll1.Visible = True
 VScroll1.Visible = True


Case 6
cmdA0.Visible = True: cmdgobacktop = 480

cmdsimplifytop = 480: cmdsimplify.Visible = True
Call printgoback
End Select

End Sub
Private Sub HScroll1_KeyPress(KeyAscii As Integer)
txtin.SetFocus
End Sub

Private Sub picTable_Click()
Select Case aIndex
Case 14
'Form1.Scale (-100, 200)-(200, -100)
Dim n As Byte
Dim z, az, angle, ar, pi, wx As Single

 pi = 3.141593: picTable.Cls
 picTable.Scale (-200, 200)-(200, -200)
 sc = 90 'saturation

 For n = 1 To 40
  ared(n) = 255 * Rnd
  agrn(n) = ared(n) + Rnd * sc - sc / 2.9
  ablu(n) = agrn(n) + Rnd * sc - sc / 2.9

  scl = 230 'larger value = more sparse cylinders
  X(n) = Rnd * scl - scl / 2
  Y(n) = Rnd * scl - scl / 2

  r(n) = 7 * Rnd + 10 'varying radii

  cs = 2 'colorshift intensity

  colorshift = Rnd * cs - cs / 2
  incred(n) = colorshift
  incgrn(n) = colorshift
  incblu(n) = colorshift
 Next n

 For z = -70 To 400 Step 25
  ar = 1250 / (1250 - z)
  For n = 1 To 40

   Call straightencolor(ared(n))
   Call straightencolor(agrn(n))
   Call straightencolor(ablu(n))

   For a = 0 To 2 * pi Step 2 * pi / 9
    picTable.PSet (ar * (r(n) * Sin(a) + X(n)), _
    ar * (r(n) * Cos(a) + Y(n))), _
    RGB(ared(n), agrn(n), ablu(n))
   Next a

   ared(n) = ared(n) - incred(n)
   agrn(n) = agrn(n) - incgrn(n)
   ablu(n) = ablu(n) - incblu(n)

  Next n
 Next z
End Select
End Sub
Private Sub straightencolor(ByRef color As Integer)
 If color < 0 Then
  color = color + 255: End If
 If color > 255 Then
  color = color - 255: End If
End Sub
Private Sub txtin_KeyPress(k As Integer)
Select Case subj
Case 1
If k > 255 Or k < 32 Then
 k = 32 'error trap
 lblcom.Caption = "Keypress out of range"
Else
 HScroll1.Value = k
 lblcom.Caption = q & Chr(k) & q & " has ANSI value " & HScroll1.Value
End If
 txtin.Text = ""
 
Case 7
Select Case npage
 Case 2
  Dim tn As Integer
  Dim teamName(1 To 5) As String
  teamName(1) = "Red Sox"
  teamName(2) = "Giants"
  teamName(3) = "White Sox"
  teamName(4) = "Cubs"
  teamName(5) = "Cubs"
  tn = k - 48
  If tn < 1 Or tn > 5 Then
   lblcom.Caption = ""
  Else
   lblcom.Caption = "The " & teamName(tn) & " won World Series " & tn
  End If
  txtin.Text = ""
 
 Case 3
  prevtext = txtin.Text
 End Select
End Select
End Sub
Private Sub txtin_change()
Select Case npage
Case 3
 Dim ascc As String
 Dim lentxtin As Byte
 lentxtin = Len(txtin.Text)
 For fn1 = 1 To lentxtin
  ascc = Asc(Mid(txtin.Text, fn1, 1))
  If ascc > 57 Or ascc < 48 Then
   txtin.Text = prevtext: Exit For: End If
 Next fn1
 If Val(txtin.Text) > 50000 Then
  txtin.Text = 50000: End If
 If Val(txtin.Text) > 49 Then
  Call cmdrun_Click: End If
End Select
End Sub
Private Sub cmdboxoffice_Click()

popcorn = 50
soda = 80
lblcom.Caption = FormatCurrency(popcorn + soda)

End Sub
Private Sub cmd1mv1_Click()
mv1 = mv1 + 1 'modular var demonstration
lblcom.Caption = mv1
End Sub
Private Sub cmd2mv1_Click()
mv1 = mv1 + 2
lblcom.Caption = mv1
End Sub
Private Sub cmd1lv1_Click()
lv1 = lv1 + 1 'local var demonstration
lblcom.Caption = lv1
End Sub
Private Sub cmd2lv1_Click()
Static lv1 As Integer
lv1 = lv1 + 2
lblcom.Caption = lv1
End Sub
Private Sub cmdrun_Click()
Select Case aIndex
Case 0
 Select Case subj
 Case 0
  lblqcom.Caption = "Value:"
  str1 = " yo."
  a = 2
  B = 3
  lblcom.Caption = B ^ (a + 1) & str1
 Case 4
  Dim ndimes As Integer
  Dim loosechange As Currency: loosechange = 0.45
  Call computedimes((loosechange), ndimes)
  lblcom.Caption = "There is a maximum of " & ndimes & " dimes in " & FormatCurrency(loosechange) & "."
 Case 7
  Select Case npage
  Case 1
  Dim lrnd As Single
  Dim localsum As Single
  picTable.Cls
  For fn1 = 1 To 3
   lrnd = Rnd * 12
   localsum = localsum + lrnd
   picTable.Print fn1; lrnd
  Next fn1
  picTable.Print "Final lcount", fn1
  picTable.Print "Rnd Average", localsum / (fn1 - 1)
 Case 3
  Dim balance As Single, numYears As Integer
  txtin.Text = Val(txtin.Text)
  If txtin.Text = "" Or txtin.Text < 50 Then
   txtin.Text = 50: End If
  balance = Val(txtin.Text)
  numYears = 0
  Do While balance < 1000000
   balance = balance + 0.07 * balance
   numYears = numYears + 1
  Loop
  picTable.Cls
  picTable.Print "About "; numYears; "years."
  End Select: End Select

Case 12
picTable.Cls
picTable.Scale (-1.8, 11)-(11, -1.8)
picTable.Line (0, 0)-(0, 10)
picTable.Line (0, 0)-(10, 0)
For n1 = 0 To 10
 picTable.Line (n1, 0.2)-(n1, -0.3), vbBlue
 picTable.CurrentX = n1 - picTable.TextWidth(n1)
 picTable.CurrentY = -0.3
 picTable.Print n1
 qy = Int(2.8 * n1 - 10)
 If qy >= 0 And qy <= 10 Then
  picTable.Line (-0.2, qy)-(0.2, qy), vbMagenta
  picTable.CurrentX = -1.4
  picTable.CurrentY = qy - picTable.TextHeight(qy) / 2
  picTable.Print qy
  End If: Next n1

Case 13
Dim c As Single
 picTable.Scale (1, 1)-(60, 60)
 picTable.Cls
 c = 2 * 3.14159
 a = 0.0000001
 B = 0.4
 picTable.FillStyle = 5
 picTable.FillColor = vbBlue
 picTable.Circle (30, 30), 20, , -a * c, -B * c
 picTable.FillStyle = 1
End Select

End Sub
Private Sub computedimes(loosechange As Currency, ndimes As Integer)
ndimes = Int(loosechange * 10)
End Sub
Private Sub cmdcopy_Click()
 piccode.Visible = False
 txtcode.Visible = True
 qs(1) = "Private Sub LoadFavoriteQuotes()"
qs(2) = " Dim lstrquote(1 To 100) As String"
  qs(3) = " Dim lstrauthor(1 To 100) As String"
  qs(4) = " Dim lbytindx As Byte: lbytindx = 1"
 qs(5) = " Open " & q & "c:\favquots.txt" & q & " For Input As #1"
 qs(6) = " 'Input into 'Buffer 1'"
 qs(7) = " 'EOF(buffer 1)(line below) is a handy 'End Of File' function"
 qs(8) = " Do While Not EOF(1)"
 qs(9) = "  Input #1, lstrquote(lbytindx), lstrauthor(lbytindx)"
 qs(10) = "  Call randomjumble(lstrquote(lbytindx), lstrauthor(lbytindx))"
 qs(11) = "  If LCase(lstrauthor(lbytindx)) = einstein Then"
 qs(12) = "   Exit Do  'A quote from Einstein is a good place to stop."
 qs(13) = "  End If"
 qs(14) = "  lbytindx = lbytindx + 1  'increment to next array slot"
 qs(15) = " Loop"
 qs(16) = " Close #1"
 qs(17) = "End Sub"
 qs(18) = ""
 'qs(18) = "Exit Do is a nice option."
 qs(19) = qps & "randomjumble(ByRef quote As String, ByRef author As String)"
 qs(20) = " Dim lsngrandom As Single, lstrjumble As String"
 qs(21) = " lstrjumble = " & q & q
 qs(22) = " Dim lintcount As Integer"
 qs(23) = " For lintcount = 1 To Len(author + quote)"
 qs(24) = "  lsngrandom = Rnd * 10  'deciding variable for lower or uppercase"
 qs(25) = "  If lsngrandom < 5 Then  'half of the time, we're making lowercase"
 qs(26) = "   lstrjumble = lstrjumble + LCase(Mid(author + quote, lintcount, 1))"
 qs(27) = "  Else"
 qs(28) = "   lstrjumble = lstrjumble + UCase(Mid(author + quote, lintcount, 1))"
 qs(29) = "  End If   'Adds 1 character from author+quote string, randomly upper or lower-cased"
 qs(30) = " Next lintcount   'each time."
 qs(31) = " picTable.Print q & Right(lstrjumble, Len(quote) & q & " & q & " - " & q & " & Left(lstrjumble, Len(author))"
 qs(32) = "'perhaps remember that i have  q = chr(34) = double quote"
 qs(33) = "End Sub"
 txtcode.Text = """"
 For mv1 = 1 To 33
  txtcode.Text = txtcode.Text & qs(mv1) & vbCrLf
 Next mv1

End Sub
Private Sub cmdrnd_click()
Select Case subj
Case 1
picTable.Cls
picTable.Print
picTable.Print Int(10 * Rnd);
picTable.Print 10 * Rnd
End Select
Select Case aIndex
Case 7
 picTable.Cls
 Dim lstrquote(1 To 100) As String
 Dim lstrauthor(1 To 100) As String
 Dim lbytindx As Byte
 For lbytindx = 1 To 6
  Call randomjumble(qs(lbytindx), qs(lbytindx + 6))
 Next lbytindx
' This works - keep for ref.
' Open "c:\favquots.txt" For Input As #1
' Do While Not EOF(1)
'  Input #1, lstrquote(lbytindx), lstrauthor(lbytindx)
'  Call randomjumble(lstrquote(lbytindx), lstrauthor(lbytindx))
'  If LCase(lstrauthor(lbytindx)) = "einstein" Then
'   Exit Do
'  End If
'  lbytindx = lbytindx + 1
' Loop
' Close #1
End Select
End Sub
Private Sub randomjumble(quote As String, author As String)
 Dim lsngrandom As Single, linttotallength As Integer
 linttotallength = Len(author) + Len(quote)
 Dim lstrjumble As String
 lstrjumble = ""
 Dim lintcount As Integer
 For lintcount = 1 To linttotallength
  lsngrandom = Rnd * 10
  If lsngrandom < 5 Then
   lstrjumble = lstrjumble + LCase(Mid(author + quote, lintcount, 1))
  Else
   lstrjumble = lstrjumble + UCase(Mid(author + quote, lintcount, 1))
  End If
 Next lintcount
 picTable.Print q & Right(lstrjumble, Len(quote)) & q & " - " & Left(lstrjumble, Len(author))
 
End Sub
Private Sub cmdsimplify_Click()
Select Case subj
Case 1
txtin.SetFocus
End Select

Call printmore

cmdsimplify.Visible = False
cmdgoback.Visible = True
End Sub
Private Sub printmore()
Select Case subj
Case 1 'called by codecommentary5()
lblmain.Caption = "This page can be used as a reference tool.  Type some keys, try the scrollbar!"
piccode.Cls
piccode.Print "Dim q as Byte"
piccode.Print "       "
piccode.Print "Private Sub txtin_Keypress(K As Integer)"
piccode.Print " q = Chr(34)"
piccode.Print " Dim qhas As String"
piccode.Print " qhas = " & q & " has ANSI Value " & q
piccode.Print " HScroll1.Value = K"
piccode.Print " lblcom.Caption = q & Chr(K) & q & qhas & HScroll1.Value"
piccode.Print " txtin.Text = " & q & q
piccode.Print "End Sub"

Case 4 'called by codecommentary2()
 Select Case npage
 Case 2
  piccode.Cls
  piccode.Print "Private Sub cmdrun_Click()"
  piccode.Print " Dim ndimes As Integer"
  piccode.Print " Dim loosechange As Currency: loosechange = 0.45"
  piccode.Print " Call computedimes(";: Call qscolor("loosechange", blue, 1): piccode.Print ",";: Call qscolor(" ndimes", mag, 1): piccode.Print ")"
  piccode.Print " .."
  piccode.Print " .."
  piccode.Print "End Sub"
  piccode.Print
  piccode.Print "Private Sub computedimes(";: Call qscolor("chg", blue, 1): piccode.Print " As Currency, ";: Call qscolor("dimez", mag, 1): piccode.Print " As Integer)"
  piccode.Print " .."
  piccode.Print "End Sub"
  cmdrun.Visible = False
  cmdruntop = 1500
  Call resizeall
  lblcom.Caption = "loosechange and chg .. same running lane (1 of 2), same Type (Currency), same baton (memory location)."
 Case 4
  piccode.Cls
  piccode.Print "Private Sub cmdrun_click()"
  piccode.Print " Dim numberofdimes as Integer"
  piccode.Print " Dim loosechange as Currency: loosechange = 0.45"
  piccode.Print " Call calculatedimes(numberofdimes, (loosechange))"
  piccode.Print " .."
  piccode.Print " .."
  piccode.Print " .."
  piccode.Print "Private Sub calculatedimes(ndime as Integer, chg as Currency)"
  piccode.Print
  lblcom.Caption = "Put extra ( ) around the choice variable in the arguments, or .. (click 'Unmore..')"
 End Select

Case 5
 lblmain.Caption = "A Function Procedure sembles a Sub Procedure, but .. see that?  Like a variable, a function holds a value.  'Area' will hold a value of type Integer."
 piccode.Cls
 piccode.Print: piccode.Print: piccode.Print: piccode.Print
 piccode.Print: piccode.Print: piccode.Print: piccode.Print
 'piccode.Print "Private Sub HScommon()"
 'piccode.Print '" Dim x As Single, y As Single"
 'piccode.Print '" x = HScroll1.Value: y = HScroll1.Value"
 'piccode.Print " lblcom.Caption = ..Area(..): End Sub"
 'piccode.Print "Private Sub VScommon()"
 'piccode.Print '" Dim x As Single, y As Single"
 'piccode.Print '" x = HScroll1.Value: y = HScroll1.Value"
 'piccode.Print " lblcom.Caption = ..Area(..): End Sub"
 piccode.Print "Private Function ";: Call qscolor("Area(", blue, 1): piccode.Print ".......................................";: Call qscolor(") As Integer", blue, 0)
 piccode.Print " .."
 Call qscolor(" Area", blue, 1): piccode.Print " = ...... : End Function"
 
Case 6
 cmdcd5.SetFocus
 lblmain.Caption = "I use Select Case in my page button routines.  In blue is a 'nested' Select Case."
 piccode.Cls: piccodewidth = 5000
 Call resizeall
 piccode.Print "Dim chapter As String"
 piccode.Print "Dim npage As Byte"
 piccode.Print " .."
 piccode.Print " .. subs between here change chapter and npage Values."
 piccode.Print " .."
 piccode.Print "Private Sub subjectmatter()"
 Call qscolor(" Select Case ", mag, 1): piccode.Print "chapter"
 Call qscolor(" Case ", mag, 1): piccode.Print q & "Variable types" & q: piccode.ForeColor = vbBlue
 piccode.Print "  Select Case ";: piccode.ForeColor = vbBlack: piccode.Print "npage": piccode.ForeColor = vbBlue
 piccode.Print "  Case";: piccode.ForeColor = vbBlack: piccode.Print " 1: lblmain.Caption = " & q & "Strings" & q: piccode.ForeColor = vbBlue
 piccode.Print "  Case";: piccode.ForeColor = vbBlack: piccode.Print " 2: lblmain.Caption = " & q & "Numeric" & q: piccode.ForeColor = vbBlue
 piccode.Print "  End Select": piccode.ForeColor = vbMagenta
 piccode.Print " Case ";: piccode.ForeColor = vbBlack: piccode.Print q & "Functions" & q
 piccode.Print "  lblmain.Caption = " & q & "Like Subs, Functions .." & q: piccode.ForeColor = vbMagenta
 piccode.Print " End Select";: piccode.ForeColor = vbBlack: piccode.Print ": End Sub"
  VScroll1.Visible = False
 End Select
End Sub
Private Sub printgoback()
Select Case subj
Case 1
lblmain.Caption = "HScroll uses another set of code."
piccode.Cls
piccode.Print "Dim q as Byte   'will be used to represent a double-quote (ANSI character 34)"
piccode.Print
piccode.Print "Private Sub txtin_Keypress(K As Integer)"
piccode.Print " q = Chr(34)"
piccode.Print " Dim qhas As String  'i use q as in qthis or qthat.."
piccode.Print " qhas = " & q & " has ANSI Value " & q
'piccode.Print " If K > 255 Or K < 32 Then  'remark  out-of-bounds HScrollbar Value test"
'piccode.Print "  K = 32  'remark: puts an out-of-bounds value in-bounds"
'piccode.Print " End If"
piccode.Print " HScroll1.Value = K"
piccode.Print " lblcom.Caption = q & Chr(K) & q & qhas & HScroll1.Value"
piccode.Print " txtin.Text = " & q & q & "   'this uhh, makes the text box hold only 1 character"
piccode.Print "End Sub"

Case 4
 Select Case npage
 Case 2
  piccode.Cls
  piccode.Print "Private Sub cmdrun_Click()"
  piccode.Print " Dim ndimes As Integer"
  piccode.Print " Dim loosechange As Currency: loosechange = 0.45"
  piccode.Print " Call computedimes(";: Call qscolor("loosechange", blue, 1): piccode.Print ",";: Call qscolor(" ndimes", mag, 1): piccode.Print ")"
  piccode.Print " lblcom.Caption = " & q & "There is a maximum of " _
   & q & "& ndimes &" & q & " dimes in " & q & " _"
  piccode.Print ; "  FormatCurrency(loosechange) & " & q & "." & q
  piccode.Print "End Sub"
  piccode.Print
  piccode.Print "Private Sub computedimes(";: Call qscolor("chg", blue, 1): piccode.Print " As Currency, ";: Call qscolor("dimez", mag, 1): piccode.Print " As Integer)"
  piccode.Print " dimez = Int(chg * 10)"
  piccode.Print "End Sub"
 
  cmdrun.Visible = True
 
  lblcom.Caption = "Call statement contains(arguments) while Sub contains(parameters)"
 Case 4
  piccode.Cls
  piccode.Print "Private Sub cmdrun_click()"
  piccode.Print " Dim numberofdimes as Integer"
  piccode.Print " Dim loosechange as Currency: loosechange = 0.45"
  piccode.Print " Call calculatedimes(numberofdimes, loosechange)"
  piccode.Print " .."
  piccode.Print " .."
  piccode.Print " .."
  piccode.Print "Private Sub calculatedimes(ndime as Integer, ByVal chg As Currency)"
  lblcom.Caption = " use ByVal in the parameters."
 End Select

Case 5
 lblmain.Caption = ""
 piccode.Cls
 piccode.Print "Private Sub HScommon()"
 piccode.Print " Dim x As Integer, y As Integer"
 piccode.Print " x = HScroll1.Value: y = HScroll1.Value"
 piccode.Print " lblcom.Caption = x & " & q & "*" & q & " & y & " & q & "=" & q & " & Area(x, y): End Sub"
 piccode.Print "Private Sub VScommon()"
 piccode.Print " Dim x As Integer, y As Integer"
 piccode.Print " x = HScroll1.Value: y = HScroll1.Value"
 piccode.Print " lblcom.Caption = x & " & q & "*" & q & " & y & " & q & "=" & q & " & Area(x, y): End Sub"
 piccode.Print "Private Function Area(x As Integer, y As Integer) As Integer"
 piccode.Print " Area = x * y"
 piccode.Print "End Function"

Case 6
lblmain.Caption = "Select Case uses expressions as does the If statement."
piccode.Visible = True: picTable.Visible = False
piccodetop = 900: piccodeleft = 450
piccodewidth = 4000: piccodeheight = 3000
Call resizeall
VScroll1.Min = 1: VScroll1.Max = 12
VScroll1.Visible = True: HScroll1.Visible = False
piccode.Cls
piccode.Print "Private Sub VScommon()"
piccode.Print " Dim caseinpoint As Byte, seasonname As String"
piccode.Print " caseinpoint = VScroll1.Value"
Call qscolor(" Select Case", blue, 1): piccode.Print " caseinpoint"
Call qscolor(" Case", blue, 1): piccode.Print " Is < 4"
piccode.Print "  seasonname = " & q & "Winter" & q
Call qscolor(" Case", blue, 1): piccode.Print " 4 To 6"
piccode.Print "  seasonname = " & q & "Spring" & q
Call qscolor(" Case", blue, 1): piccode.Print " 7, 8, 9"
piccode.Print "  seasonname = " & q & "Summer" & q
Call qscolor(" Case", blue, 1): piccode.Print " Else"
piccode.Print "  seasonname = " & q & "Fall" & q
Call qscolor(" End Select", blue, 0)
piccode.Print " lblchaptr.Caption = seasonname"
piccode.Print "End Sub"

End Select
End Sub

Private Sub cmdgoback_Click()
Call printgoback

Select Case subj
Case 1
txtin.SetFocus
Case 4
If firsttime(1) = 0 Then
firsttime(1) = 1
'cmdc6.Enabled = True
cmdcd3.Enabled = True
cmdcd4.Enabled = True
cmdcd5.Enabled = True
'cmdbc4.Enabled = True
End If
End Select

cmdgoback.Visible = False
cmdsimplify.Visible = True
End Sub
Private Sub lblalg_Click()
subj = 8
Call clearform
piccode.Visible = True
 
 If lblalg.Caption = "Algebra Reference" Then
'  If firsttime(4) = 0 Then
'   firsttime(4) = 1
'   'cmdvari.Enabled = False
'   End If
 picTable.Visible = False
 lblchaptr.Caption = "Algebra"
 lblqpage.Visible = True
 piccode.Visible = True
 cmdcd1.Visible = True
 cmdcd2.Visible = True
 cmdcd3.Visible = True
 cmdcd4.Visible = True
 'cmdcd5.Visible = True
 piccode.Cls: piccodeheight = 1860: piccodewidth = 3500
 piccodetop = 1440: piccodeleft = 220
 Call resizeall
 Call codecommentary1
 lblalg.Caption = "back"
 lblqprop.Caption = ""
Else
 Call clearform
 lblqpage.Visible = False
 lblalg.Caption = "Algebra Reference"
 Call contentspre
 lblqprop.Caption = "Properties"
 picTable.Visible = True
 lblqpage.Visible = False
 picTable.Visible = True
End If

End Sub
Private Sub HScroll1_Change()

Select Case subj
Case 0  'For message6()
lblchaptr.FontSize = HScroll1.Value
lblcom.Caption = HScroll1.Value
End Select

Call HScommon

End Sub
Private Sub HScommon()
'lblcom.Caption = HScroll1.Value
Select Case subj
Case 1  'For cmdcd5()
lblcom.Caption = q & Chr(HScroll1.Value) & q & " has ANSI value " & HScroll1.Value
txtin.Text = Chr(HScroll1.Value)
Case 2
 ltab = Int(HScroll1.Value)
 lblcom.Caption = ltab
 picTable.Cls
 lblqcom.Caption = "ltab:"
 picTable.Print qhello & 1; qhello
 picTable.Print qhello & 1; 1 + 1, qhello
 picTable.Print -1; 2
 picTable.Print 1,
 picTable.Print 12; Tab(ltab); qhello, 1
 picTable.Print qhello + qhello
 picTable.Print "Error - Type Mismatch"
 picTable.Print
 picTable.Print "The ; and , are not allowed for label or variable formatting."
 picTable.Print "A positive number in a label is not preceded by a space."

Case 5  'used by codecommentary5
Dim X As Integer, Y As Integer: X = HScroll1.Value
Y = VScroll1.Value
lblcom.Caption = X & "*" & Y & "=" & Area(X, Y)

Case 6
Select Case npage
 Case 2
 If HScroll1.Value < 50 Then
  lblcom.Caption = "   " & HScroll1.Value
 Else
  lblcom.Caption = HScroll1.Value
 End If
 lblqcom.Caption = ""
 If HScroll1.Value >= 20 And HScroll1.Value <= 30 Then
  lblqcom.Caption = "Rain": End If
 
 Case 4 'called by codecommentary4
 Dim t As Single
 If clscount > 14 Then
 picTable.Cls: clscount = 0: End If
 t = HScroll1.Value
 If t / 5 = Int(t / 5) Or _
  t / 4 = Int(t / 4) Then
 clscount = clscount + 1
 picTable.Print Tab(t); t;
 End If

 Case 5
 End Select
End Select

If pIndex = 6 Then
 If clicktwice(1) = 1 Then
  lblcom.Caption = HScroll1.Value
  lblchaptr.FontSize = HScroll1.Value
 End If
End If

End Sub
Private Sub HScroll1_Scroll()
Call HScommon
End Sub

Private Function Area(X As Integer, Y As Integer) As Integer
Area = X * Y
End Function

Private Sub VScroll1_LostFocus()
Select Case subj
Case 6
 lblchaptr.Caption = "Decisions"
End Select
End Sub
Private Sub VScroll1_Change()
Call VScommon
End Sub
Private Sub VScroll1_Scroll()
Call VScommon
End Sub
Private Sub VScommon()
Select Case aIndex
Case 0
 Select Case subj
 Case 5
  Dim X As Integer, Y As Integer: Y = VScroll1.Value
  X = HScroll1.Value
  lblcom.Caption = X & "*" & Y & "=" & Area(X, Y)
 Case 6
  Select Case npage
 Case 5
   caseinpoint = VScroll1.Value
   Select Case caseinpoint
   Case Is < 4: sn = "Winter"
   Case 4 To 6: sn = "Spring"
   Case 7 To 9: sn = "Summer"
   Case Is > 9: sn = "Fall"
   End Select
   lblchaptr.Caption = sn: End Select: End Select

Case 12, 14
 piccode.Cls
 For fn1 = VScroll1.Value To VScroll1.Value + 15
  piccode.Print qs(fn1): Next fn1

Case 15
 piccode.Cls
 For fn1 = VScroll1.Value To VScroll1.Value + 10
  If fn1 > 6 And fn1 < 10 Or fn1 > 17 And fn1 < 21 Then
   Call qscolor(qs(fn1), blue, 0)
  Else
   If fn1 = 5 Then
    Call qscolor(qs(fn1), mag, 0)
   Else
    If fn1 = 1 Or fn1 = 17 Or fn1 = 23 Then
     Call qscolor(qs(fn1), red, 0)
    Else
     piccode.Print qs(fn1): End If: End If: End If
 Next fn1
' qs(1) = "Dim prevtext As String"
' qs(2) = "Private Sub txtin_change()"
' qs(3) = "Dim asc1 As String"
' qs(4) = "Dim lentxtin As Byte"
' qs(5) = " lentxtin = Len(txtin.Text)"
' qs(6) = " For fn1 = 1 To lentxtin  'no non-numeric keystrokes"
' qs(7) = "  asc1 = Asc(Mid(txtin.Text, fn1, 1))"
' qs(8) = "  If asc1 > 57 Or asc1 < 48 Then"
' qs(9) = "   txtin.Text = prevtext: Exit For: End If"
' qs(10) = " Next fn1"
' qs(11) = " If Val(txtin.Text) > 50000 Then"
' qs(12) = "  txtin.Text = 50000: End If"
' qs(13) = " If Val(txtin.Text) > 49 Then"
' qs(14) = "  'cmdrun would calculate for each keystroke"
' qs(15) = "  Call cmdrun_Click: End If"
' qs(16) = "End Sub"
' qs(17) = "Private Sub txtin_KeyPress(k As Integer)"
' qs(18) = " 'txtin_Change executes before _KeyPress."
' qs(19) = " 'txtin_Change uses prevtext if keystroke"
' qs(20) = " 'is non-numeric"
' qs(21) = " prevtext = txtin.Text"
' qs(22) = "End Sub"
' qs(23) = "Private Sub cmdrun_Click()"
' qs(24) = "Dim balance As Single, numYears As Integer"
' qs(25) = " If txtin.Text = "" Or txtin.Text < 50 Then"
' qs(26) = "  txtin.Text = 50: End If"
' qs(27) = " .."
 'qs(28) = " balance = Val(txtin.Text)"
 'qs(29) = " numYears = 0"
 'qs(30) = " Do While balance < 1000000"
 'qs(31) = "  balance = balance + 0.07 * balance"
 'qs(32) = "  numYears = numYears + 1: Loop"
 'qs(33) = ""
 'qs(34) = " picTable.Cls"
 'qs(35) = " picTable.Print " & q & "About " & q & "; numYears; " & q & "years." & q
 'qs(36) = "End Sub"
 txtin.Visible = True
 txtin.SetFocus
 
End Select
End Sub

