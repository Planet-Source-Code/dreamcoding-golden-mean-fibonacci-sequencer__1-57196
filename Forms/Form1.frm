VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Fibonacci Number Sequencer ( Golden Mean / Ratio / Phi ) "
   ClientHeight    =   9135
   ClientLeft      =   -165
   ClientTop       =   1575
   ClientWidth     =   15135
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   15135
   Begin VB.TextBox txtTotal 
      Height          =   1455
      Left            =   5520
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   7560
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Help"
      Height          =   375
      Left            =   14040
      TabIndex        =   22
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cboCount 
      Height          =   315
      Left            =   11400
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox txtCurrentResult 
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Top             =   6720
      Width           =   13575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "View Golden Ratio"
      Height          =   255
      Left            =   4200
      TabIndex        =   11
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Max Number"
      Height          =   255
      Left            =   8040
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.ListBox lstLen 
      Height          =   5520
      Left            =   1200
      TabIndex        =   9
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtStarting 
      Height          =   375
      Left            =   9480
      TabIndex        =   8
      Top             =   8640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtProceed 
      Height          =   375
      Left            =   9480
      TabIndex        =   7
      Top             =   8280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtTemp 
      Height          =   375
      Left            =   9480
      TabIndex        =   6
      Top             =   7920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ListBox lstCount 
      Height          =   5520
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.ListBox lstGoldenMean 
      Height          =   1425
      Left            =   1440
      TabIndex        =   4
      Top             =   7560
      Width           =   2295
   End
   Begin VB.ListBox lstResults 
      Height          =   5520
      Left            =   2040
      TabIndex        =   2
      Top             =   1080
      Width           =   12975
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   1560
      MaxLength       =   308
      TabIndex        =   1
      Top             =   120
      Width           =   6375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go !"
      Default         =   -1  'True
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "What's so great about fibonacci numbers?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   8880
      TabIndex        =   27
      Top             =   7320
      Width           =   3135
   End
   Begin VB.Label Label12 
      Caption         =   "The ratio to the left is found by taking the proceeding number and dividing it by the initial number. "
      Height          =   1575
      Left            =   120
      TabIndex        =   26
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   $"Form1.frx":08CA
      Height          =   1815
      Left            =   3960
      TabIndex        =   25
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Total of All up to Selected:"
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
      Left            =   5520
      TabIndex        =   24
      Top             =   7320
      Width           =   2415
   End
   Begin VB.Label Label9 
      Caption         =   "(n)th Fibonacci  Num :"
      Height          =   255
      Left            =   9480
      TabIndex        =   21
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "The textboxes to the right are used to calculate large additions. They are kept invisible at runtime."
      Height          =   1335
      Left            =   8160
      TabIndex        =   19
      Top             =   7680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Note: Fibonacci sequence starts with 1,1,2,3,5,8... etc. In this example we start at 2."
      Height          =   375
      Left            =   6480
      TabIndex        =   18
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label6 
      Caption         =   "Current Result:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Golden Ratio"
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
      Left            =   1440
      TabIndex        =   15
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Fibonacci Numbers"
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Chr Len"
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Count"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Calculate to Num:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------
'Programmer:   Jeff Nickell (phishbowlerz@yahoo.com)
'Program Name: Fibonacci Number Sequencer (Golden Mean / Ratio)
'Created On:   11/10/2004
'
'Credit to:    Philipp Emanuel Weidmann for modLNHS.bas
'              which calculates very large numbers. Found on PSC.
'              ListBoxLocking function: Unknown
'              IsNumber function: Shon
'
'http://www.dreamstruct.com
'----------------------------------------------------------------------




Private Sub cboCount_Click()
Form1.txtInput = ""
End Sub

Private Sub Check1_Click()
'Resizes form to display the GoldenMean listbox (Golden Ratio 1.61803398874989)
If Form1.Check1.Value = 1 Then
Form1.Height = 9645
End If
If Form1.Check1.Value = 0 Then
Form1.Height = 7740
End If

End Sub

Private Sub Command1_Click()
On Error GoTo Err1
Dim Fib As Double

'Clear all lists
lstResults.Clear
lstGoldenMean.Clear
lstCount.Clear
lstLen.Clear

'We are starting at the 3rd Fibonacci Number in the Sequence which is "2"
'1,1,2,3,5,8 etc. So we initialize the count variable (x) at 3, which is
'actually a list index of 2 since first item in count list is a list index of 0.
x = 2

'Are they using the count or nth fibonacci dropdown box?
If Form1.cboCount.ListIndex = 0 Then
'They are calculating using the input box
nthFib = False
Else
'They are calculating using the dropdown box
nthFib = True
End If

If nthFib = False Then

  'Make sure they entered a numerical value
  If IsNumber(Trim(Form1.txtInput)) = True Then
  Else
    msg = MsgBox("Input is not a Number!")
    Exit Sub
  End If
  
End If

'Check to Make Sure a Number to Calculate to Has Been Entered
If Form1.txtInput = "" And nthFib = False Then

msg = MsgBox("Need Num to Count", vbOKOnly)
    Form1.txtInput.SetFocus
    Exit Sub
Else
    

    Form1.txtStarting = 1
    Form1.txtProceed = 1
    
  Form1.Caption = "Please Wait.. Calculating Sequence"
    Do
      txtTemp = txtProceed
      
      'Add the Proceeding Number with the Number we are currently on.
      'Note: Fibonacci Numbers are found by taking the proceeding number
      '      and adding it to the current number.
      '      So 1+1=2, 2+3=5, 3+5=8, 5+8=13, 8+13=21, 13+21=34 ...etc
      '         /      /     /       /      /
      '      1,1,    2,  3,    5   8,   13,     21,   34  (Fibonacci Sequence)
      '
      '      So sequence generated by program will look like:
      '      1,1,2,3,5,8,13,21,34..etc
      
      
      txtProceed = LargeAdd(txtProceed, txtStarting)
      
      'The two lines below calculate the Golden Mean
      'This is done by taking the Proceeding number divided by the initial number.
      'And then the square root of the result to give approx 1.61803398874989
      'This is just for proof of concept, and is displayed only if you check
      'the checkbox in the program.
      
      
      
      Fib = txtProceed / txtStarting
      Form1.lstGoldenMean.AddItem Sqr(Fib)
      
      'Grab the Temp Number and Make it the New Starting Number
      txtStarting = txtTemp
      
      'x here is a counting variable to count what
      'fibonacci number we are on. 1st, 2nd, 3rd, etc.
      x = x + 1
      
      'Add the Count variable to the Count list box
      Form1.lstCount.AddItem x
      
      'Take the Length of the number that comes after the number we are currently on
      'And display this in the CHR Len list
      Form1.lstLen.AddItem Len(txtProceed)
      
      'Display the addition of proceeding + starting number in results list.
      lstResults.AddItem txtProceed
      
      'The iif statement below is an if then statement and based on whether
      'nthFib is True (nthFib being whether they are using the combo box) if it is True,
      'it will loop until the count variable has reached the selected combobox number.
      'If false it will loop until the calculation is either equal to or has surpassed
      'the number in the input textbox.
      
      
    Loop Until IIf(nthFib, x = Form1.cboCount.List(Form1.cboCount.ListIndex), Val(txtProceed) >= Val(Form1.txtInput))


End If

'Set initial list indexes for list locking
Form1.lstCount.ListIndex = 0
Form1.lstLen.ListIndex = 0
Form1.lstResults.ListIndex = 0

'Reset Status Label
Form1.Caption = "Fibonacci Number Sequencer ( Golden Mean / Ratio / Phi )"

GoTo ExitThis

Err1:
 msg = MsgBox(err.Description)
 msg = MsgBox("Too Big")
ExitThis:

End Sub

Private Sub Command2_Click()
Form1.txtInput = "99999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999"
Command1_Click
End Sub

Private Sub Command3_Click()
helpmsg = "There are two ways to go about the sequencing. You can either use the Calculate to Num textbox, or the nth Fibonacci Num dropdown box." & vbNewLine & vbNewLine
helpmsg = helpmsg & "The first option allows you to enter a number into the textbox and then hit Go! After you hit Go! number sequencing will start" & vbNewLine
helpmsg = helpmsg & "and end when the sequence reaches approximately the number you entered into the box." & vbNewLine
helpmsg = helpmsg & vbNewLine & "The second option, is to simply select a number from the dropdown box." & vbNewLine
helpmsg = helpmsg & "This will step through the sequence up to the nth number in the sequence you selected." & vbNewLine
helpmsg = helpmsg & "The max number for this (in this program) is 1476." & vbNewLine
helpmsg = helpmsg & vbNewLine & "You can easily switch back and forth between using the textbox and combo with out any problems." & vbNewLine
helpmsg = helpmsg & "The combobox will clear as soon as you type text, and the text will clear as soon as you select from the combobox." & vbNewLine
helpmsg = helpmsg & vbNewLine & "Select the View Golden Ratio textbox to view the actual golden ratio calculated for each number."
helpmsg = helpmsg & vbNewLine & "You will also see the Total Up To Selected. It is the sum of each number up to and including the selected."
helpmsg = helpmsg & vbNewLine & "Contact: phishbowlerz@yahoo.com"
msg = MsgBox(helpmsg, vbOKOnly, "Fibonacci Number Sequencer")


End Sub

Private Sub Form_Load()
'----------------------------------------------------------------------
'Programmer:   Jeff Nickell (phishbowlerz@yahoo.com)
'Program Name: Fibonacci Number Sequencer (Golden Mean / Ratio)
'Created On:   11/10/2004
'
'Credit to:    Philipp Emanuel Weidmann for modLNHS.bas
'              which calculates very large numbers. Found on PSC.
'              ListBoxLocking function: Unknown
'              IsNumber function: Shon
'
'http://www.dreamstruct.com
'-----------------------------------------------------------------------

'Initialize form size information
Form1.Left = 50
Form1.Height = 7740

cboCount.AddItem " "
For x = 3 To 1476 'Max nth number we can calculate for now (1476)
  cboCount.AddItem x
Next x
cboCount.ListIndex = 0

'We will display the entire form
Form1.Check1.Value = 1
End Sub

Private Sub Label13_Click()
frmExplain.Show
End Sub

Private Sub lstCount_Click()
'Lock listboxes lstResults and lstLen with lstCount
Call ListLock(Form1.lstResults, Form1.lstCount, True)
Call ListLock(Form1.lstLen, Form1.lstCount, True)
Call ListLock(Form1.lstGoldenMean, Form1.lstCount, True)
End Sub

Private Sub lstGoldenMean_Click()
Call ListLock(Form1.lstCount, Form1.lstGoldenMean, True)
Call ListLock(Form1.lstResults, Form1.lstGoldenMean, True)
Call ListLock(Form1.lstLen, Form1.lstGoldenMean, True)
End Sub

Private Sub lstLen_Click()
'Lock listboxes lstCount and lstResults with lstLen
Call ListLock(Form1.lstCount, Form1.lstLen, True)
Call ListLock(Form1.lstResults, Form1.lstLen, True)
Call ListLock(Form1.lstGoldenMean, Form1.lstLen, True)
End Sub

Private Sub lstResults_Click()
'Lock listboxes lstCount and lstLen with lstResults
Call ListLock(Form1.lstCount, Form1.lstResults, True)
Call ListLock(Form1.lstLen, Form1.lstResults, True)
Call ListLock(Form1.lstGoldenMean, Form1.lstResults, True)

'Copy the Current Selected Item to the Current Results Textbox
Form1.txtCurrentResult = Form1.lstResults.List(Form1.lstResults.ListIndex)

'Calculate the Total of all digits up to and including the selected number

'Take the 2nd entry after the selected item and subtract 1
'This is a known equation that works for all fibonacci numbers
'The draw back is that we will get an error for the last two items
'in our list since we have not calculated the next two numbers.
'For now I am just leaving this as it is.

OurIndex = Form1.lstResults.ListIndex
OurIndex = LargeAdd(OurIndex, 2)
TotalTally = LargeSubtract(Form1.lstResults.List(OurIndex), 1)
Form1.txtTotal = TotalTally

End Sub

Private Sub txtInput_Change()
If Len(txtInput) = 308 Then
  msg = MsgBox("Number input cannot be more than 308 digits long")
End If
If Len(Form1.txtInput) > 0 Then
  Form1.cboCount.ListIndex = 0
End If
End Sub

Public Function IsNumber(strNum As String) As Boolean

    Dim strChar As String

    While strNum <> vbNullString
        strChar = Left(strNum, 1)


        If IsNumeric(strChar) Then
            strNum = Right(strNum, Len(strNum) - 1)
        Else
            IsNumber = False
            Exit Function
        End If

    Wend

    IsNumber = True
End Function
