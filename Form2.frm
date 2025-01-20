VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Новое задание"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9300
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4290
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3480
      TabIndex        =   17
      Text            =   "Combo6"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3960
      TabIndex        =   15
      Text            =   "Combo5"
      Top             =   600
      Width           =   2175
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6480
      TabIndex        =   13
      Text            =   "Combo4"
      Top             =   2640
      Width           =   2295
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6600
      TabIndex        =   11
      Top             =   600
      Width           =   2175
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   600
      TabIndex        =   10
      Top             =   600
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   600
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "m/d/yy h:nn"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   3480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "m/d/yy h:nn"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   3480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   8175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Добавить"
      Height          =   495
      Left            =   6720
      TabIndex        =   0
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "Выполнивший сотрудник"
      Height          =   255
      Left            =   3480
      TabIndex        =   18
      Top             =   2280
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label8 
      Caption         =   "Статус"
      Height          =   255
      Left            =   3960
      TabIndex        =   16
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Приоритет"
      Height          =   375
      Left            =   6480
      TabIndex        =   14
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Тип задания"
      Height          =   255
      Left            =   6600
      TabIndex        =   12
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Сотрудник"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Выполнено"
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Получено"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Направление"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Описание"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Description As String
Dim Taskdate As Date
Dim Taskestimatedfinishdate As Date


Private Sub Form_Load()
Set cn1 = New ADODB.Connection
    cn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\TariffTasks\tasks.mdb;Persist Security Info=False"
    cn1.Open
    
  
    
    On Error GoTo 0
   
    
    Set rs1 = New ADODB.Recordset
    rs1.LockType = adLockOptimistic
    Set rs1.ActiveConnection = cn1
    
    rs1.Source = "SELECT Destinationname, Destinationcode FROM Destinations ORDER BY Destinationname"
    rs1.Open
    rs1.MoveFirst
    'Combo2.Text = rs1.Fields(0).Value
                      
    Do While Not rs1.EOF
       Combo2.AddItem rs1.Fields(0).Value
       rs1.MoveNext
    Loop
    
    rs1.Close
     
       
    Set rs1 = New ADODB.Recordset
    rs1.LockType = adLockOptimistic
    Set rs1.ActiveConnection = cn1
    
    rs1.Source = "SELECT Employeename, Employeecode  FROM Employees ORDER BY Employeename"
    rs1.Open
    rs1.MoveFirst
    'Combo1.Text = rs1.Fields(0)
                      
    Do While Not rs1.EOF
       Combo1.AddItem rs1.Fields(0).Value
       rs1.MoveNext
    Loop
    
    rs1.Close

Set rs1 = New ADODB.Recordset
    rs1.LockType = adLockOptimistic
    Set rs1.ActiveConnection = cn1
    
    rs1.Source = "SELECT Statusname, Statuscode  FROM Status ORDER BY Statusname"
    rs1.Open
    rs1.MoveFirst
    Combo5.Text = rs1.Fields(0)
                      
    Do While Not rs1.EOF
       Combo5.AddItem rs1.Fields(0).Value
       rs1.MoveNext
    Loop
    
    rs1.Close


    'Тип задания
    Set rs1 = New ADODB.Recordset
    rs1.LockType = adLockOptimistic
    Set rs1.ActiveConnection = cn1
    
    rs1.Source = "SELECT Tasktypename, Tasktypecode  FROM Tasktypes ORDER BY Tasktypename"
    rs1.Open
    rs1.MoveFirst
    'Combo1.Text = rs1.Fields(0)
                      
    Do While Not rs1.EOF
       Combo3.AddItem rs1.Fields(0).Value
       rs1.MoveNext
    Loop
    
    rs1.Close
    
    
    'Приоритет
    Set rs1 = New ADODB.Recordset
    rs1.LockType = adLockOptimistic
    Set rs1.ActiveConnection = cn1
    
    rs1.Source = "SELECT Priorityname, Prioritiecode  FROM Priorities ORDER BY Priorityname"
    rs1.Open
    rs1.MoveFirst
    Combo4.Text = rs1.Fields(0)
                      
    Do While Not rs1.EOF
       Combo4.AddItem rs1.Fields(0).Value
       rs1.MoveNext
    Loop
    
    rs1.Close
    
    'Combo4.DataField (1)
    'Выполнивший сотрудник
    Set rs1 = New ADODB.Recordset
    rs1.LockType = adLockOptimistic
    Set rs1.ActiveConnection = cn1
    
    rs1.Source = "SELECT Employeename, Employeecode  FROM Employees ORDER BY Employeename"
    rs1.Open
    rs1.MoveFirst
    'Combo1.Text = rs1.Fields(0)
                      
    Do While Not rs1.EOF
       Combo6.AddItem rs1.Fields(0).Value
       rs1.MoveNext
    Loop
    
    rs1.Close
    
    
End Sub

Private Sub Command1_Click()

'Private Sub AddRecord_Click()

Description = Text1.Text
'Taskdate = Text3.Text
'Taskestimatedfinishdate = Text4.Text

Set cn1 = New ADODB.Connection
    cn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\TariffTasks\tasks.mdb;Persist Security Info=False"
    cn1.Open
On Error GoTo BACK
cn1.BeginTrans
cn1.Execute "INSERT INTO Tasklist (Taskname, Destinationname, Tasktypename, Priorityname, Statusname) VALUES ('" & Text1 & "', '" & Combo2.Text & "', '" & Combo3.Text & "', '" & Combo4.Text & "', '" & Combo5.Text & "')"
cn1.CommitTrans
'End Sub



Form2.Hide
Unload Me

MsgBox ("Запись успешно добавлена в базу данных")
cn1.Close
    Exit Sub
BACK:
MsgBox "Ошибка! Информация не записана"


End Sub

