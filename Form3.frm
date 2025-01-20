VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Редактировать задание"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9045
   LinkTopic       =   "Form3"
   ScaleHeight     =   5835
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   480
      TabIndex        =   20
      Top             =   4320
      Width           =   8175
   End
   Begin VB.ComboBox Combo7 
      Height          =   315
      Left            =   6360
      TabIndex        =   18
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Добавить"
      Height          =   495
      Left            =   6720
      TabIndex        =   8
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   480
      TabIndex        =   7
      Top             =   1560
      Width           =   8175
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
      Height          =   285
      Left            =   480
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd. MM. yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   2895
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   6480
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   6480
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   3720
      TabIndex        =   1
      Top             =   2760
      Width           =   2175
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label Label11 
      Caption         =   "Оценка"
      Height          =   255
      Left            =   6360
      TabIndex        =   19
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "Примечание"
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Описание"
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Направление"
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Получено"
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   3240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Выполнено"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   3240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Тип задания"
      Height          =   255
      Left            =   6480
      TabIndex        =   12
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Приоритет"
      Height          =   375
      Left            =   6480
      TabIndex        =   11
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Статус"
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "Выполнивший сотрудник"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   2400
      Width           =   2775
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()

    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\TariffTasks\tasks.mdb;Persist Security Info=False"
cn.Open
mNote = Text2
mEmployeefulfilled = Combo6.Text
mGradename = Combo7.Text
mTaskcode = Form1.Adodc1.Recordset.Fields(0).Value
aDestinationname = Combo2.Text
aStatusname = Combo5.Text
aTasktypename = Combo3.Text
aPriorityname = Combo4.Text

'On Error GoTo BACK

cn.BeginTrans
cn.Execute "UPDATE Tasklist SET Notes = '" & mNote & "', Taskname = '" & Text1 & "', Taskfinishdate = Now(), Statusname = '" & aStatusname & "', Destinationname = '" & aDestinationname & "', Employeefulfilled = '" & mEmployeefulfilled & "', Gradename = '" & mGradename & "', Tasktypename = '" & aTasktypename & "' WHERE Taskcode = " & Str(mTaskcode) & ""
cn.CommitTrans
MsgBox "Информация записана"

'vbOK vbInformation, mTaskcode

cn.Close

    Unload Form3
    Form3.Hide
Exit Sub
'BACK:
'cn.RollbackTrans
'MsgBox "Ошибка! Информация не записана"

End Sub

