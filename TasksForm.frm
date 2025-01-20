VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   11100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Добавить задание"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Искать"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   2295
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Text            =   "Тип задания"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Text            =   "Статус"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Text            =   "Сотрудник"
      Top             =   840
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "Направление"
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    'Combo1.Text = rs1.Fields(0)
                      
    Do While Not rs1.EOF
       Combo1.AddItem rs1.Fields(0)
       rs1.MoveNext
    Loop
    
    rs1.Close
     
    On Error GoTo 0

End Sub
