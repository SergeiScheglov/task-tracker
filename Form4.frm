VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   10830
   LinkTopic       =   "Form4"
   ScaleHeight     =   7215
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "O:\sheglov\TariffTasks\tasks.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3375
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5953
      _Version        =   393216
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Set cn1 = New ADODB.Connection
    cn1.ConnectionString = "Data Source=C:\TariffTasks\tasks.mdb;Persist Security Info=False"
    cn1.Open
Data1.RecordSource = "SELECT Employeename as [Сотрудник], Phonemobile as [Мобильный], Phonehome as [Домашний] FROM Employees ORDER BY Employeename DESC"
Data1.Refresh
cn1.Close
End Sub
