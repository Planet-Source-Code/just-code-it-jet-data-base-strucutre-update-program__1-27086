VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Upgrade"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   Icon            =   "upgrade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmd 
      Left            =   2520
      Top             =   3540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdTarget 
      Caption         =   "..."
      Height          =   255
      Left            =   6420
      TabIndex        =   8
      Top             =   420
      Width           =   315
   End
   Begin VB.CommandButton cmdSource 
      Caption         =   "..."
      Height          =   255
      Left            =   6420
      TabIndex        =   7
      Top             =   60
      Width           =   315
   End
   Begin VB.TextBox txtTarget 
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Top             =   360
      Width           =   5655
   End
   Begin VB.TextBox txtSource 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   60
      Width           =   5655
   End
   Begin VB.CommandButton cmdUpgrade 
      Caption         =   "Upgrade"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   3420
      Width           =   915
   End
   Begin VB.ListBox lstTbl 
      Height          =   2595
      Left            =   60
      TabIndex        =   1
      Top             =   720
      Width           =   6675
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   3420
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Target"
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Source"
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim dbt As Database
Dim dbs As Database

Dim miAuto%



Private Sub cmdClose_Click()
    Unload Me
End Sub

Sub chkqry()
Dim l%, n%, flg%
Dim qdfNew As QueryDef
           
    For n = 0 To dbs.QueryDefs.Count - 1      'loop through source database queries
        flg = False                           'set found query flag to false
        lstTbl.AddItem "Checking Query - " & dbs.QueryDefs(n).Name, 0
        For l = 0 To dbt.QueryDefs.Count - 1                         'loop through target database queries
            If dbs.QueryDefs(n).Name = dbt.QueryDefs(l).Name Then     'if found check sql
                If dbs.QueryDefs(n).SQL <> dbt.QueryDefs(l).SQL Then 'if sql different then update target sql with source sql
                    lstTbl.AddItem "***** Updated Query - " & dbs.QueryDefs(n).Name, 0
                    dbt.QueryDefs(l).SQL = dbs.QueryDefs(n).SQL
                End If
                flg = True          'flag query as existing
                Exit For            'exit for loop
            End If
        Next l
        '***** add new query if it doesn't exist in slave
        If Not flg And Left(dbs.QueryDefs(n).Name, 4) <> "MSys" Then 'add a new query
            lstTbl.AddItem "Created Query - " & dbs.QueryDefs(n).Name, 0
            Set qdfNew = dbt.CreateQueryDef(dbs.QueryDefs(n).Name, dbs.QueryDefs(n).SQL)
        End If
        DoEvents
    Next n

End Sub
Sub chktbl()
Dim l%, n%, flg%, sconnect$
    
    sconnect = "[;database=" & dbs.Name & "]."
    For n = 0 To dbs.TableDefs.Count - 1        'loop through source database queries
        flg = False                             'set found query flag to false
        lstTbl.AddItem "Checking Table - " & dbs.TableDefs(n).Name, 0
        For l = 0 To dbt.TableDefs.Count - 1                         'loop through target database queries
            If dbs.TableDefs(n).Name = dbt.TableDefs(l).Name Then     'if found check sql
                Call chkfld(dbt.TableDefs(l).Name)
                flg = True                      'flag query as existing
                Exit For                        'exit for loop
            End If
        Next l
        '***** add new query if it doesn't exist in target
        If Not flg And Left(dbs.TableDefs(n).Name, 4) <> "MSys" Then 'add a new query
            lstTbl.AddItem "***** Created Table - " & dbs.TableDefs(n).Name, 0
            dbt.Execute "select * into " & dbs.TableDefs(n).Name & " from " & sconnect & dbs.TableDefs(n).Name
        End If
        DoEvents
    Next n

End Sub

Sub chkfld(tb$)
Dim tdt As TableDef
Dim tds As TableDef
Dim l%, n%, flg%, maxloop%, wrd$

    Set tdt = dbt.TableDefs(tb)
    Set tds = dbs.TableDefs(tb)
    
    For n = 0 To tds.Fields.Count - 1
        flg = False
        For l = 0 To tdt.Fields.Count - 1
            If tds.Fields(n).Name = tdt.Fields(l).Name Then
                flg = True                                  'fld found so drop out
                Exit For
            End If
        Next l
        '***** add a new field if it doesn't exist
        If Not flg And Left(tds.Fields(n).Name, 4) <> "MSys" Then
            AppendDeleteField tdt, "APPEND", tds.Fields(n).Name, tds.Fields(n).Type, tds.Fields(n).Size
            lstTbl.AddItem "***** Created Field - " & tds.Fields(n).Name, 0
        End If
    Next n
    
    maxloop = tdt.Fields.Count - 1
    For n = 0 To maxloop
        If n > maxloop Then Exit For            'need this in case a column is dropped
        If n = 26 Then Stop
        flg = False
        For l = 0 To tds.Fields.Count - 1
            If tds.Fields(l).Name = tdt.Fields(n).Name Then
                flg = True                                  'fld found so drop out
                Exit For
            End If
        Next l
        '***** delete the field from the target table
        If Not flg And Left(tdt.Fields(n).Name, 4) <> "MSys" Then
            wrd = tdt.Fields(n).Name
            AppendDeleteField tdt, "DELETE", tdt.Fields(n).Name, tdt.Fields(n).Type, tdt.Fields(n).Size
            maxloop = maxloop - 1           'dropped a column so decrease maxloop
            lstTbl.AddItem "***** Deleted Field - " & wrd, 0
        End If
    Next n

End Sub

Sub AppendDeleteField(tdfTemp As TableDef, strCommand As String, strName As String, Optional varType, Optional varSize)
    
   With tdfTemp
      If .Updatable = False Then
         MsgBox "TableDef not Updatable! " & "Unable to complete task."
         Exit Sub
      End If
      If strCommand = "APPEND" Then
         .Fields.Append .CreateField(strName, varType, varSize)
      Else
         If strCommand = "DELETE" Then .Fields.Delete strName
      End If
   End With

End Sub

Private Sub cmdSource_Click()
    cmd.Filter = "*.mdb|*.mdb"
    cmd.FileName = ""
    cmd.ShowOpen
    If cmd.FileName = "" Then Exit Sub
    Set dbs = OpenDatabase(cmd.FileName)
    txtSource = dbs.Name
    
End Sub

Private Sub cmdTarget_Click()
    cmd.Filter = "*.mdb|*.mdb"
    cmd.FileName = ""
    cmd.ShowOpen
    If cmd.FileName = "" Then Exit Sub
    Set dbt = OpenDatabase(cmd.FileName)
    txtTarget = dbt.Name
End Sub

Private Sub cmdUpgrade_Click()
Dim r%, fl%, l%

    Set dbt = OpenDatabase(txtTarget)    'need to reset the database objects in case a table was added or dropped in last update
    Set dbs = OpenDatabase(txtSource)
        
    lstTbl.Clear
    If txtSource = "" Or txtTarget = "" Then
        r = MsgBox("You need to set a Source and Detsination database.", vbInformation)
        Exit Sub
    End If
    
    Screen.MousePointer = 11
        Call chktbl         'Check Tables
        Call chkqry         'Check Queries
    Screen.MousePointer = 0
    If Not miAuto Then
        r = MsgBox("Database Upgrade Complete.", vbInformation, "DataBase Upgrade")
    End If
    
    fl = FreeFile
    On Error Resume Next
    Kill App.Path & "\dbupgrade.log"
    On Error GoTo 0
    Open App.Path & "\dbupgrade.log" For Output As #fl
    For l = lstTbl.ListCount - 1 To 0 Step -1
        Print #fl, lstTbl.List(l)
    Next l
    Close #fl
    
End Sub

Private Sub Form_Load()
Dim wrd$, s%, e%, r%
    If Command$ <> "" Then  'must be in the format Source=filename;Target=filename;
                            'ie    Source=c:\my documents\db source.mdb;target=c:\my documents\db target.mdb;
        wrd = Command$
        txtSource = getwrd(1, 0, wrd, "Source=", ";")
        txtTarget = getwrd(1, 0, wrd, "Target=", ";")
        If Not isfile(txtSource) Or txtSource = "" Then
            r = MsgBox("The source file " & txtSource & " does not exist.", vbExclamation)
            End
        End If
        If Not isfile(txtTarget) Or txtTarget = "" Then
            r = MsgBox("The target file " & txtTarget & " does not exist.", vbExclamation)
            End
        End If
        miAuto = True
        Me.Hide
        Call cmdUpgrade_Click
        End
    End If
End Sub
Function getwrd(s%, e%, b$, swrd$, ewrd$) As String
Dim s1%
    s1 = s                              'hold old value incase can't find word and reverts to 0
    s = InStr(s, b, swrd) + Len(swrd)
    If s = Len(swrd) Then
        s = s1 + 1                      'set value back + 1
        getwrd = ""
        Exit Function
    End If
    e = InStr(s + 1, b, ewrd)
    If e = 0 Then
        e = s1
        getwrd = ""
        Exit Function
    End If
    getwrd = Mid(b, s, e - s)
    s = s + 1
End Function
Function isfile(fl$) As Integer
Dim a$
    a = Replace(fl, Chr(34), "")
    isfile = False
    If Dir(a) <> "" Then isfile = True
End Function
