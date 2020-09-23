VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmMain 
   Caption         =   " Memory Test --- API vs MMX/SSE ASM vs CRT (C RunTime)"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11355
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   11355
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame fmCache 
      Caption         =   "Cache"
      Height          =   1200
      Left            =   9270
      TabIndex        =   12
      Top             =   3603
      Width           =   1875
      Begin VB.PictureBox piicHolder 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   0
         Left            =   180
         ScaleHeight     =   645
         ScaleWidth      =   1410
         TabIndex        =   13
         Top             =   375
         Width           =   1410
         Begin VB.OptionButton optCache 
            Caption         =   "Pollute"
            Height          =   210
            Index           =   0
            Left            =   15
            TabIndex        =   15
            Top             =   0
            Width           =   1320
         End
         Begin VB.OptionButton optCache 
            Caption         =   "Populate"
            Height          =   210
            Index           =   1
            Left            =   15
            TabIndex        =   14
            Top             =   412
            Width           =   1320
         End
      End
   End
   Begin VB.ComboBox cboChart 
      Height          =   330
      ItemData        =   "frmMain.frx":030A
      Left            =   9270
      List            =   "frmMain.frx":0334
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   5235
      Width           =   1875
   End
   Begin MSChart20Lib.MSChart Chart 
      Height          =   6165
      Left            =   240
      OleObjectBlob   =   "frmMain.frx":03AD
      TabIndex        =   1
      Top             =   255
      Width           =   7740
   End
   Begin VB.Frame fmTest 
      Caption         =   "Test"
      Height          =   1200
      Left            =   9270
      TabIndex        =   3
      Top             =   821
      Width           =   1875
      Begin VB.PictureBox picHolder 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   705
         Index           =   0
         Left            =   180
         ScaleHeight     =   705
         ScaleWidth      =   1560
         TabIndex        =   6
         Top             =   375
         Width           =   1560
         Begin VB.OptionButton opTest 
            Caption         =   "FillMemory"
            Enabled         =   0   'False
            Height          =   270
            Index           =   1
            Left            =   0
            TabIndex        =   8
            Top             =   412
            Width           =   1560
         End
         Begin VB.OptionButton opTest 
            Caption         =   "CopyMemory"
            Height          =   270
            Index           =   0
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   1560
         End
      End
   End
   Begin VB.Frame fmAlign 
      Caption         =   "Data Alignment"
      Height          =   1200
      Left            =   9270
      TabIndex        =   2
      Top             =   2212
      Width           =   1875
      Begin VB.PictureBox piicHolder 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   660
         Index           =   1
         Left            =   180
         ScaleHeight     =   660
         ScaleWidth      =   1410
         TabIndex        =   9
         Top             =   375
         Width           =   1410
         Begin VB.OptionButton optAlign 
            Caption         =   "Aligned"
            Height          =   210
            Index           =   0
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   1320
         End
         Begin VB.OptionButton optAlign 
            Caption         =   "Aligned +1"
            Height          =   210
            Index           =   1
            Left            =   15
            TabIndex        =   10
            Top             =   412
            Width           =   1320
         End
      End
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Benchmark"
      Default         =   -1  'True
      Height          =   390
      Left            =   9270
      TabIndex        =   0
      Top             =   240
      Width           =   1875
   End
   Begin VB.Label lblChart 
      AutoSize        =   -1  'True
      Caption         =   "Chart type:"
      Height          =   210
      Left            =   9270
      TabIndex        =   5
      Top             =   4995
      Width           =   1080
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub RtlFillMemory Lib "kernel32" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

Private Const ONE_MB        As Long = (1024& * 1024&)

Private b1(1 To ONE_MB + 1) As Byte                         'Copy/Fill destination buffer
Private b2(1 To ONE_MB + 1) As Byte                         'Copy source buffer
Private b3(1 To ONE_MB + 1) As Byte                         'Cache pollution buffer
Private bAlign              As Boolean                      'Copy/Fill data alignment
Private bPollute            As Boolean                      'Cache populate/pollute
Private bCopy               As Boolean                      'Copy/Fill test
Private nHz                 As Currency                     'CPU frequency
Private nOverhead           As Currency                     'CpuClk overhead
Private nSIMD               As eSIMD_Support                'CPU MMX/SSE support
Private mem                 As cMemTest                     'MMX/SSE memory copy/fill class
Private bmk                 As cBenchMark                   'Benchmark class
Private crt                 As cCDECL                       'C cdecl class - call the C runtime library

Private Sub Form_Initialize()
  Call InitCommonControls
End Sub

Private Sub Form_Load()
  Dim c1 As Currency
  Dim c2 As Currency
  
  Set bmk = New cBenchMark                                  'Create benchmark class
  Set mem = New cMemTest                                    'MMX/SSE memory class
  Set crt = New cCDECL                                      'C run-time dll class
  
  Call crt.DllLoad("MSVCRT20")                              'Load the cdecl dll, MS Visual C Run-Time
  
  nHz = bmk.MHz * 1000000@                                  'Calc cpu frequency
  nSIMD = mem.SIMD_Support                                  'Determine the cpu's level of MMX/SSE SIMD support
  
  'Determine CpuClk overhead - twice to eliminate cache fx
  Call bmk.CpuClk(c1)
  Call bmk.CpuClk(c2)
  Call bmk.CpuClk(c1)
  Call bmk.CpuClk(c2)
  nOverhead = c2 - c1
  
  cboChart.ListIndex = 1
  opTest(0).Value = True
  optAlign(0).Value = True
  optCache(0).Value = True
End Sub

Private Sub Form_Resize()
  Dim nLeft As Long
  
  On Error Resume Next
    Call Chart.Move(240, 240, ScaleWidth - 2580, ScaleHeight - 480)
    nLeft = ScaleWidth - 2085
    cmdTest.Left = nLeft
    fmTest.Left = nLeft
    fmAlign.Left = nLeft
    fmCache.Left = nLeft
    lblChart.Left = nLeft
    cboChart.Left = nLeft
  On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set mem = Nothing
  Set bmk = Nothing
  Set crt = Nothing
End Sub

Private Sub cboChart_Click()
  Chart.chartType = cboChart.ItemData(cboChart.ListIndex)
End Sub

Private Sub cmdTest_Click()
  Dim i           As Long
  Dim j           As Long
  Dim nSum        As Long
  Dim nCol        As Long
  Dim nRow        As Long
  Dim nBlockSize  As Long
  
  Screen.MousePointer = vbHourglass
  DoEvents
  Chart.Repaint = False
  
  'Initially, 32k
  nBlockSize = &H8000&
  
  'For each blocksize, 32kB to 1MB
  For nRow = 1 To 6
    With Chart
      .Row = nRow
      
      'For each test type
      For nCol = 1 To .ColumnCount
        nSum = 0
        
        'Average over 10 tests
        For j = 1 To 10
          .Column = nCol
          If bCopy Then
            nSum = nSum + TestCopyMemory(nCol, nBlockSize)
          Else
            nSum = nSum + TestFillMemory(nCol, nBlockSize)
          End If
        Next j
        
        .Data = nSum / 10
      Next nCol
    End With
    
    nBlockSize = nBlockSize * 2
  Next nRow
  
  Chart.Repaint = True
  Screen.MousePointer = vbDefault
End Sub

Private Sub optAlign_Click(Index As Integer)
  bAlign = (Index = 0)
End Sub

Private Sub optCache_Click(Index As Integer)
  bPollute = (Index = 0)
End Sub

Private Sub opTest_Click(Index As Integer)
  Dim nCol As Long
  Dim nRow As Long
  
  bCopy = (Index = 0)
  
  With Chart
    If bCopy Then
      .ColumnCount = 5
      .Column = 1
      .ColumnLabel = "RtlMoveMemory"
      .Column = 2
      .ColumnLabel = "CopyMemMMX"
      .Column = 3
      .ColumnLabel = "CopyMemSSE"
      .Column = 4
      .ColumnLabel = "memmove"
      .Column = 5
      .ColumnLabel = "memcpy"
    Else
      .ColumnCount = 4
      .Column = 1
      .ColumnLabel = "RtlFillMemory"
      .Column = 2
      .ColumnLabel = "FillMemMMX"
      .Column = 3
      .ColumnLabel = "FillMemSSE"
      .Column = 4
      .ColumnLabel = "memset"
    End If
    
    For nCol = 1 To .ColumnCount
      For nRow = 1 To .RowCount
        .Column = nCol
        .Row = nRow
        .Data = 0
      Next nRow
    Next nCol
  End With
End Sub

Private Function TestCopyMemory(nTest As Long, nBlockSize As Long) As Long
  Dim n   As Long
  Dim c1  As Currency
  Dim c2  As Currency
  
  If bAlign Then
    n = 1
  Else
    n = 2
  End If
  
  If bPollute Then
    'Pollute the cache - should work for all but the P4-EE
    Call RtlFillMemory(ByVal VarPtr(b3(n)), ONE_MB, &HAA)
  Else
    'Populate the cache
    Select Case nTest
      Case 1: Call RtlMoveMemory(ByVal VarPtr(b1(n)), ByVal VarPtr(b2(n)), nBlockSize)
      Case 2: Call mem.CopyMemMMX(ByVal VarPtr(b1(n)), ByVal VarPtr(b2(n)), nBlockSize)
      Case 3: Call mem.CopyMemSSE(ByVal VarPtr(b1(n)), ByVal VarPtr(b2(n)), nBlockSize)
      Case 4: Call crt.CallFunc("memcpy", VarPtr(b1(1)), VarPtr(b2(1)), nBlockSize)
      Case 5: Call crt.CallFunc("memmove", VarPtr(b1(1)), VarPtr(b2(1)), nBlockSize)
    End Select
  End If
  
  Select Case nTest
  Case 1 'API
    Call bmk.CpuClk(c1)
      Call RtlMoveMemory(ByVal VarPtr(b1(n)), ByVal VarPtr(b2(n)), nBlockSize)
    Call bmk.CpuClk(c2)
  
  Case 2 'MMX
    If nSIMD > esNone Then
      Call bmk.CpuClk(c1)
        Call mem.CopyMemMMX(ByVal VarPtr(b1(n)), ByVal VarPtr(b2(n)), nBlockSize)
      Call bmk.CpuClk(c2)
    Else
      TestCopyMemory = 0
      Exit Function
    End If
    
  Case 3 'SSE
    If nSIMD > esMMX Then
      Call bmk.CpuClk(c1)
        Call mem.CopyMemSSE(ByVal VarPtr(b1(n)), ByVal VarPtr(b2(n)), nBlockSize)
      Call bmk.CpuClk(c2)
    Else
      TestCopyMemory = 0
      Exit Function
    End If
  
  Case 4 'CRT memcpy
    Call bmk.CpuClk(c1)
      Call crt.CallFunc("memcpy", VarPtr(b1(1)), VarPtr(b2(1)), nBlockSize)
    Call bmk.CpuClk(c2)

  Case 5 'CRT memmove
    Call bmk.CpuClk(c1)
      Call crt.CallFunc("memmove", VarPtr(b1(1)), VarPtr(b2(1)), nBlockSize)
    Call bmk.CpuClk(c2)
  
  End Select
  
  c2 = (c2 - c1 - nOverhead) * 10000
  TestCopyMemory = nHz / c2 * CCur(nBlockSize) / CCur(ONE_MB)
End Function

Private Function TestFillMemory(nTest As Long, nBlockSize As Long) As Long
  Dim n   As Long
  Dim c1  As Currency
  Dim c2  As Currency
  
  If bAlign Then
    n = 1
  Else
    n = 2
  End If
  
  If bPollute Then
    'Pollute the cache
    Call RtlFillMemory(ByVal VarPtr(b3(n)), ONE_MB, &HAA)
  Else
    'Populate the cache
    Select Case nTest
      Case 1: Call RtlFillMemory(ByVal VarPtr(b1(n)), nBlockSize, &H55)
      Case 2: Call mem.FillMemMMX(ByVal VarPtr(b1(n)), nBlockSize, &HAA)
      Case 3: Call mem.FillMemSSE(ByVal VarPtr(b1(n)), nBlockSize, &HAA)
      Case 4: Call crt.CallFunc("memset", VarPtr(b1(1)), 0, nBlockSize)
    End Select
  End If
  
  Select Case nTest
  Case 1 'API
    Call bmk.CpuClk(c1)
      Call RtlFillMemory(ByVal VarPtr(b1(n)), nBlockSize, &H55)
    Call bmk.CpuClk(c2)
  
  Case 2 'MMX
    If nSIMD > esNone Then
      Call bmk.CpuClk(c1)
        Call mem.FillMemMMX(ByVal VarPtr(b1(n)), nBlockSize, &HAA)
      Call bmk.CpuClk(c2)
    Else
      TestFillMemory = 0
      Exit Function
    End If
    
  Case 3 'SSE
    If nSIMD > esMMX Then
      Call bmk.CpuClk(c1)
        Call mem.FillMemSSE(ByVal VarPtr(b1(n)), nBlockSize, &HAA)
      Call bmk.CpuClk(c2)
    Else
      TestFillMemory = 0
      Exit Function
    End If
    
  Case 4 'CRT memset
    Call bmk.CpuClk(c1)
      Call crt.CallFunc("memset", VarPtr(b1(1)), 0, nBlockSize)
    Call bmk.CpuClk(c2)
  
  End Select
  
  c2 = (c2 - c1 - nOverhead) * 10000
  TestFillMemory = nHz / c2 * CCur(nBlockSize) / CCur(ONE_MB)
End Function
