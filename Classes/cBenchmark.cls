VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBenchMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=======================================================================================
' cBenchmark
'
Option Explicit

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private pMe           As Long                               'Address of this class instance's vtable
Private nEntry        As Long                               'vtable public method entry index
Private sCode()       As String                             'Machine code lives here
Private cCpuStart     As Currency                           'BenchStart cpu clock cycles
Private cCpuStop      As Currency                           'BenchStop cpu clock cycles
Private cQpcFreq      As Currency                           'QueryPerformance frequency
Private cQpcStart     As Currency                           'QueryPerformance start
Private cQpcStop      As Currency                           'QueryPerformance stop

Private m_MHz         As Long                               'CPU MHz

Private Sub Class_Initialize()
  Call QueryPerformanceFrequency(cQpcFreq)                  'Get the QueryPerformance frequency
  
  Call RtlMoveMemory(pMe, ByVal ObjPtr(Me), 4)              'Get the address of this class instance's vtable
  Call Redirect("0F318B4C2408890189510433C0C20800")         'Patch the CpuClk function
End Sub

Public Sub CpuClk(Cycles As Currency)
End Sub

'Benchmark start
Public Sub BenchStart()
  Call QueryPerformanceCounter(cQpcStart)
  Call CpuClk(cCpuStart)
End Sub

'Benchmark stop
Public Sub BenchStop(fSeconds As Double, cCpuCycles As Currency)
  Call QueryPerformanceCounter(cQpcStop)
  Call CpuClk(cCpuStop)
  
  fSeconds = CDbl((cQpcStop - cQpcStart) / cQpcFreq)
  cCpuCycles = (cCpuStop - cCpuStart) * 10000
End Sub

Public Property Get MHz() As Long
  Dim i       As Long
  Dim fSecs   As Double
  Dim cCycles As Currency

  If m_MHz = 0 Then
    Call BenchStart
      For i = 0 To 20000
        DoEvents
      Next i
    Call BenchStop(fSecs, cCycles)
    m_MHz = Round((cCycles / fSecs) / 1000000#, 0)
  End If
  
  MHz = m_MHz
End Property

Private Sub Redirect(ByVal sHexCode As String)
  Dim i     As Long
  Dim nLen  As Long
  Dim s     As String
  
  nLen = Len(sHexCode)
  
  For i = 1 To nLen Step 2
    s = s & ChrB$(Val("&H" & Mid$(sHexCode, i, 2)))
  Next i
  
  ReDim Preserve sCode(0 To nEntry)
  sCode(nEntry) = s
  Call RtlMoveMemory(ByVal pMe + &H1C + (nEntry * 4), StrPtr(sCode(nEntry)), 4)
  nEntry = nEntry + 1
End Sub
