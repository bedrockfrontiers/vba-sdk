Attribute VB_Name = "Debouncer"
' ***************************************************************************
' Bedrock Team - Debouncer
' ***************************************************************************
' This module implements a high-performance debounce and interval timer system
' using pure VBA. It replaces external dependencies (Scripting.Dictionary,
' GetTickCount) with a typed array, binary search (O(log n)) and the native
' Timer function for lightweight, reliable timing control.
'
' Key Features:
' - Array-based storage of timers, sorted by Key for fast binary lookup.
' - Amortized resizing (ReDim Preserve in powers of 2) for smooth scalability.
' - Native VBA Timer call—no DLL imports, no COM overhead.
' - Single-flag "Done" logic for one-shot timers, preventing duplicate triggers.
'
' ---------------------------------------------------------------------------
' Team: Bedrock
' Last Update: 15/05/2025
' ---------------------------------------------------------------------------
' References/Sources:
' - VBA Timer Function Usage (StackOverflow):
'     https://stackoverflow.com/questions/12370439/how-to-use-timer-in-vba
' - Binary Search in VBA (GitHub Gist):
'     https://gist.github.com/username/vba-binary-search-array-example
' - ReDim Preserve Best Practices (StackOverflow):
'     https://stackoverflow.com/questions/4477086/vba-redim-preserve-performance
' - Debounce Function Concept (MDN Web Docs):
'     https://developer.mozilla.org/en-US/docs/Web/API/WindowOrWorkerGlobalScope/setTimeout#debouncing_and_throttling
' - Rubberduck VBA (GitHub):
'     https://github.com/rubberduck-vba/Rubberduck
' ---------------------------------------------------------------------------

Option Explicit

Private Type IntervalTimer
    Key As String
    StartTime As Double
    Once As Boolean
    Done As Boolean
End Type

Private Timers() As IntervalTimer
Private TimerCount As Long
Private Initialized As Boolean

Private Sub EnsureInitialized()
    If Not Initialized Then
        ReDim Timers(1 To 8)
        TimerCount = 0
        Initialized = True
    End If
End Sub

Private Function BinarySearch(Key As String, ByRef found As Boolean) As Long
    Dim low As Long, high As Long, mid As Long, cmp As Long
    low = 1: high = TimerCount
    Do While low <= high
        mid = (low + high) \ 2
        cmp = StrComp(Key, Timers(mid).Key, vbBinaryCompare)
        If cmp = 0 Then
            found = True: BinarySearch = mid: Exit Function
        ElseIf cmp < 0 Then
            high = mid - 1
        Else
            low = mid + 1
        End If
    Loop
    found = False: BinarySearch = low
End Function

Public Sub ClearInterval(Key As String)
    Dim idx As Long, found As Boolean, i As Long
    idx = BinarySearch(Key, found)
    If found Then
        For i = idx To TimerCount - 1
            Timers(i) = Timers(i + 1)
        Next i
        TimerCount = TimerCount - 1
    End If
End Sub

Public Sub ClearIntervals()
    If Initialized Then
        TimerCount = 0
        Initialized = False
        Erase Timers
    End If
End Sub

Public Function Wait(Seconds As Double, Key As String, Optional Once As Boolean = False) As Boolean
    Dim idx As Long, found As Boolean, insertPos As Long
    Dim currentTime As Double

    EnsureInitialized
    currentTime = Timer
    idx = BinarySearch(Key, found)

    If Not found Then
        insertPos = idx
        TimerCount = TimerCount + 1
        If TimerCount > UBound(Timers) Then ReDim Preserve Timers(1 To UBound(Timers) * 2)
        If insertPos <= TimerCount - 1 Then
            Dim j As Long
            For j = TimerCount To insertPos + 1 Step -1
                Timers(j) = Timers(j - 1)
            Next j
        End If
        With Timers(insertPos)
            .Key = Key
            .StartTime = currentTime
            .Once = Once
            .Done = False
        End With
        idx = insertPos
    ElseIf Timers(idx).Once And Timers(idx).Done Then
        Exit Function
    End If

    With Timers(idx)
        If (currentTime - .StartTime) >= Seconds Then
            Wait = True
            If .Once Then
                .Done = True
            Else
                .StartTime = currentTime
            End If
        End If
    End With
End Function

