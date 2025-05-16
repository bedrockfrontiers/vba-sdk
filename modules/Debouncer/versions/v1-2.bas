Attribute VB_Name = "Temporizer"
Option Explicit

Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

Private Intervals As Dictionary, OnceIntervals As Dictionary

Private Function tick_count() As Double
    tick_count = GetTickCount / 1000
End Function

Public Sub ClearInterval(Key As String)
    If Intervals.Exists(Key) Then
        Intervals.Remove (Key)
    End If
End Sub

Public Sub ClearIntervals()
    If Intervals Is Nothing Then
        Exit Sub
    Else
        If Not OnceIntervals Is Nothing Then
            OnceIntervals.RemoveAll
        End If
        Intervals.RemoveAll
    End If
End Sub

Public Function Wait(Seconds As Double, Key As String, Optional Once As Boolean = False) As Boolean
    If Intervals Is Nothing Then
        Set Intervals = New Dictionary
        Set OnceIntervals = New Dictionary
    End If
    
    If Not Intervals.Exists(Key) Then
        Intervals.Add Key, tick_count
        If Once And Not OnceIntervals.Exists(Key) Then
            OnceIntervals.Add Key, False
        End If
    End If
    
    If Once Then
        If OnceIntervals(Key) Then
            Exit Function
        End If
    End If
    
    If (tick_count - Intervals(Key)) >= Seconds Then
        If Once Then
            OnceIntervals(Key) = True
        End If
        ClearInterval (Key)
        Wait = True
    End If
End Function
