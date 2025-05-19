VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ***************************************************************************
' List - High-performance Dynamic Array for VBA
' ***************************************************************************
' This class module implements a dynamic array-like structure with fast
' resizing, insertion, removal, and search functionalities. It uses pure VBA
' and optimizes performance by managing capacity doubling and minimal memory
' reallocations.
'
' Key Features:
' - Dynamic capacity management (doubling strategy)
' - Indexed access with bounds checking
' - Add, InsertAt, RemoveAt methods for flexible array operations
' ---------------------------------------------------------------------------
' Team: Bedrock Team
' Last Update: 17/05/2025
' ***************************************************************************

Option Explicit

Private Type TList
    Items() As Variant
    Count As Long
    Capacity As Long
End Type

Private Const DEFAULT_CAP As Long = 8
Private data As TList

Private Sub Class_Initialize()
    data.Capacity = DEFAULT_CAP
    data.Count = 0
    ReDim data.Items(0 To data.Capacity - 1)
End Sub

Private Sub EnsureCapacity(Optional ByVal additional As Long = 1)
    Dim required As Long: required = data.Count + additional
    If required > data.Capacity Then
        If data.Capacity < 1024 Then
            data.Capacity = data.Capacity * 2
        Else
            data.Capacity = CLng(data.Capacity * 1.5)
        End If
        If data.Capacity < required Then data.Capacity = required
        ReDim Preserve data.Items(0 To data.Capacity - 1)
    End If
End Sub

Public Property Get Length() As Long: Length = data.Count: End Property
Public Property Get Capacity() As Long: Capacity = data.Capacity: End Property

Public Property Get Item(Index As Long) As Variant
    If Index < 0 Or Index >= data.Count Then Err.Raise vbObjectError + 1, "List: Bedrock", "Index out of bounds"
    Item = data.Items(Index)
End Property

Public Property Let Item(Index As Long, ByVal Value As Variant)
Attribute Item.VB_UserMemId = 0
    If Index < 0 Or Index >= data.Count Then Err.Raise vbObjectError + 1, "List: Bedrock", "Index out of bounds"
    data.Items(Index) = Value
End Property

Public Property Set Item(Index As Long, ByVal Value As Object)
    If Index < 0 Or Index >= data.Count Then Err.Raise vbObjectError + 1, "List: Bedrock", "Index out of bounds"
    Set data.Items(Index) = Value
End Property

Public Property Get Items() As Variant
    Dim Arr() As Variant
    If data.Count > 0 Then
        ReDim Arr(0 To data.Count - 1)
        Dim i As Long
        For i = 0 To data.Count - 1
            Arr(i) = data.Items(i)
        Next i
    Else
        ReDim Arr(-1 To -1)
    End If
    Items = Arr
End Property

Public Property Let Items(ByVal newItems As Variant)
    If Not IsArray(newItems) Then Err.Raise vbObjectError + 2, "List: Bedrock", "Assigned value must be an array"
    Dim newCount As Long: newCount = UBound(newItems) - LBound(newItems) + 1
    
    data.Count = newCount
    data.Capacity = IIf(newCount < DEFAULT_CAP, DEFAULT_CAP, newCount)
    
    ReDim data.Items(0 To data.Capacity - 1)
    Dim i As Long, src As Long: src = LBound(newItems)
    
    For i = 0 To newCount - 1
        If IsObject(newItems(src + i)) Then Set data.Items(i) = newItems(src + i) Else data.Items(i) = newItems(src + i)
    Next i
End Property

Public Sub Add(ByVal Value As Variant)
    EnsureCapacity
    
    If IsObject(Value) Then
        Set data.Items(data.Count) = Value
    Else
        data.Items(data.Count) = Value
    End If
    data.Count = data.Count + 1
End Sub

Public Sub AddRange(ByVal Arr As Variant)
    If Not IsArray(Arr) Then Err.Raise vbObjectError + 3, "List: Bedrock", "AddRange requires an array"
    
    Dim cntNew As Long: cntNew = UBound(Arr) - LBound(Arr) + 1
    EnsureCapacity cntNew
    
    Dim base As Long: base = data.Count
    Dim i As Long
    
    For i = 0 To cntNew - 1: data.Items(base + i) = Arr(LBound(Arr) + i): Next i
    data.Count = data.Count + cntNew
End Sub

Public Sub InsertAt(Index As Long, ByVal Value As Variant)
    If Index < 0 Or Index > data.Count Then Err.Raise vbObjectError + 1, "List: Bedrock", "Index out of bounds"
    EnsureCapacity
    
    Dim i As Long
    
    For i = data.Count To Index + 1 Step -1
        data.Items(i) = data.Items(i - 1)
    Next i
    
    data.Items(Index) = Value
    data.Count = data.Count + 1
End Sub

Public Sub RemoveAt(Index As Long)
    If Index < 0 Or Index >= data.Count Then Err.Raise vbObjectError + 1, "List: Bedrock", "Index out of bounds"
    Dim i As Long
    
    For i = Index To data.Count - 2
        data.Items(i) = data.Items(i + 1)
    Next i
    
    data.Count = data.Count - 1
    data.Items(data.Count) = Empty
End Sub

Public Sub Clear()
    data.Capacity = DEFAULT_CAP
    data.Count = 0
    ReDim data.Items(0 To data.Capacity - 1)
End Sub

Public Sub Trim()
    If data.Count > 0 Then
        data.Capacity = data.Count
        ReDim Preserve data.Items(0 To data.Capacity - 1)
    Else
        Clear
    End If
End Sub

Public Sub Shuffle()
    Dim i As Long
    Dim j As Long
    Dim temp As Variant
    
    Randomize
    
    For i = data.Count - 1 To 1 Step -1
        j = Int((i + 1) * Rnd)
        temp = data.Items(i)
        data.Items(i) = data.Items(j)
        data.Items(j) = temp
    Next i
End Sub

Public Function Join(delimeter As String) As String
    Join = VBA.Join(Me.ToArray, delimeter)
End Function

Public Sub Reverse()
    Dim i As Long, tmp As Variant, last As Long: last = data.Count - 1
    For i = 0 To (data.Count \ 2) - 1
        tmp = data.Items(i)
        data.Items(i) = data.Items(last - i)
        data.Items(last - i) = tmp
    Next i
End Sub

Public Function ToArray() As Variant
    Dim Arr() As Variant
    If data.Count > 0 Then
        ReDim Arr(0 To data.Count - 1)
        Dim i As Long
        For i = 0 To data.Count - 1
            Arr(i) = data.Items(i)
        Next i
    Else
        ReDim Arr(-1 To -1)
    End If
    ToArray = Arr
End Function

Public Function IndexOf(ByVal Value As Variant) As Long
    Dim i As Long
    For i = 0 To data.Count - 1
        If IsObject(data.Items(i)) And IsObject(Value) Then
            If data.Items(i) Is Value Then
                IndexOf = i
                Exit Function
            End If
        Else
            If data.Items(i) = Value Then
                IndexOf = i
                Exit Function
            End If
        End If
    Next i
    IndexOf = -1
End Function

Public Function Contains(ByVal Value As Variant) As Boolean
    Contains = (IndexOf(Value) <> -1)
End Function

Public Function LastIndexOf(ByVal Value As Variant) As Long
    Dim i As Long
    For i = data.Count - 1 To 0 Step -1
        If IsObject(data.Items(i)) And IsObject(Value) Then
            If data.Items(i) Is Value Then
                LastIndexOf = i
                Exit Function
            End If
        Else
            If data.Items(i) = Value Then
                LastIndexOf = i
                Exit Function
            End If
        End If
    Next i
    LastIndexOf = -1
End Function

Public Function Concat(ByVal Arr As List) As List
    Dim result As New List
    result.AddRange Me.ToArray
    
    If TypeOf other Is List Then
        result.AddRange Arr.ToArray
    Else
        Err.Raise vbObjectError + 6, "List: Bedrock", "Concat accepts List or VBA array"
    End If
    
    Set Concat = result
End Function

Public Sub Sort(Optional ByVal ascending As Boolean = True)
    If data.Count < 2 Then Exit Sub
    QuickSort 0, data.Count - 1, ascending
End Sub

Private Sub QuickSort(ByVal low As Long, ByVal high As Long, ByVal asc As Boolean)
    Dim i As Long, j As Long, pivot As Variant, tmp As Variant
    i = low: j = high: pivot = data.Items((low + high) \ 2)
    Do While i <= j
        If asc Then
            Do While data.Items(i) < pivot: i = i + 1: Loop
            Do While data.Items(j) > pivot: j = j - 1: Loop
        Else
            Do While data.Items(i) > pivot: i = i + 1: Loop
            Do While data.Items(j) < pivot: j = j - 1: Loop
        End If
        If i <= j Then
            tmp = data.Items(i): data.Items(i) = data.Items(j): data.Items(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If low < j Then QuickSort low, j, asc
    If i < high Then QuickSort i, high, asc
End Sub

Public Function Clone() As List
    Dim result As New List
    result.AddRange Me.ToArray
    
    Set Clone = result
End Function
