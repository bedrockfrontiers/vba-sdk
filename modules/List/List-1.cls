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
' Last Update: 18/05/2025
' ***************************************************************************

Option Explicit

Private Type TList
    Items() As Variant
    Count As Long
    Capacity As Long
End Type

Private Const DEFAULT_CAP As Long = 8
Private Const TOLERANCE As Long = 16
Private data As TList

Private Sub Class_Initialize()
    data.Capacity = DEFAULT_CAP
    data.Count = 0
    ReDim data.Items(0 To data.Capacity - 1)
End Sub

Private Sub EnsureCapacity(Optional ByVal additional As Long = 1)
    Dim required As Long: required = data.Count + additional
    If required <= data.Capacity Then Exit Sub

    If data.Capacity < 1024 Then
        data.Capacity = data.Capacity * 2
    Else
        data.Capacity = CLng(data.Capacity * 1.5)
    End If
    
    If data.Capacity < required Then data.Capacity = required
    ReDim Preserve data.Items(0 To data.Capacity - 1)
End Sub

Public Property Get Length() As Long: Length = data.Count: End Property
Public Property Get Capacity() As Long: Capacity = data.Capacity: End Property

Public Function At(ByVal Index As Long) As Variant
    If Index < 0 Or Index >= data.Count Then Err.Raise vbObjectError + 1, "List: Bedrock", "Index out of bounds"
    If IsObject(data.Items(Index)) Then
        Set At = data.Items(Index)
    Else
        At = data.Items(Index)
    End If
End Function

Public Property Get Item(Index As Long) As Variant
    If Index < 0 Or Index >= data.Count Then Err.Raise vbObjectError + 1, "List: Bedrock", "Index out of bounds"
    
    If IsObject(data.Items(Index)) Then
        Set Item = data.Items(Index)
    Else
        Item = data.Items(Index)
    End If
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
    If data.Count + TOLERANCE > data.Capacity Then
        EnsureCapacity
    End If
    
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
    
    If IsObject(Value) Then
        Set data.Items(Index) = Value
    Else
        data.Items(Index) = Value
    End If
    data.Count = data.Count + 1
End Sub

Public Sub ForEach(callback As Object)
    Dim i As Long
    For i = 0 To data.Count - 1
        CallByName callback, "Invoke", VbMethod, data.Items(i)
    Next i
End Sub

Public Function First() As Variant
    If data.Count = 0 Then Err.Raise vbObjectError + 1, "List: Bedrock", "List is empty"
    
    If IsObject(data.Items(0)) Then
        Set First = data.Items(0)
    Else
        First = data.Items(0)
    End If
End Function

Public Function Last() As Variant
    If data.Count = 0 Then Err.Raise vbObjectError + 1, "List: Bedrock", "List is empty"
    
    If IsObject(data.Items(data.Count - 1)) Then
        Set Last = data.Items(data.Count - 1)
    Else
        Last = data.Items(data.Count - 1)
    End If
End Function

Public Sub RemoveAt(Index As Long)
    If Index < 0 Or Index >= data.Count Then Err.Raise vbObjectError + 1, "List: Bedrock", "Index out of bounds"
    Dim i As Long
    
    For i = Index To data.Count - 2
        data.Items(i) = data.Items(i + 1)
    Next i
    
    data.Count = data.Count - 1
    data.Items(data.Count) = Empty
End Sub

Public Function Remove(ByVal Value As Variant) As Boolean
    Dim idx As Long
    idx = IndexOf(Value)
    If idx <> -1 Then
        RemoveAt idx
        Remove = True
    Else
        Remove = False
    End If
End Function

Public Function Pop() As Variant
    If data.Count = 0 Then Err.Raise vbObjectError + 1, "List: Bedrock", "List is empty"
    Pop = data.Items(data.Count - 1)
    data.Count = data.Count - 1
End Function

Public Function Shift() As Variant
    If data.Count = 0 Then Err.Raise vbObjectError + 1, "List: Bedrock", "List is empty"
    Shift = data.Items(0)
    RemoveAt 0
End Function

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
    Dim i As Long, j As Long, tmp As Variant
    Dim isObjI As Boolean, isObjJ As Boolean
    Randomize
    For i = data.Count - 1 To 1 Step -1
        j = Int((i + 1) * Rnd)
        isObjI = IsObject(data.Items(i)): isObjJ = IsObject(data.Items(j))
        If isObjI Then Set tmp = data.Items(i) Else tmp = data.Items(i)
        If isObjJ Then Set data.Items(i) = data.Items(j) Else data.Items(i) = data.Items(j)
        If isObjI Then Set data.Items(j) = tmp Else data.Items(j) = tmp
    Next i
End Sub

Public Function Join(delimeter As String) As String
    Join = VBA.Join(Me.ToArray, delimeter)
End Function

Public Sub Reverse()
    Dim i As Long, j As Long
    j = data.Count - 1
    For i = 0 To (data.Count \ 2) - 1
        Swap i, j
        j = j - 1
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

Public Function BinarySearch(Value As Variant, Optional ascending As Boolean = True) As Long
    Dim low As Long, high As Long, mid As Long
    Dim midVal As Variant
    low = 0: high = data.Count - 1
    
    Do While low <= high
        mid = (low + high) \ 2
        midVal = data.Items(mid)
        If (IsObject(midVal) And IsObject(Value) And midVal Is Value) _
           Or (Not IsObject(midVal) And Not IsObject(Value) And midVal = Value) Then
            BinarySearch = mid: Exit Function
        End If
        If IsObject(midVal) Or IsObject(Value) Then Exit Do
        If (midVal < Value) Xor (Not ascending) Then
            low = mid + 1
        Else
            high = mid - 1
        End If
    Loop
    
    BinarySearch = -1
End Function


Public Function Every(callback As Object) As Boolean
    Dim i As Long
    For i = 0 To data.Count - 1
        If Not CallByName(callback, "Invoke", VbMethod, data.Items(i)) Then
            Every = False: Exit Function
        End If
    Next i
    Every = True
End Function

Public Function Some(callback As Object) As Boolean
    Dim i As Long
    For i = 0 To data.Count - 1
        If CallByName(callback, "Invoke", VbMethod, data.Items(i)) Then
            Some = True: Exit Function
        End If
    Next i
    Some = False
End Function

Public Function Map(callback As Object) As List
    Dim result As New List, i As Long, mapped As Variant
    For i = 0 To data.Count - 1
        mapped = CallByName(callback, "Invoke", VbMethod, data.Items(i))
        result.Add mapped
    Next i
    Set Map = result
End Function

Public Function Filter(callback As Object) As List
    Dim result As New List, i As Long, keep As Boolean
    For i = 0 To data.Count - 1
        keep = CallByName(callback, "Invoke", VbMethod, data.Items(i))
        If keep Then result.Add data.Items(i)
    Next i
    Set Filter = result
End Function

Public Function Reduce(callback As Object, initial As Variant) As Variant
    Dim acc As Variant: acc = initial
    Dim i As Long
    For i = 0 To data.Count - 1
        acc = CallByName(callback, "Invoke", VbMethod, acc, data.Items(i))
    Next i
    Reduce = acc
End Function

Public Function FindAll(callback As Object) As List
    Set FindAll = Filter(callback)
End Function

Public Function Slice(startIndex As Long, endIndex As Long) As List
    Dim result As New List, i As Long
    If startIndex < 0 Then startIndex = 0
    If endIndex > data.Count Then endIndex = data.Count
    For i = startIndex To endIndex - 1
        result.Add data.Items(i)
    Next i
    Set Slice = result
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

Private Sub Swap(i As Long, j As Long)
    Dim tmp As Variant
    Dim isObjI As Boolean: isObjI = IsObject(data.Items(i))
    Dim isObjJ As Boolean: isObjJ = IsObject(data.Items(j))
    
    If isObjI Then
        Set tmp = data.Items(i)
    Else
        tmp = data.Items(i)
    End If
    
    If isObjJ Then
        Set data.Items(i) = data.Items(j)
    Else
        data.Items(i) = data.Items(j)
    End If

    If isObjI Then
        Set data.Items(j) = tmp
    Else
        data.Items(j) = tmp
    End If
End Sub

Private Sub InsertionSort(low As Long, high As Long, asc As Boolean)
    Dim i As Long, j As Long, key As Variant
    For i = low + 1 To high
        key = data.Items(i)
        j = i - 1
        If asc Then
            Do While j >= low And data.Items(j) > key
                data.Items(j + 1) = data.Items(j): j = j - 1
            Loop
        Else
            Do While j >= low And data.Items(j) < key
                data.Items(j + 1) = data.Items(j): j = j - 1
            Loop
        End If
        data.Items(j + 1) = key
    Next i
End Sub

Public Sub Sort(Optional ByVal ascending As Boolean = True)
    If data.Count < 2 Then Exit Sub
    QuickSort 0, data.Count - 1, ascending, 0
End Sub

Private Sub QuickSort(low As Long, high As Long, asc As Boolean, depth As Long)
    Const MAX_DEPTH As Long = 64
    If high - low < 16 Then
        InsertionSort low, high, asc
        Exit Sub
    End If
    If depth > MAX_DEPTH Then
        InsertionSort low, high, asc
        Exit Sub
    End If
    Dim i As Long, j As Long, pivot As Variant
    i = low: j = high: pivot = data.Items((low + high) \ 2)
    Do While i <= j
        If asc Then
            Do While data.Items(i) < pivot: i = i + 1: Loop
            Do While data.Items(j) > pivot: j = j - 1: Loop
        Else
            Do While data.Items(i) > pivot: i = i + 1: Loop
            Do While data.Items(j) < pivot: j = j - 1: Loop
        End If
        If i <= j Then Swap i, j: i = i + 1: j = j - 1
    Loop
    If low < j Then QuickSort low, j, asc, depth + 1
    If i < high Then QuickSort i, high, asc, depth + 1
End Sub

Public Function Clone() As List
    Dim result As New List
    result.AddRange Me.ToArray
    
    Set Clone = result
End Function
