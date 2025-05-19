# List Module Documentation

A high-performance, dynamic array–backed collection class for VBA, delivering flexible list operations (Add, InsertAt, RemoveAt, Sort, Shuffle, Map, Filter, Reduce, and more) with automatic capacity management and zero external dependencies.

## Download

* [Latest Version](/modules/List/List.cls)
* [Previous Versions](/modules/List/versions)

## Table of Contents

* [Features](#features)
* [Requirements](#requirements)
* [Installation](#installation)
* [API Reference](#api-reference)

  * [Properties](#properties)
  * [Methods](#methods)
  * [Callback Classes](#callback-classes)
* [Usage Examples](#usage-examples)
* [Implementation Details](#implementation-details)
* [Performance and Scalability](#performance-and-scalability)

## Features

* **Pure VBA**: No external libraries or DLL imports.
* **Automatic Resizing**: Capacity doubles (×2) up to 1 024 elements, then grows by 1.5× to minimize reallocations.
* **Rich API**: Add, AddRange, InsertAt, RemoveAt, Remove, Pop, Shift, Clear, Trim, Shuffle, Reverse, Join, Concat, Clone.
* **Functional & Callback Methods**: ForEach, Map, Filter, Reduce, Every, Some, FindAll, Slice.
* **Variant & Object Storage**: Holds any data type or object reference, handling `Set` where appropriate.
* **Indexed Access**: `Item` and `At` with bounds checking.
* **Binary Search**: Fast lookup in sorted lists via `BinarySearch`.

## Requirements

* Microsoft Office VBA host (Excel, Word, Access, etc.)
* VBA project with `Option Explicit` enabled

## Installation

1. Open the VBA editor (**Alt + F11**).
2. In the **Project Explorer**, right-click your target project and choose **Import File…**.
3. Select **`List.cls`** and confirm.
4. Save your VBA project.

## API Reference

### Properties

| Name          | Type             | Description                                                   |
| ------------- | ---------------- | ------------------------------------------------------------- |
| `Length`      | Long             | Number of elements currently in the list.                     |
| `Capacity`    | Long             | Allocated storage slots (always ≥ `Length`).                  |
| `At(Index)`   | Variant / Object | Returns element at zero-based `Index` (with bounds checking). |
| `Item(Index)` | Variant / Object | Get or set element at zero-based `Index` (bounds-checked).    |
| `Items`       | Variant Array    | Snapshot of all elements as a zero-based VBA array (get/set). |

### Methods

#### Add(Value As Variant)

Appends `Value` to the end of the list, resizing if needed.

```vb
Dim lst As New List
lst.Add 42
lst.Add "Hello"
```

#### AddRange(Arr As Variant)

Appends all elements from a VBA array `Arr`. Raises error if `Arr` is not an array.

```vb
lst.AddRange Array(1, 2, 3)
```

#### InsertAt(Index As Long, Value As Variant)

Inserts `Value` at position `Index` (0 ≤ Index ≤ Length), shifting subsequent items right.

```vb
lst.InsertAt 1, "Middle"
```

#### RemoveAt(Index As Long)

Removes the element at `Index`, shifting subsequent items left.

```vb
lst.RemoveAt 0  ' removes first element
```

#### Remove(Value As Variant) As Boolean

Removes the first occurrence of `Value`; returns `True` if removed.

```vb
If lst.Remove("Banana") Then Debug.Print "Removed!"
```

#### Pop() As Variant

Removes and returns the last element. Error if list is empty.

```vb
Dim lastItem As Variant
lastItem = lst.Pop
```

#### Shift() As Variant

Removes and returns the first element. Error if list is empty.

```vb
Dim firstItem As Variant
firstItem = lst.Shift
```

#### Clear()

Empties the list and resets capacity to default (8).

```vb
lst.Clear
```

#### Trim()

Reduces capacity to match current `Length`, reclaiming unused space.

```vb
lst.Trim
```

#### Shuffle()

Randomly permutes list elements (Fisher–Yates shuffle).

```vb
lst.Shuffle
```

#### Reverse()

Reverses the order of elements in-place.

```vb
lst.Reverse
```

#### Join(delimiter As String) As String

Concatenates all elements into a single string, separated by `delimiter`.

```vb
Debug.Print lst.Join(", ")
```

#### ToArray() As Variant

Returns a zero-based VBA array containing all current elements.

```vb
Dim arr As Variant
arr = lst.ToArray()
```

#### IndexOf(Value As Variant) As Long

Returns the first index of `Value`, or –1 if not found (uses reference equality for objects).

```vb
Dim idx As Long
idx = lst.IndexOf("Hello")
```

#### LastIndexOf(Value As Variant) As Long

Returns the last index of `Value`, or –1 if not found.

```vb
Dim lastIdx As Long
lastIdx = lst.LastIndexOf(42)
```

#### Contains(Value As Variant) As Boolean

Returns `True` if `Value` exists in the list.

```vb
If lst.Contains("Hello") Then …
```

#### BinarySearch(Value As Variant, Optional ascending As Boolean = True) As Long

Performs binary search on a sorted list; returns index or –1 if not found.

```vb
Dim pos As Long
pos = lst.BinarySearch(5, True)
```

#### Concat(other As List) As List

Returns a new `List` combining current elements followed by those of `other`.

```vb
Dim merged As List
Set merged = lst.Concat(anotherList)
```

#### Clone() As List

Creates a shallow copy of the list and its elements.

```vb
Dim copy As List
Set copy = lst.Clone
```

### Callback Classes

The following methods accept a **callback object**—a class instance exposing an `Invoke` method. Use these to encapsulate logic and promote reuse.

| Method    | Signature                                                     | Description                                       |
| --------- | ------------------------------------------------------------- | ------------------------------------------------- |
| `ForEach` | `Sub Invoke(item As Variant)`                                 | Executes once per element (no return value).      |
| `Map`     | `Function Invoke(item As Variant) As Variant`                 | Transforms each item; returns new list of values. |
| `Filter`  | `Function Invoke(item As Variant) As Boolean`                 | Includes item if return is `True`.                |
| `Reduce`  | `Function Invoke(acc As Variant, item As Variant) As Variant` | Aggregates values into a single result.           |
| `Every`   | `Function Invoke(item As Variant) As Boolean`                 | Returns `True` only if all calls return `True`.   |
| `Some`    | `Function Invoke(item As Variant) As Boolean`                 | Returns `True` if any call returns `True`.        |
| `FindAll` | Alias for `Filter`                                            | —                                                 |
| `Slice`   | *No callback; returns sublist.*                               | —                                                 |

#### How to Implement a Callback Class

1. **Add a new Class Module**: e.g. `clsPrinter`, `clsSquarer`, etc.
2. **Define the `Invoke` method** matching the required signature.
3. **Instantiate** and pass to list method.

```vb
'-- clsPrinter.cls --
Option Explicit

Public Sub Invoke(item As Variant)
    Debug.Print item
End Sub
```

```vb
'-- clsSquarer.cls --
Option Explicit

Public Function Invoke(item As Variant) As Variant
    Invoke = item * item
End Function
```

#### Tips & Best Practices

* **Single Responsibility**: Each callback class should do one job.
* **Reusable**: Define generic callbacks (e.g. `clsEqualityChecker`) for common tasks.
* **Naming**: Prefix classes with `cls` and methods clearly (`Invoke`).
* **Error Handling**: Validate input inside `Invoke` and raise meaningful errors.

## Usage Examples

```vb
Sub Example_Functional()
    Dim numbers As New List
    numbers.AddRange Array(2, 3, 4, 5)

    ' Print each element
    numbers.ForEach New clsPrinter

    ' Square each element
    Dim squares As List
    Set squares = numbers.Map(New clsSquarer)
    Debug.Print "Squares: " & squares.Join(", ")

    ' Filter even numbers
    Dim evens As List
    Set evens = numbers.Filter(New clsEvenChecker)
    Debug.Print "Evens: " & evens.Join(", ")

    ' Sum all numbers
    Dim total As Variant
    total = numbers.Reduce(New clsAdder, 0)
    Debug.Print "Sum: " & total
End Sub
```

## Implementation Details

```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  ' True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Type TList
    Items()   As Variant
    Count     As Long
    Capacity  As Long
End Type

Private Const DEFAULT_CAP As Long = 8
Private Const TOLERANCE    As Long = 16
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
```

## Performance and Scalability

* **Add**: Amortized O(1)
* **InsertAt / RemoveAt**: O(n)
* **Index Access**: O(1)
* **Sort**: Average O(n log n), worst O(n²)
* **BinarySearch**: O(log n)

This `List` class provides a robust foundation for dynamic collections in VBA, balancing speed, memory efficiency, and a rich feature set.
