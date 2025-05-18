# List Documentation

A dynamic, array-backed collection class for VBA, delivering flexible list operations (add, insert, remove, sort, shuffle, and more) with automatic capacity management and no external dependencies.

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
* [Usage Examples](#usage-examples)
* [Implementation Details](#implementation-details)
* [Performance and Scalability](#performance-and-scalability)

## Features

* **Pure VBA**: No DLL imports or external libraries.
* **Automatic Resizing**: Grows capacity (×2 up to a threshold, then ×1.5) to minimize reallocations.
* **Rich API**: Add, Insert, Remove, Clear, Trim, Shuffle, Reverse, Sort, Concat, Clone, and more.
* **Variant Storage**: Holds any data type or object reference, handling `Set` for objects.
* **Snapshot Arrays**: `Items` and `ToArray` return a zero-based VBA array of current elements.
* **QuickSort**: In-place, configurable ascending/descending sort.

## Requirements

* Microsoft Office VBA (Excel, Word, Access, etc.)
* VBA project with `Option Explicit` enabled.

## Installation

1. Open the VBA editor (**Alt + F11**).
2. In the **Project Explorer**, right-click your target project and choose **Import File…**.
3. Select `List.cls` and confirm.
4. Save your project.

## API Reference

### Properties

| Name                                     | Type             | Description                                                |
| ---------------------------------------- | ---------------- | ---------------------------------------------------------- |
| `Length`                                 | Long             | Number of elements currently in the list.                  |
| `Capacity`                               | Long             | Allocated storage slots (≥ Length).                        |
| `Item(Index)`<br/>(Property Get/Let/Set) | Variant / Object | Get or set element at zero-based `Index` (bounds-checked). |
| `Items`<br/>(Property Get/Let)           | Variant Array    | Snapshot of all elements as a zero-based VBA array.        |

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

#### Clear()

Empties the list and resets capacity to default (8).

```vb
lst.Clear
```

#### Trim()

Reduces capacity to match current length, reclaiming unused space.

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

Returns first index of `Value`, or –1 if not found (uses reference equality for objects).

```vb
Dim idx As Long
idx = lst.IndexOf "Hello"
```

#### LastIndexOf(Value As Variant) As Long

Returns last index of `Value`, or –1 if not found.

```vb
Dim lastIdx As Long
lastIdx = lst.LastIndexOf 42
```

#### Contains(Value As Variant) As Boolean

Returns `True` if `Value` exists in the list.

```vb
If lst.Contains("Hello") Then …
```

#### Concat(other As List) As List

Returns a new `List` combining current elements followed by all elements of `other`.

```vb
Dim lst2 As New List
lst2.AddRange Array("A", "B")
Dim merged As List
Set merged = lst.Concat(lst2)
```

#### Sort(Optional ascending As Boolean = True)

Sorts elements in-place using QuickSort. `ascending = False` for descending order.

```vb
lst.Sort False  ' descending
```

#### Clone() As List

Creates a shallow copy of the list and its elements.

```vb
Dim copy As List
Set copy = lst.Clone
```

## Usage Examples

> [!TIP]
> You can do `list(<index>)` instead of `list.Item(<index>)`.

```vb
Sub Example_ListOperations()
    Dim fruits As New List
    fruits.AddRange Array("Apple", "Banana", "Cherry")
    fruits.InsertAt 1, "Blueberry"
    Debug.Print "Count: " & fruits.Length
    Debug.Print "Second item: " & fruits.Item(1)
    
    fruits.RemoveAt 2
    fruits.Sort
    Debug.Print fruits.Join(" | ")
    
    fruits.Shuffle
    Debug.Print "Shuffled: " & fruits.Join(", ")
End Sub
```

```vb
Sub Example_Advanced()
    Dim numbers As New List
    numbers.AddRange Array(5, 3, 8, 1, 9)
    numbers.Sort False   ' descending
    Debug.Print "Max: " & numbers.Item(0)
    
    Dim copy As List
    Set copy = numbers.Clone
    copy.Clear
    Debug.Print "Original still has " & numbers.Length & " items"
End Sub
```

## Implementation Details

```vb
Private Type TList
    Items()   As Variant
    Count     As Long
    Capacity  As Long
End Type

Const DEFAULT_CAP As Long = 8
Private data As TList
```

* **EnsureCapacity** doubles capacity (or ×1.5 above threshold) when needed.
* **QuickSort** recursively sorts between low/high indices.
* **Variant vs. Object** storage handled via `IsObject` and appropriate `Set`.

## Performance and Scalability

* **Add**: Amortized O(1)
* **Insert/Remove**: O(n) (shifts elements)
* **Lookup**: O(1) by index
* **Sort**: Average O(n log n), worst O(n²) in degenerate pivot cases
* **Memory**: Grows in powers of two until threshold, then 1.5×, balancing allocations and usage

For questions or contributions, please contact **Team Bedrock** or open an issue in your project repository.
