# Debouncer Documentation

A high-performance debounce and interval timer system implemented in pure VBA. Eliminates external dependencies (e.g., `Scripting.Dictionary`, Win32 APIs) by leveraging a typed array, binary search (O(log n)), and the native `Timer` function.

## Download

- [Latest Version](/modules/Debouncer/Debouncer.bas)
- [Previous Versions](/modules/Debouncer/versions)

## Table of Contents

- [Features](#features)  
- [Requirements](#requirements)  
- [Installation](#installation)  
- [API Reference](#api-reference)  
  - [`Wait(Seconds, Key, [Once]) As Boolean`](#waitseconds-key-once-as-boolean)  
  - [`ClearInterval(Key)`](#clearintervalkey)  
  - [`ClearIntervals()`](#clearintervals)  
- [Usage Examples](#usage-examples)  
- [Implementation Details](#implementation-details)  
- [Performance and Scalability](#performance-and-scalability)

## Features

- **Pure VBA**: No DLL imports or COM objects.  
- **Array-Based Storage**: Timers stored in a dynamic, sorted array keyed by string.  
- **Binary Search**: O(log n) lookup for timer insertions/queries.  
- **Automatic Resizing**: `ReDim Preserve` in powers of two to minimize reallocations.  
- **One-Shot & Recurring**: Support for single-trigger (`Once = True`) and repeating intervals.  
- **Lightweight Timing**: Uses the built-in `Timer` function for millisecond-precision.

---

## Requirements

- Microsoft Office VBA (Excel, Word, Access, etc.)  
- Optionally: VBIDE access to import/insert modules

---

## Installation

1. Open the VBA editor (e.g., **Alt + F11**).  
2. In the **Project Explorer**, right-click your target project and choose **Import File…**.  
3. Select `Debouncer.bas` and confirm.  
4. Save your project.

---

## API Reference

### Wait(Seconds, Key, [Once]) As Boolean

Registers or checks a timer keyed by `Key`. Returns `True` once the specified interval has elapsed.

| Parameter | Type    | Description                                                                 |
|-----------|---------|-----------------------------------------------------------------------------|
| Seconds   | Double  | Interval duration in seconds (can be fractional).                           |
| Key       | String  | Unique identifier for this timer.                                           |
| Once      | Boolean | Optional. If `True`, timer fires only once and then becomes inactive. Defaults to `False`. |

**Usage**  
- On first call with a new `Key`, initializes the timer and returns `False`.  
- Subsequent calls return `False` until `Seconds` have elapsed.  
- Returns `True` once the duration has elapsed.  
- If `Once = True`, further calls with the same `Key` will return `False` (timer is “done”).  
- If `Once = False`, timer resets its start time each time it fires.

> [!TIP]
> Use the `Wait` function inside a loop.  
> This function checks if a given amount of time has passed **since the last call** using a unique `Key`.  
> It will keep returning `False` until the interval completes, and only then return `True`—perfect for debounce or periodic tasks.  
>  
> Example usage:
>
> ```vba
> Do
>     If Wait(1.5, "my_timer") Then
>         Debug.Print "Executed every 1.5 seconds"
>     End If
>     DoEvents
> Loop
> ```
>
> Without the loop, `Wait` would only check once and might never return `True`.

### ClearInterval(Key)

Removes the timer entry identified by `Key`. Subsequent calls to `Wait` with the same `Key` will re-register it from scratch.

```vb
Debouncer.ClearInterval "MyTimer"
```

### ClearIntervals()

Resets the entire timer system: clears all timers and deallocates internal storage.

```vb
Debouncer.ClearIntervals
```

## Usage Examples

```vb
Sub Example_Debounce()
    Do
        If Debouncer.Wait(1#, "ChangeHandler") Then
            Call ProcessLargeDataset
        End If
        DoEvents
    Loop
End Sub
````

### Repeated Interval Logging

```vb
Sub Example_IntervalLogging()
    ' Writes a timestamp to the Immediate window every 5 seconds
    Do
        If Debouncer.Wait(5#, "Logger") Then
            Debug.Print "Tick: " & Format(Now, "HH:NN:SS")
        End If
        DoEvents
    Loop
End Sub
```

### One-Shot Cleanup Scheduling

```vb
Sub Example_OneShotCleanup()
    ' Triggers cleanup only once, 30 seconds after the first call
    Do
        If Debouncer.Wait(30#, "CleanupTask", True) Then
            Call PerformCleanup
        End If
        DoEvents
    Loop
End Sub
```

> \[!TIP]
> Always run `Debouncer.Wait` inside a loop.
> It continuously checks the time condition and returns `True` only when the target interval is met.
> For `Once=True`, it triggers just a single time and never again unless cleared or reset.

## Implementation Details

* **Data Structure**

  ```vb
  Private Type IntervalTimer
      Key       As String
      StartTime As Double
      Once      As Boolean
      Done      As Boolean
  End Type
  ```
* **Lookup**: `BinarySearch(Key, found)` returns insertion index or existing position.
* **Resizing**: On overflow, array doubles its size (`1 → 2 → 4 → 8 → …`).
* **Timer Source**: Uses VBA’s `Timer` (seconds since midnight with fractional part).

## Performance and Scalability

* **Lookup Complexity:** O(log n) via binary search on a sorted array.
* **Insertion/Deletion:** O(n) for shifting elements, amortized over infrequent resizes.
* **Memory:** Dynamic array grows in powers of two, minimizing memory churn.
* **Precision:** Dependent on VBA’s `Timer`, typically \~10 ms resolution on Windows.

> For questions or contributions, please contact **Team Bedrock** or open an issue in your project repository.
