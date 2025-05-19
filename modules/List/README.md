# List (v1.2)

[Documentation](/docs/List)

## Overview

The **List** class module provides a dynamic, array-backed collection type for VBA, delivering familiar list operations such as adding, inserting, removing, and querying items. It abstracts away manual array management, delivering efficient resizing, indexing, and high-level utilities—ideal for scenarios where a native VBA array falls short in flexibility and performance.

## Purpose and Problem Addressed

Native VBA arrays require manual `ReDim Preserve` calls and careful index tracking, which can be error-prone and tedious when you need:

* Dynamic growth or shrinkage without pre-knowing final size.
* Seamless insertion and removal at arbitrary positions.
* Utility methods for sorting, searching, and shuffling.
* Encapsulation of boundary checks and capacity logic.

The **List** module addresses these pain points by wrapping array behavior in a robust class, offering an intuitive API that automatically handles capacity, bounds validation, and common list algorithms.

## Core Technical Concepts

### Mechanism

At its core, the List uses a private dynamic array (`Items() As Variant`) along with two counters—`Count` (number of stored elements) and `Capacity` (allocated slots). All items live in `Items(0 … Capacity–1)`, but only indices `0 … Count–1` are considered valid.

### Implementation Details

* **Automatic Resizing (`EnsureCapacity`)**
  When adding or inserting elements, the module checks if `Count + needed > Capacity`. It then doubles the capacity (up to a threshold) or grows by 1.5× for large lists, minimizing reallocations while bounding memory overhead.

* **Index Validation**
  Each getter, setter, and mutation method validates the target index, raising a descriptive VBA error if out of bounds, preventing silent data corruption.

* **Variant Storage**
  Storing items as `Variant` allows for any data type—scalars or object references. The code distinguishes via `IsObject` to use `Set` appropriately.

* **Rich API Surface**
  Methods include:

  * `Add`, `AddRange`
  * `InsertAt`, `RemoveAt`
  * `Clear`, `Trim` (shrink to fit)
  * `Shuffle`, `Reverse`
  * `Join` (concatenate into a delimited string)
  * `ToArray`, `Items` property (snapshot)
  * `IndexOf`, `LastIndexOf`, `Contains`
  * `Concat` (merge two lists)
  * `Sort` (in-place QuickSort)
  * `Clone` (deep copy)

* **QuickSort Implementation**
  Provides an in-place, recursive QuickSort with configurable ascending/descending order, achieving average O(n log n) sort performance.

## Advantages Over Conventional Methods

* **Simplicity and Safety**
  Eliminates repetitive `ReDim Preserve` boilerplate and manual bounds checks.

* **Performance**
  Amortized O(1) `Add` operations and efficient resizing strategy avoid frequent memory churn.

* **Feature-Rich**
  Offers built-in algorithms (sort, shuffle, reverse) and helpers (ToArray, Join), reducing custom code.

* **Portability**
  Pure VBA implementation with no external dependencies; drop the class module into any Office project.

## Theoretical Foundations

This module treats the dynamic array as the underlying storage, mirroring the classic **vector** or **ArrayList** pattern found in many languages. Resizing factors (×2, ×1.5) balance time vs. space complexity. QuickSort provides reliable average-case sorting performance, reflecting well-studied divide-and-conquer algorithms.

## Practical Considerations

* **Initial Capacity**
  Default is 8. For known large workloads, consider adding a constructor or factory to pre-set `Capacity` via repeated `AddRange` of an empty array.

* **Large Lists**
  For very large collections, memory fragmentation in VBA may occur; monitor `Capacity` and use `Trim` to reclaim unused space.

* **Type Homogeneity**
  While the list accepts mixed types, sorting and comparison assume homogeneous, comparable elements; mixing incomparable types may raise run-time errors.

* **Threading and Reentrancy**
  As with all VBA, this class is not thread-safe. Avoid modifying a list from multiple event handlers without guarding logic.

## Summary

The **List** module equips VBA developers with a flexible, high-performance collection type, abstracting array mechanics and enriching the language with modern list operations. Whether you need dynamic growth, complex sorting, or bulk manipulation, this class simplifies code, reduces bugs, and elevates maintainability in Office macros and add-ins.
