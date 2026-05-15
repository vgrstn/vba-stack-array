# Generic Stack (VBA, Array-Backed)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

A lightweight **LIFO stack** implemented on top of a dynamic **array** with doubling growth and lazy shrink.  
Designed for **speed and memory efficiency** versus a `Collection`-backed stack, while preserving a clean API.

---

## 📦 Features

- **Fast push/pop** using a dynamic **array** with exponential growth (`GrowthFactor = 2`)
- **Lower memory use** than `Collection` (≈2× less per item on x64; extra headspace applies)
- **Shrink logic**: reclaims capacity when occupancy drops below `ShrinkTrigger` (25%)
- **Enumeration**: `For Each` (top → bottom) via a snapshot `Collection`
- **Utility export**: `Items([base])` returns a 0- or 1-based array copy
- Pure VBA, no external references

---

## ⚙️ Public Interface

| Member             | Type       | Description |
|-------------------|------------|-------------|
| `Push(Item)`       | `Sub`      | Adds an item at the **top**. |
| `Pop()`            | `Function` | Returns **and removes** the top item. Raises error 5 if empty. |
| `Peek` *(Default)* | `Property` | Returns the top item **without** removing it. Raises error 5 if empty. |
| `Count`            | `Property` | Number of items currently stored. |
| `IsEmpty`          | `Property` | `True` if empty, else `False`. |
| `Clear`            | `Sub`      | Clears the stack and resets capacity to `InitialSize`. |
| `Items([base])`    | `Function` | Copies items into `Variant()`; order is **top → bottom**; `arr(base)` = top. |
| `For Each`         | Enumerator | Iterates **top → bottom** on a snapshot. (Mutation during iteration is safe but iterates the old snapshot.) |

**Errors**  
- Empty stack on `Peek`/`Pop` raises **`vbErrorInvalidProcedureCall (=5)`** with source `"Stack.Peek"` / `"Stack.Pop"`.

---

## 🚀 Quick Start

```vb
Dim s As New Stack

s.Push "alpha"
s.Push "beta"
Debug.Print s.Peek        ' -> beta
Debug.Print s.Pop         ' -> beta (removed)
Debug.Print s.Pop         ' -> alpha (removed)
Debug.Print s.IsEmpty     ' -> True

```

## ⚡ Performance Notes

The array-based implementation provides **constant-time O(1)** amortized `Push` and `Pop` operations,  
while outperforming the `Collection`-based version in speed and memory use.

### Timings (ms) per Push + Pop cycle

| Count   | Array   | Collection |
|:-------:|--------:|-----------:|
| 10      | 0.00030 | 0.00053    |
| 100     | 0.00030 | 0.00053    |
| 1 000   | 0.00029 | 0.00052    |
| 10 000  | 0.00028 | 0.00052    |
| 100 000 | 0.00029 | 0.00052    |

### Observations

- The **array version** is roughly **40% faster** than the `Collection` version across all stack sizes.
- It uses about **half the memory** per stored item on x64 systems (excluding the headspace buffer).
- The **doubling growth strategy** (`GrowthFactor = 2`) keeps `ReDim Preserve` calls rare and amortized O(1).
- **Shrink** fires only when occupancy drops below `ShrinkTrigger` (25%), preventing thrashing on alternating push/pop.
- The **collection version** stays stable at very large sizes but pays a constant late-binding penalty internally.

### Practical Guidance

- For computational stacks, recursion simulations, or evaluation engines prefer the **array stack** for speed and memory efficiency.
- For unbounded or long-lived stacks (e.g., message buffers), the **collection stack** may offer simpler code.

> All timings were measured in VBA 7 (x64) on Windows 11 using a high-resolution `Stopwatch` based on `QueryPerformanceCounter`.

---
