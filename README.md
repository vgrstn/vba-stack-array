# Generic Stack (VBA, Array-Backed)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

A lightweight **LIFO stack** implemented on top of a dynamic **array** with chunked growth/shrink.  
Designed for **speed and memory efficiency** versus a `Collection`-backed stack, while preserving a clean API.

---

## 📦 Features

- **Fast push/pop** using a dynamic **array** with headspace (`ChunkSize`)
- **Lower memory use** than `Collection` (≈2× less per item on x64; extra headspace applies)
- **Shrink logic**: reclaims capacity when headspace grows to `2 * ChunkSize`
- **Enumeration**: `For Each` (top → bottom) via an intermediate `Collection`
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
| `Clear`            | `Sub`      | Clears the stack and resets capacity to one `ChunkSize`. |
| `Items([base])`    | `Function` | Copies items into `Variant()`; order is **top → bottom**. |
| `For Each`         | Enumerator | Iterates **top → bottom**. (Don’t mutate while iterating.) |

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
