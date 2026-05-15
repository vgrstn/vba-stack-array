VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Generic stack class."
'@IgnoreModule UseMeaningfulName, ProcedureNotUsed
'@Folder("Class")
'@ModuleDescription("Generic stack class.")

'------------------------------------------------------------------------------
' MIT License
'
' Copyright (c) 2025 Vincent van Geerestein
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'------------------------------------------------------------------------------

Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Author: Vincent van Geerestein
' E-mail: vincent@vangeerestein.com
' Description: Generic stack class
' Enumeration: For Each item in object
' Add-in: RubberDuck (https://rubberduckvba.com/)
' Version: 2025.10.17
'
' Methods
' Clear             Clears the stack
' Push item         Adds an item at the top of the stack
' Pop               Returns and removes the item at the top of the stack
' Items([base])     Exports all stack items to an array
'
' Properties
' Count             Returns the number of items in the stack
' IsEmpty           Returns True if the stack is empty or False if not
' Peek              Returns the item at the top of the stack (default)
'
' This generic Stack Class is implemented using an encapsulated VB Array.
'
' Timings (ms) for one push + one pop.
'     count      Array   Collection
' 1      10    0.00030      0.00053
' 2     100    0.00030      0.00053
' 3    1000    0.00029      0.00052
' 4   10000    0.00028      0.00052
' 5  100000    0.00029      0.00052
'
' The performance of the Array version is about 40% faster than the Collection
' version and it uses 50% less memory (for x64) per stored item, although extra
' memory is needed for the headspace (InitialSize).
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Performance Notes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Array-based stack is faster and more memory-efficient than using a Collection.
' Shrink logic avoids fragmentation and excessive resizing.
' Enumeration is rarely used in stack workflows, but supported for completeness.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Selected VB errors.
Private Enum VBERROR
    vbErrorInvalidProcedureCall = 5
End Enum

' InitialSize is the initial allocation size of the dynamic array.
Private Const InitialSize As Long = 100

' GrowthFactor is the multiplier applied on both grow and shrink.
Private Const GrowthFactor As Double = 2

' ShrinkTrigger is the occupancy fraction below which a shrink fires.
' Must be < 1 / GrowthFactor to prevent thrashing on alternating push/pop.
Private Const ShrinkTrigger As Double = 0.5 / GrowthFactor

' Wrapper for private data.
Private Type TPRIVATE
    Top As Long
    Stack() As Variant
    List As Collection
End Type
Private this As TPRIVATE


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Class_Initialize()

    ReDim this.Stack(1 To InitialSize)

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Enumerator
'@Description "Enumerate stack from top to bottom."
Public Property Get Enumerate() As IEnumVARIANT
' Called by For Each by setting Attribute Enumerate.VB_UserMemId = -4.
' [_NewEnum] is the name of the hidden method to enumerate a VB Collection.
' The internal stack array is first copied to an intermediate VB Collection.
'
' A snapshot Collection is built on each call and stored in this.List.
' Push or Pop sets this.List = Nothing, but the enumerator holds its own
' reference, so the snapshot survives until the For Each loop completes.
' Mutation during iteration is therefore safe but iterates the old snapshot.

    Set this.List = New Collection
    Dim i As Long
    For i = this.Top To 1 Step -1
        this.List.Add this.Stack(i)
    Next
    Set Enumerate = this.List.[_NewEnum]

End Property


'@Description "Removes all items from the stack."
Public Sub Clear()

    ReDim this.Stack(1 To InitialSize)
    this.Top = 0
    Set this.List = Nothing

End Sub


'@Description "Returns the number of items contained in the stack."
Public Property Get Count() As Long

    Count = this.Top

End Property


'@Description "Returns True if the stack is empty or False if not."
Public Property Get IsEmpty() As Boolean

    IsEmpty = (this.Top = 0)

End Property


'@DefaultMember
'@Description "Returns the item at the top of the stack."
Public Property Get Peek() As Variant

    If this.Top = 0 Then Err.Raise vbErrorInvalidProcedureCall, VBA.TypeName(Me) & ".Peek", "Stack is empty"

    If VBA.IsObject(this.Stack(this.Top)) Then
        Set Peek = this.Stack(this.Top)
    Else
        Peek = this.Stack(this.Top)
    End If

End Property


'@Description "Returns and removes the item at the top of the stack."
Public Function Pop() As Variant

    If this.Top = 0 Then Err.Raise vbErrorInvalidProcedureCall, VBA.TypeName(Me) & ".Pop", "Stack is empty"

    If VBA.IsObject(this.Stack(this.Top)) Then
        Set Pop = this.Stack(this.Top)
    Else
        Pop = this.Stack(this.Top)
    End If

    ' Tidy up especially for an object or a string.
    this.Stack(this.Top) = Empty

    this.Top = this.Top - 1

    ' Reclaim space.
    If this.Top < ShrinkTrigger * UBound(this.Stack) Then
        Dim newSize As Long: newSize = Int(UBound(this.Stack) / GrowthFactor)
        If newSize < InitialSize Then newSize = InitialSize
        ReDim Preserve this.Stack(1 To newSize)
    End If

    Set this.List = Nothing

End Function


'@Description "Adds an item at the top of the stack."
Public Sub Push(ByVal Item As Variant)

    ' Increase the headspace if there is none left.
    If this.Top = UBound(this.Stack) Then
        ReDim Preserve this.Stack(1 To Int(GrowthFactor * this.Top))
    End If

    this.Top = this.Top + 1

    If VBA.IsObject(Item) Then
        Set this.Stack(this.Top) = Item
    Else
        this.Stack(this.Top) = Item
    End If

    Set this.List = Nothing

End Sub


'@Description "Exports all stack items to an array with most recently pushed item as first array element."
Public Function Items(Optional ByVal base As Long) As Variant()

    Dim arr() As Variant: ReDim arr(base To base + this.Top - 1)
    Dim j As Long: j = base
    Dim i As Long
    For i = this.Top To 1 Step -1
        If VBA.IsObject(this.Stack(i)) Then
            Set arr(j) = this.Stack(i)
        Else
            arr(j) = this.Stack(i)
        End If
        j = j + 1
    Next

    Items = arr

End Function
