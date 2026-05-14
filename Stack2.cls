VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack2"
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
' 1      10    0.00025      0.00049
' 2     100    0.00025      0.00049
' 3    1000    0.00025      0.00050
' 4   10000    0.00025      0.00049
' 5  100000    0.00281      0.00050
'
' The performance of the Array version is about 2x faster than the Collection
' version and it uses 2x less memory (for x64) per stored item, although extra
' memory is needed for the headspace (ChunkSize). At very large stack size the
' Collection version is still stable, the Array version not.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Performance Notes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Array-based stack is faster and more memory-efficient than using a Collection.
' Shrink logic avoids fragmentation and excessive resizing.
' Enumeration is rarely used in queue workflows, but supported for completeness.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Selected VB errors.
Private Enum VBERROR
    vbErrorInvalidProcedureCall = 5
End Enum

' ChunkSize is the size by which the dynamic array grows or shrinks.
Private Const ChunkSize As Long = 100

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
    
    ReDim this.Stack(1 To ChunkSize)
    this.Top = 0
    Set this.List = Nothing

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Enumerator
'@Description "Enumerate stack from top to bottom."
Public Property Get Enumerate() As IEnumVARIANT
Attribute Enumerate.VB_Description = "Enumerate stack from top to bottom."
Attribute Enumerate.VB_UserMemId = -4
' Called by For Each by setting Attribute Enumerate.VB_UserMemId = -4.
' [_NewEnum] is the name of the hidden method to enumerate a VB Collection.
' The internal stack array is first copied to an intermediate VB Collection.
'
' Mutating Stack during the For Each loop does not break the enumeration but
' will give unpredictable results. After a For Each loop completes, this.List
' remains intact. VB does not expose a standard hook to detect when enumeration
' has finished. To keep it simple, any Push or Pop will reset the enumerator.
'
' While the current enumerator remains valid, modifying the stack during
' iteration may lead to unpredictable behavior. This reflects VB's native
' semantics: altering a collection mid-enumeration can produce undefined or
' erratic runtime results.

    Set this.List = New Collection
    Dim i As Long
    For i = this.Top To 1 Step -1
        this.List.Add this.Stack(i)
    Next
    Set Enumerate = this.List.[_NewEnum]

End Property


'@Description "Removes all items from the stack."
Public Sub Clear()
Attribute Clear.VB_Description = "Removes all items from the stack."
    
    ReDim this.Stack(1 To ChunkSize)
    this.Top = 0

End Sub


'@Description "Returns the number of items contained in the stack."
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of items contained in the stack."
    
    Count = this.Top

End Property


'@Description "Returns True if the stack is empty or False if not."
Public Property Get IsEmpty() As Boolean
Attribute IsEmpty.VB_Description = "Returns True if the stack is empty or False if not."
    
    IsEmpty = (this.Top = 0)

End Property


'@DefaultMember
'@Description "Returns the item at the top of the stack."
Public Property Get Peek() As Variant
Attribute Peek.VB_Description = "Returns the item at the top of the stack."
Attribute Peek.VB_UserMemId = 0
    
    If this.Top = 0 Then Err.Raise vbErrorInvalidProcedureCall, TypeName(Me) & ".Peek", "Stack is empty"
    If VBA.IsObject(this.Stack(this.Top)) Then
        Set Peek = this.Stack(this.Top)
    Else
        Peek = this.Stack(this.Top)
    End If

End Property


'@Description "Returns and removes the item at the top of the stack."
Public Function Pop() As Variant
Attribute Pop.VB_Description = "Returns and removes the item at the top of the stack."
    
    If this.Top = 0 Then Err.Raise vbErrorInvalidProcedureCall, TypeName(Me) & ".Pop", "Stack is empty"
    If VBA.IsObject(this.Stack(this.Top)) Then
        Set Pop = this.Stack(this.Top)
    Else
        Pop = this.Stack(this.Top)
    End If
    
    ' Tidy up especially for an object or a string.
    this.Stack(this.Top) = Empty
    
    this.Top = this.Top - 1
    
    ' Reclaim the headspace if it has become two chunks large.
    Const ChunkSize2 As Long = 2 * ChunkSize
    If UBound(this.Stack) = this.Top + ChunkSize2 Then
        ReDim Preserve this.Stack(1 To this.Top + ChunkSize)
    End If

    Set this.List = Nothing

End Function


'@Description "Adds an item at the top of the stack."
Public Sub Push(ByVal Item As Variant)
Attribute Push.VB_Description = "Adds an item at the top of the stack."
    
    ' Increase the headspace if there is none left.
    If UBound(this.Stack) = this.Top Then
        ReDim Preserve this.Stack(1 To this.Top + ChunkSize)
    End If
    
    this.Top = this.Top + 1
    
    If VBA.IsObject(Item) Then
        Set this.Stack(this.Top) = Item
    Else
        this.Stack(this.Top) = Item
    End If
    
    Set this.List = Nothing

End Sub


'@Description "Exports all stack items to an array."
Public Function Items(Optional ByVal base As Long) As Variant()
Attribute Items.VB_Description = "Exports all stack items to an array."

    If this.Top = 0 Then
        Items = Array()
        Exit Function
    End If
    
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
