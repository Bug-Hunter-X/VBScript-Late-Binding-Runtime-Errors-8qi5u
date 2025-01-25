Error handling is crucial when using late binding in VBScript. This example adds error handling to gracefully manage situations where a method or property doesn't exist:

```vbscript
On Error Resume Next  'Enable error handling

Dim obj
Set obj = CreateObject("Some.Object")

If Err.Number <> 0 Then
    MsgBox "Error accessing object method or property: " & Err.Description
    Err.Clear
Else
    MsgBox obj.NonExistentMethod 'This might still fail if method doesn't exist
End If

Set obj = Nothing
```

Alternatively, for known objects, early binding is preferred for type safety and avoiding runtime errors:

```vbscript
'Requires adding a reference to the object library if not already present
Dim obj as Object ' Or specific object type if known
Set obj = CreateObject("Some.Object")
MsgBox obj.ValidMethod 'Only call existing methods to avoid errors
Set obj = Nothing
```