Late Binding can cause unexpected runtime errors if an object doesn't support a method or property you're trying to access.  Consider this:

```vbscript
Dim obj
Set obj = CreateObject("Some.Object")
MsgBox obj.NonExistentMethod
```

If `Some.Object` doesn't have `NonExistentMethod`, this will throw a runtime error.  Early binding (declaring object type) helps prevent this but requires more upfront coding.