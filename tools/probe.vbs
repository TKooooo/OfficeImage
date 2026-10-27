Option Explicit
' Office-equivalent probe: create the add-in through IDispatch late binding
' and call OnConnection the same way Word/Excel does.
' custom is a real VARIANT array (VT_VARIANT|VT_ARRAY) passed ByRef,
' Application is Nothing so the method body hits the catch block,
' which still proves the marshalling layer works.

Dim obj
Set obj = CreateObject("ImageZoomAddin.Connect")
WScript.Echo "created: " & TypeName(obj)

Dim custom()
On Error Resume Next
obj.OnConnection Nothing, 2, Nothing, custom
If Err.Number <> 0 Then
    WScript.Echo "invoke FAILED hr=0x" & Hex(Err.Number) & "  " & Err.Description
Else
    WScript.Echo "invoke OK"
End If
On Error GoTo 0
