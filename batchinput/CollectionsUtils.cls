Public Function ExistKey(ByRef collectionInstance As Collection, _
    key As String) As Boolean
On Error GoTo KeyDontExists
    Dim o As Variant
    o = collectionInstance.Item(key)
    ExistKey = True
    Exit Function
KeyDontExists:
    ExistKey = False
End Function
