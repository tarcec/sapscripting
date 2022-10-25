Function Decrypt(ByVal szData As String) As String
    ''' This key value can be changed to alter the
    ''' encryption, but it must be the same for both
    ''' encryption and decryption.
    Const KEY_TEXT As String = "CORVFhMK7wZqi0oqNsQ5LJ5"
    
    ''' The KEY_OFFSET is optional, and may be any
    ''' value 0-64.
    ''' Likewise, it needs to be the same coming/going.
    Const KEY_OFFSET As Long = 38

    Dim bytKey() As Byte
    Dim bytData() As Byte
    Dim lNum As Long
    Dim szKey As String

    For lNum = 1 To ((Len(szData) \ Len(KEY_TEXT)) + 1)
        szKey = szKey & KEY_TEXT
    Next lNum

    bytKey = Left$(szKey, Len(szData))
    bytData = szData

    For lNum = LBound(bytData) To UBound(bytData)
        If lNum Mod 2 Then
            bytData(lNum) = bytData(lNum) Xor (bytKey(lNum) _
                + KEY_OFFSET)
        Else
            bytData(lNum) = bytData(lNum) Xor (bytKey(lNum) _
                - KEY_OFFSET)
        End If
    Next lNum

    szEncryptDecrypt = bytData
End Function

Function szEncryptDecrypt(ByVal szData As String) As String
    ''' This key value can be changed to alter the
    ''' encryption, but it must be the same for both
    ''' encryption and decryption.
    Const KEY_TEXT As String = "CORVFhMK7wZqi0oqNsQ5LJ5"
    
    ''' The KEY_OFFSET is optional, and may be any
    ''' value 0-64.
    ''' Likewise, it needs to be the same coming/going.
    Const KEY_OFFSET As Long = 38

    Dim bytKey() As Byte
    Dim bytData() As Byte
    Dim lNum As Long
    Dim szKey As String

    For lNum = 1 To ((Len(szData) \ Len(KEY_TEXT)) + 1)
        szKey = szKey & KEY_TEXT
    Next lNum

    bytKey = Left$(szKey, Len(szData))
    bytData = szData

    For lNum = LBound(bytData) To UBound(bytData)
        If lNum Mod 2 Then
            bytData(lNum) = bytData(lNum) Xor (bytKey(lNum) _
                + KEY_OFFSET)
        Else
            bytData(lNum) = bytData(lNum) Xor (bytKey(lNum) _
                - KEY_OFFSET)
        End If
    Next lNum

    szEncryptDecrypt = CStr(bytData)
End Function





