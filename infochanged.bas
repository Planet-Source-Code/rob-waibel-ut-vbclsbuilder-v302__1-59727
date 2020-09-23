Public Function InfoChanged(vFirst As Variant, vSecond As Variant) As Boolean
    
    If IsNull(vFirst) And IsNull(vSecond) Then
        InfoChanged = False
        Exit Function
    End If
    If IsNull(vFirst) And Not IsNull(vSecond) Then
        InfoChanged = True
        Exit Function
    End If
    If Not IsNull(vFirst) And IsNull(vSecond) Then
        InfoChanged = True
        Exit Function
    End If
    If vFirst <> vSecond Then
        InfoChanged = True
        Exit Function
    End If

End Function
