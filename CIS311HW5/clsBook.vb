Public Class clsBook
    Public Property strCategory As String
    Public Property intQuantity As Integer
    Public Property sngPrice As Single
    Public Property strTitle As String
    Public Property sngInventoryTotal As Single

    Public Sub New(ByVal strCategory As String, ByVal intQuantity As Integer, ByVal sngPrice As Single, ByVal strTitle As String, ByVal sngTotal As Single)
        Me.strCategory = strCategory
        Me.intQuantity = intQuantity
        Me.sngPrice = sngPrice
        Me.strTitle = strTitle
        Me.sngInventoryTotal = sngTotal
    End Sub

End Class
