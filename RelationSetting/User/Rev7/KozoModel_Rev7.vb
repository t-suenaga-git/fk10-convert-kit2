Namespace Njc.Model

    ''' <summary>
    ''' V7構造モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_kozo_Rev7_Model

        Public Property Kozo_no() As Njc.Common.SetArrayData    '構造№
        Public Property Kozo_name() As Njc.Common.SetArrayData  '構造名称
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Kozo_no = New Njc.Common.SetArrayData(int)
            Kozo_name = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1
        End Sub

    End Class

End Namespace
