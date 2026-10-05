Namespace Njc.Model

    ''' <summary>
    ''' V7口座種別モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_kozasyu_Rev7_Model

        Public Property Kosyu_no() As Njc.Common.SetArrayData    '口座種別№
        Public Property Kosyu_name() As Njc.Common.SetArrayData  '口座種別名称
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Kosyu_no = New Njc.Common.SetArrayData(int)
            Kosyu_name = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1
        End Sub

    End Class

End Namespace
