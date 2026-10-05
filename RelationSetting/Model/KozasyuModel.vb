Namespace Njc.Model

    ''' <summary>
    ''' V10口座種別モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_kozasyu_Model

        Public Property Kosyu_no() As Njc.Common.SetArrayData
        Public Property Kosyu_name() As Njc.Common.SetArrayData
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Kosyu_no = New Njc.Common.SetArrayData(int)
            Kosyu_name = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

End Namespace
