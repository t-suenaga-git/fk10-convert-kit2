Namespace Njc.Model

    ''' <summary>
    ''' V7取引態様モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_TosiYoto_Rev7_Model

        Public Property TosiYoto_no() As Njc.Common.SetArrayData
        Public Property TosiYoto_name() As Njc.Common.SetArrayData
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            TosiYoto_no = New Njc.Common.SetArrayData(int)
            TosiYoto_name = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

End Namespace
