Namespace Njc.Model

    ''' <summary>
    ''' V7取引態様モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_toritaiyo_Rev7_Model

        Public Property Taiyo_no() As Njc.Common.SetArrayData   '取引態様№
        Public Property Taiyo_name() As Njc.Common.SetArrayData '取引態様名称
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Taiyo_no = New Njc.Common.SetArrayData(int)
            Taiyo_name = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

End Namespace
