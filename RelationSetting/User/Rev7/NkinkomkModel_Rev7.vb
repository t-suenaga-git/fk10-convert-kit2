Namespace Njc.Model

    ''' <summary>
    ''' V7入金項目モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_nkin_Rev7_Model

        Public Property Nkin_no() As Njc.Common.SetArrayData    '入金項目№
        Public Property Nkin_name() As Njc.Common.SetArrayData  '入金項目名称
        Public Property Nkin_kbn() As Njc.Common.SetArrayData   '入金項目区分
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Nkin_no = New Njc.Common.SetArrayData(int)
            Nkin_name = New Njc.Common.SetArrayData(int)
            Nkin_kbn = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1
        End Sub

    End Class

End Namespace
