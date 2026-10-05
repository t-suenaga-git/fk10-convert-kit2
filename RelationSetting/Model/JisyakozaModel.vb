Namespace Njc.Model

    ''' <summary>
    ''' 10入金項目モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_jisyakoza_Model

        Public Property Jisya_no() As Njc.Common.SetArrayData
        Public Property Jisya_name() As Njc.Common.SetArrayData
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Jisya_no = New Njc.Common.SetArrayData(int)
            Jisya_name = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

End Namespace
