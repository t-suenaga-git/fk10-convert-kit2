Namespace Njc.Model

    ''' <summary>
    ''' V7入金区分モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_nkbn_Rev7_Model

        Public Property Nkbn_no() As Njc.Common.SetArrayData        '入金区分No
        Public Property Nkbn_name() As Njc.Common.SetArrayData      '入金区分名称
        Public Property Nkbn_sname() As Njc.Common.SetArrayData     '入金区分名称
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Nkbn_no = New Njc.Common.SetArrayData(int)
            Nkbn_name = New Njc.Common.SetArrayData(int)
            Nkbn_sname = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1
        End Sub

    End Class

End Namespace
