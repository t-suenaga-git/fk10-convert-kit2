Namespace Njc.Model

    ''' <summary>
    ''' V7設備モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_setubi_Rev7_Model

        Public Property Setubi_no() As Njc.Common.SetArrayData    '設備№
        Public Property Setubi_name() As Njc.Common.SetArrayData  '設備名称
        Public Property Setubi_lstno() As Njc.Common.SetArrayData  '設備内容№
        Public Property Setubi_lstname() As Njc.Common.SetArrayData  '設備内容名称
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Setubi_no = New Njc.Common.SetArrayData(int)
            Setubi_name = New Njc.Common.SetArrayData(int)
            Setubi_lstno = New Njc.Common.SetArrayData(int)
            Setubi_lstname = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1
        End Sub

    End Class

End Namespace
