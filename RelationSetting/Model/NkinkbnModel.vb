Namespace Njc.Model

    ''' <summary>
    ''' V10入金区分モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_nkbn_Model

        Public Property Nkbn_no() As Njc.Common.SetArrayData            '入金区分№
        Public Property Nkbn_name() As Njc.Common.SetArrayData          '入金区分名称
        Public Property Nkbn_zokusei_no() As Njc.Common.SetArrayData    '入金区分属性№
        Public Property Nkbn_zokusei_name() As Njc.Common.SetArrayData  '入金区分属性名称
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Nkbn_no = New Njc.Common.SetArrayData(int)
            Nkbn_name = New Njc.Common.SetArrayData(int)
            Nkbn_zokusei_no = New Njc.Common.SetArrayData(int)
            Nkbn_zokusei_name = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

    ''' <summary>
    ''' V10入金区分モデル(属性)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_nkbn_z_Model

        Public Property Nkbn_zokusei_no() As Njc.Common.SetArrayData    '入金区分属性№
        Public Property Nkbn_zokusei_name() As Njc.Common.SetArrayData  '入金区分属性名称
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Nkbn_zokusei_no = New Njc.Common.SetArrayData(int)
            Nkbn_zokusei_name = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

End Namespace
