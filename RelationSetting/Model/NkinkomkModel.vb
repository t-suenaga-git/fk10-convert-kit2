Namespace Njc.Model

    ''' <summary>
    ''' 10入金項目モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_nkin_Model

        Public Property Nkin_no() As Njc.Common.SetArrayData            '入金項目№
        Public Property Nkin_name() As Njc.Common.SetArrayData          '入金項目名称
        Public Property Nkin_ruino() As Njc.Common.SetArrayData         '入金区分
        Public Property Nkin_ruiname() As Njc.Common.SetArrayData       '入金区分名
        Public Property Nkin_zkseino() As Njc.Common.SetArrayData       '入金属性No
        Public Property Nkin_zkseiname() As Njc.Common.SetArrayData     '入金属性名
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Nkin_no = New Njc.Common.SetArrayData(int)
            Nkin_name = New Njc.Common.SetArrayData(int)
            Nkin_ruino = New Njc.Common.SetArrayData(int)
            Nkin_ruiname = New Njc.Common.SetArrayData(int)
            Nkin_zkseino = New Njc.Common.SetArrayData(int)
            Nkin_zkseiname = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

    ''' <summary>
    ''' 10入金項目モデル(属性)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_nkin_z_Model

        Public Property Nkin_zkseino() As Njc.Common.SetArrayData        '入金項目属性№
        Public Property Nkin_zkseiname() As Njc.Common.SetArrayData      '入金項目属性名称
        Public Property Nkin_ruino() As Njc.Common.SetArrayData         '入金区分
        Public Property Nkin_ruiname() As Njc.Common.SetArrayData       '入金区分名
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Nkin_zkseino = New Njc.Common.SetArrayData(int)
            Nkin_zkseiname = New Njc.Common.SetArrayData(int)
            Nkin_ruino = New Njc.Common.SetArrayData(int)
            Nkin_ruiname = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

    ''' <summary>
    ''' 10入金項目モデル(変動費メーター分類)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_nkin_hendometer_Model

        Public Property Nkin_hendometerno() As Njc.Common.SetArrayData        '入金項目属性№
        Public Property Nkin_hendometername() As Njc.Common.SetArrayData      '入金項目属性名称
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Nkin_hendometerno = New Njc.Common.SetArrayData(int)
            Nkin_hendometername = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

End Namespace
