Namespace Njc.Model

    ''' <summary>
    ''' V7FB情報モデル
    ''' </summary>
    ''' <remarks></remarks>
    Public Class M_FBInfo_Furiirai_Rev7_Model

        Public Property Fmt_syubetu() As Njc.Common.SetArrayData
        Public Property Fkom_no() As Njc.Common.SetArrayData
        Public Property Fkom_name() As Njc.Common.SetArrayData
        Public Property Kinyu_no() As Njc.Common.SetArrayData
        Public Property Kinyu_name() As Njc.Common.SetArrayData
        Public Property Ten_no() As Njc.Common.SetArrayData
        Public Property Ten_name() As Njc.Common.SetArrayData
        Public Property Kosyu_no() As Njc.Common.SetArrayData
        Public Property Koza_no() As Njc.Common.SetArrayData
        Public Property Koza_meigi() As Njc.Common.SetArrayData
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Fmt_syubetu = New Njc.Common.SetArrayData(int)
            Fkom_no = New Njc.Common.SetArrayData(int)
            Fkom_name = New Njc.Common.SetArrayData(int)
            Kinyu_no = New Njc.Common.SetArrayData(int)
            Kinyu_name = New Njc.Common.SetArrayData(int)
            Ten_no = New Njc.Common.SetArrayData(int)
            Ten_name = New Njc.Common.SetArrayData(int)
            Kosyu_no = New Njc.Common.SetArrayData(int)
            Koza_no = New Njc.Common.SetArrayData(int)
            Koza_meigi = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

    Public Class M_FBInfo_Kozafurikae_Rev7_Model

        Public Property Fmt_syubetu() As Njc.Common.SetArrayData
        Public Property Fkae_no() As Njc.Common.SetArrayData
        Public Property Fkae_name() As Njc.Common.SetArrayData
        Public Property Kinyu_no() As Njc.Common.SetArrayData
        Public Property Kinyu_name() As Njc.Common.SetArrayData
        Public Property Ten_no() As Njc.Common.SetArrayData
        Public Property Ten_name() As Njc.Common.SetArrayData
        Public Property Kosyu_no() As Njc.Common.SetArrayData
        Public Property Koza_no() As Njc.Common.SetArrayData
        Public Property Koza_meigi() As Njc.Common.SetArrayData
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Fmt_syubetu = New Njc.Common.SetArrayData(int)
            Fkae_no = New Njc.Common.SetArrayData(int)
            Fkae_name = New Njc.Common.SetArrayData(int)
            Kinyu_no = New Njc.Common.SetArrayData(int)
            Kinyu_name = New Njc.Common.SetArrayData(int)
            Ten_no = New Njc.Common.SetArrayData(int)
            Ten_name = New Njc.Common.SetArrayData(int)
            Kosyu_no = New Njc.Common.SetArrayData(int)
            Koza_no = New Njc.Common.SetArrayData(int)
            Koza_meigi = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

    Public Class M_FBInfo_Nssetting_Rev7_Model

        Public Property Fmt_syubetu() As Njc.Common.SetArrayData
        Public Property Ns_no() As Njc.Common.SetArrayData
        Public Property Ns_name() As Njc.Common.SetArrayData
        Public Property ItemCnt As Integer

        Public Sub New(int As Integer)

            Fmt_syubetu = New Njc.Common.SetArrayData(int)
            Ns_no = New Njc.Common.SetArrayData(int)
            Ns_name = New Njc.Common.SetArrayData(int)
            ItemCnt = int + 1

        End Sub

    End Class

End Namespace
