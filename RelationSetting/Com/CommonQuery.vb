Namespace Njc.Query

    Public Class CommonQuery

        ''' <summary>
        ''' 仮テーブル削除(DROP)クエリ
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Function Qry_TmpTbl_Drop(tblname As String)

            Dim strsql As String = ""

            strsql = strsql & " IF EXISTS ("
            strsql = strsql & " SELECT * FROM dbo.sysobjects WHERE id = object_id(N'[dbo].["
            strsql = strsql & tblname
            strsql = strsql & "]') AND OBJECTPROPERTY(id, N'IsTable') = 1)"
            strsql = strsql & " DROP TABLE [dbo].["
            strsql = strsql & tblname
            strsql = strsql & "]"

            Return strsql

        End Function

        ''' <summary>
        ''' 仮テーブル挿入クエリ
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Function Qry_TmpTbl_Insert(tblname As String, ByVal value As String)

            Dim strsql As String = ""

            strsql = strsql & " INSERT INTO "
            strsql = strsql & tblname
            strsql = strsql & " VALUES ("
            strsql = strsql & value
            strsql = strsql & ")"
            Return strsql

        End Function

        ''' <summary>
        ''' 仮テーブル作成クエリ
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Function Qry_TmpTbl_Create(tblname As String, ByVal value As Object)

            Dim strsql As String = ""
            Dim tmp_str As String = ""

            For cntii = 0 To value.fieldcnt
                Select Case cntii
                    Case 0
                        tmp_str = tmp_str & "[" & value.FieldName(cntii) & "]"
                    Case Else
                        tmp_str = tmp_str & ",[" & value.FieldName(cntii) & "]"
                End Select

                If InStr(value.FieldName(cntii), "№") <> 0 Then
                    tmp_str = tmp_str & " INT"
                Else
                    tmp_str = tmp_str & " NCHAR(50)"
                End If
            Next

            strsql = strsql & " CREATE TABLE " & tblname & " ("
            strsql = strsql & tmp_str
            strsql = strsql & " )"

            Return strsql

        End Function

    End Class

End Namespace