Imports System.Data.SqlClient
Imports Converter10.Njc.Common
Imports Converter10.Njc.N3Lib.Utys
Imports System.Text
Imports System.Security.Cryptography

Module DataCondModule   '20160929 データ調整用メソッドの作成 -add

    ''' <summary>
    ''' 契約者口座情報更新クエリ 20160929 契約者口座情報の移行仕様変更に伴う修正 -add
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Get_List_Qry_KysKozaCond() As List(Of String)

        Dim rtn_list As New List(Of String)

        Dim tmp_sql As String = ""

        '金融機関区分
        tmp_sql = tmp_sql & " UPDATE kysdata_koza "
        tmp_sql = tmp_sql & " SET koza_yucyoflg = 1 "
        tmp_sql = tmp_sql & " WHERE koza_printkbn = 2 and koza_kbn = 1 "
        rtn_list.Add(tmp_sql)
        tmp_sql = ""

        '金融機関No
        tmp_sql = tmp_sql & " UPDATE kysdata_koza "
        tmp_sql = tmp_sql & " SET kinyu_no = 9900  "
        tmp_sql = tmp_sql & " WHERE koza_yucyoflg = 1 and koza_kbn = 1 and len(yucyokoza_kigo1) = 5 "
        rtn_list.Add(tmp_sql)
        tmp_sql = ""

        '支店No
        tmp_sql = tmp_sql & " UPDATE kysdata_koza "
        tmp_sql = tmp_sql & " SET kinyu_tenno = substring(yucyokoza_kigo1,2,3) "
        tmp_sql = tmp_sql & " WHERE koza_yucyoflg = 1 and koza_kbn = 1 and len(yucyokoza_kigo1) = 5 "
        rtn_list.Add(tmp_sql)
        tmp_sql = ""

        '種別
        tmp_sql = tmp_sql & " UPDATE kysdata_koza "
        tmp_sql = tmp_sql & " SET koza_syubetu = iif(left(yucyokoza_kigo1,1)=1,1,2) "
        tmp_sql = tmp_sql & " WHERE koza_yucyoflg = 1 and koza_kbn = 1 and len(yucyokoza_kigo1) = 5  "
        rtn_list.Add(tmp_sql)
        tmp_sql = ""

        '番号
        tmp_sql = tmp_sql & " UPDATE kysdata_koza "
        tmp_sql = tmp_sql & " SET koza_bango = left(yucyokoza_bango,7) "
        tmp_sql = tmp_sql & " WHERE koza_yucyoflg = 1 and koza_kbn = 1 and len(yucyokoza_bango) = 8 "
        rtn_list.Add(tmp_sql)
        tmp_sql = ""

        tmp_sql = tmp_sql & " UPDATE kysdata_koza "
        tmp_sql = tmp_sql & " SET koza_bango = left(yucyokoza_bango,6) "
        tmp_sql = tmp_sql & " WHERE koza_yucyoflg = 1 and koza_kbn = 1 and len(yucyokoza_bango) = 7 "
        rtn_list.Add(tmp_sql)
        tmp_sql = ""

        tmp_sql = tmp_sql & " UPDATE kysdata_koza "
        tmp_sql = tmp_sql & " SET koza_bango = left(yucyokoza_bango,5) "
        tmp_sql = tmp_sql & " WHERE koza_yucyoflg = 1 and koza_kbn = 1 and len(yucyokoza_bango) = 6 "
        rtn_list.Add(tmp_sql)
        tmp_sql = ""

        tmp_sql = tmp_sql & " UPDATE kysdata_koza "
        tmp_sql = tmp_sql & " SET koza_bango = left(yucyokoza_bango,4) "
        tmp_sql = tmp_sql & " WHERE koza_yucyoflg = 1 and koza_kbn = 1 and len(yucyokoza_bango) = 5 "
        rtn_list.Add(tmp_sql)
        tmp_sql = ""

        tmp_sql = tmp_sql & " UPDATE kysdata_koza "
        tmp_sql = tmp_sql & " SET koza_bango = left(yucyokoza_bango,3) "
        tmp_sql = tmp_sql & " WHERE koza_yucyoflg = 1 and koza_kbn = 1 and len(yucyokoza_bango) = 4 "
        rtn_list.Add(tmp_sql)
        tmp_sql = ""

        tmp_sql = tmp_sql & " UPDATE kysdata_koza "
        tmp_sql = tmp_sql & " SET koza_bango = left(yucyokoza_bango,2) "
        tmp_sql = tmp_sql & " WHERE koza_yucyoflg = 1 and koza_kbn = 1 and len(yucyokoza_bango) = 3 "
        rtn_list.Add(tmp_sql)
        tmp_sql = ""

        tmp_sql = tmp_sql & " UPDATE kysdata_koza "
        tmp_sql = tmp_sql & " SET koza_bango = left(yucyokoza_bango,1) "
        tmp_sql = tmp_sql & " WHERE koza_yucyoflg = 1 and koza_kbn = 1 and len(yucyokoza_bango) = 2 "
        rtn_list.Add(tmp_sql)
        tmp_sql = ""

        tmp_sql = tmp_sql & " UPDATE kysdata_koza "
        tmp_sql = tmp_sql & " SET koza_bango = NULL "
        tmp_sql = tmp_sql & " WHERE koza_yucyoflg = 1 and koza_kbn = 1 and len(yucyokoza_bango) = 1 "
        rtn_list.Add(tmp_sql)
        tmp_sql = ""

        Return rtn_list

    End Function

    ''' <summary>
    ''' 自社口座情報のゆうちょ→銀行変換処理を行うクエリ作成処理
    ''' </summary>
    ''' <returns>データ調整クエリ(複数の文字列を格納したオブジェクト)</returns>
    ''' <remarks>
    ''' 20161125 自社口座情報のゆうちょ→銀行変換処理の追加 新規追加
    ''' </remarks>
    Public Sub Get_List_Qry_JisyaKozaCond(ByVal sqlcnnv10 As SqlConnection)

        Dim tmp_sql As String = ""
        Dim tmp_cnt As Integer = 0  'クエリ実行用の作業用変数

        '対象となる自社口座情報を取得
        Dim list_updatejisyakozainfo As New List(Of String)
        tmp_sql = tmp_sql & " SELECT "
        tmp_sql = tmp_sql & " 	CONVERT(VARCHAR(MAX),jisya_no) + '-' + CONVERT(VARCHAR(MAX),jisya_kozano) "
        tmp_sql = tmp_sql & " FROM jisyadata_koza WHERE kinyu_no IS NULL AND kinyu_tenno IS NULL AND koza_bango IS NULL "
        DBExec.Exec_DataReader_Col_List(tmp_sql, sqlcnnv10, list_updatejisyakozainfo)

        '対象となる自社口座情報が存在する場合は抽出用に成形(無ければ処理を抜ける)
        Dim tmp_jisyakey As String = ""
        If list_updatejisyakozainfo.Count = 0 Then
            Exit Sub
        Else
            For Each jisyakey In list_updatejisyakozainfo
                tmp_jisyakey = tmp_jisyakey & "," & "'" & jisyakey & "'"
            Next
            tmp_jisyakey = "(" & tmp_jisyakey.Remove(0, 1) & ")"
        End If

        '金融機関No
        tmp_sql = tmp_sql & " UPDATE jisyadata_koza "
        tmp_sql = tmp_sql & " SET kinyu_no = 9900  "
        tmp_sql = tmp_sql & " WHERE CONVERT(VARCHAR(MAX),jisya_no) + '-' + CONVERT(VARCHAR(MAX),jisya_kozano) IN " & tmp_jisyakey
        tmp_sql = tmp_sql & " AND len(yucyokoza_kigo1) = 5 "
        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)
        tmp_sql = ""

        '支店No
        tmp_sql = tmp_sql & " UPDATE jisyadata_koza "
        tmp_sql = tmp_sql & " SET kinyu_tenno = substring(yucyokoza_kigo1,2,3) "
        tmp_sql = tmp_sql & " WHERE CONVERT(VARCHAR(MAX),jisya_no) + '-' + CONVERT(VARCHAR(MAX),jisya_kozano) IN " & tmp_jisyakey
        tmp_sql = tmp_sql & " AND len(yucyokoza_kigo1) = 5 "
        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)
        tmp_sql = ""

        '種別
        tmp_sql = tmp_sql & " UPDATE jisyadata_koza "
        tmp_sql = tmp_sql & " SET koza_syubetu = iif(left(yucyokoza_kigo1,1)=1,1,2) "
        tmp_sql = tmp_sql & " WHERE CONVERT(VARCHAR(MAX),jisya_no) + '-' + CONVERT(VARCHAR(MAX),jisya_kozano) IN " & tmp_jisyakey
        tmp_sql = tmp_sql & " AND len(yucyokoza_kigo1) = 5 "
        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)
        tmp_sql = ""

        '番号
        tmp_sql = tmp_sql & " UPDATE jisyadata_koza "
        tmp_sql = tmp_sql & " SET koza_bango = left(yucyokoza_bango,7) "
        tmp_sql = tmp_sql & " WHERE CONVERT(VARCHAR(MAX),jisya_no) + '-' + CONVERT(VARCHAR(MAX),jisya_kozano) IN " & tmp_jisyakey
        tmp_sql = tmp_sql & " AND len(yucyokoza_bango) = 8 "
        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)
        tmp_sql = ""

        tmp_sql = tmp_sql & " UPDATE jisyadata_koza "
        tmp_sql = tmp_sql & " SET koza_bango = left(yucyokoza_bango,6) "
        tmp_sql = tmp_sql & " WHERE CONVERT(VARCHAR(MAX),jisya_no) + '-' + CONVERT(VARCHAR(MAX),jisya_kozano) IN " & tmp_jisyakey
        tmp_sql = tmp_sql & " AND len(yucyokoza_bango) = 7 "
        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)
        tmp_sql = ""

        tmp_sql = tmp_sql & " UPDATE jisyadata_koza "
        tmp_sql = tmp_sql & " SET koza_bango = left(yucyokoza_bango,5) "
        tmp_sql = tmp_sql & " WHERE CONVERT(VARCHAR(MAX),jisya_no) + '-' + CONVERT(VARCHAR(MAX),jisya_kozano) IN " & tmp_jisyakey
        tmp_sql = tmp_sql & " AND len(yucyokoza_bango) = 6 "
        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)
        tmp_sql = ""

        tmp_sql = tmp_sql & " UPDATE jisyadata_koza "
        tmp_sql = tmp_sql & " SET koza_bango = left(yucyokoza_bango,4) "
        tmp_sql = tmp_sql & " WHERE CONVERT(VARCHAR(MAX),jisya_no) + '-' + CONVERT(VARCHAR(MAX),jisya_kozano) IN " & tmp_jisyakey
        tmp_sql = tmp_sql & " AND len(yucyokoza_bango) = 5 "
        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)
        tmp_sql = ""

        tmp_sql = tmp_sql & " UPDATE jisyadata_koza "
        tmp_sql = tmp_sql & " SET koza_bango = left(yucyokoza_bango,3) "
        tmp_sql = tmp_sql & " WHERE CONVERT(VARCHAR(MAX),jisya_no) + '-' + CONVERT(VARCHAR(MAX),jisya_kozano) IN " & tmp_jisyakey
        tmp_sql = tmp_sql & " AND len(yucyokoza_bango) = 4 "
        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)
        tmp_sql = ""

        tmp_sql = tmp_sql & " UPDATE jisyadata_koza "
        tmp_sql = tmp_sql & " SET koza_bango = left(yucyokoza_bango,2) "
        tmp_sql = tmp_sql & " WHERE CONVERT(VARCHAR(MAX),jisya_no) + '-' + CONVERT(VARCHAR(MAX),jisya_kozano) IN " & tmp_jisyakey
        tmp_sql = tmp_sql & " AND len(yucyokoza_bango) = 3 "
        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)
        tmp_sql = ""

        tmp_sql = tmp_sql & " UPDATE jisyadata_koza "
        tmp_sql = tmp_sql & " SET koza_bango = left(yucyokoza_bango,1) "
        tmp_sql = tmp_sql & " WHERE CONVERT(VARCHAR(MAX),jisya_no) + '-' + CONVERT(VARCHAR(MAX),jisya_kozano) IN " & tmp_jisyakey
        tmp_sql = tmp_sql & " AND len(yucyokoza_bango) = 2 "
        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)
        tmp_sql = ""

        tmp_sql = tmp_sql & " UPDATE jisyadata_koza "
        tmp_sql = tmp_sql & " SET koza_bango = NULL "
        tmp_sql = tmp_sql & " WHERE CONVERT(VARCHAR(MAX),jisya_no) + '-' + CONVERT(VARCHAR(MAX),jisya_kozano) IN " & tmp_jisyakey
        tmp_sql = tmp_sql & " AND len(yucyokoza_bango) = 1 "
        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)
        tmp_sql = ""

    End Sub

    ''' <summary>
    ''' 適用税率一括更新処理 20161007 適用税率データ調整処理の追加 -add
    ''' </summary>
    ''' <param name="sqlcnnv10"></param>
    ''' <remarks></remarks>
    Public Sub Set_ZeiMst(ByVal sqlcnnv10 As SqlConnection)

        Dim tmptblname As String = "tmp_m_zei"

        '-------------------------------------------------------
        '消費税マスタの仮テーブルを作成する
        '-------------------------------------------------------
        '初期化
        Dim tmp_cnt As Integer = 0
        Dim tmp_sql_drop As String = DBQuery.Qry_DropInfo(tmptblname, True)
        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmp_cnt)

        '作成
        Dim tmp_sql_create As String = ""
        tmp_sql_create = tmp_sql_create & " CREATE TABLE " & tmptblname
        tmp_sql_create = tmp_sql_create & " 	( "
        tmp_sql_create = tmp_sql_create & " 		[No] INT, "
        tmp_sql_create = tmp_sql_create & " 		[ZeiRit] INT, "
        tmp_sql_create = tmp_sql_create & " 		[StartYmd] DATE, "
        tmp_sql_create = tmp_sql_create & " 		[EndYmd] DATE "
        tmp_sql_create = tmp_sql_create & " 	) "
        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_create, tmp_cnt)

        '挿入用のデータ取得
        Dim tmp_no() As String = DBExec.Exec_Scalar("SELECT REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX),zeidefine.query('/ArrayOfZeiDefine/ZeiDefine/No')),'</No><No>',','),'<No>',''),'</No>','') FROM profile_fk", sqlcnnv10).ToString.Split(",")
        Dim tmp_zeirit() As String = DBExec.Exec_Scalar("SELECT REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX),zeidefine.query('/ArrayOfZeiDefine/ZeiDefine/ZeiRit')),'</ZeiRit><ZeiRit>',','),'<ZeiRit>',''),'</ZeiRit>','') FROM profile_fk", sqlcnnv10).ToString.Split(",")
        Dim tmp_staymd() As String = DBExec.Exec_Scalar("SELECT REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX),zeidefine.query('/ArrayOfZeiDefine/ZeiDefine/StartYmd')),'</StartYmd><StartYmd>',','),'<StartYmd>',''),'</StartYmd>','') FROM profile_fk", sqlcnnv10).ToString.Split(",")
        Dim tmp_endymd() As String = DBExec.Exec_Scalar("SELECT REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX),zeidefine.query('/ArrayOfZeiDefine/ZeiDefine/EndYmd')),'</EndYmd><EndYmd>',','),'<EndYmd>',''),'</EndYmd>','') FROM profile_fk", sqlcnnv10).ToString.Split(",")

        '挿入
        For cntii = 0 To UBound(tmp_no)
            Dim tmp_sql_insert As String = " INSERT INTO " & tmptblname & " VALUES(" & tmp_no(cntii) & "," & tmp_zeirit(cntii) & "," & "'" & tmp_staymd(cntii) & "'" & "," & "'" & tmp_endymd(cntii) & "'" & ")"
            DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_insert, tmp_cnt)
        Next

        '-------------------------------------------------------
        '仮テーブルを元にデータ調整
        '-------------------------------------------------------
        Dim tmp_sql As String = ""

        '送金ルール控除項目情報(システム日付基準)
        tmp_sql = tmp_sql & " UPDATE sorule_kojo_cmrule SET "
        tmp_sql = tmp_sql & " 	zei_rit = "
        tmp_sql = tmp_sql & " 		( "
        tmp_sql = tmp_sql & " 			SELECT ZeiRit FROM tmp_m_zei WHERE GETDATE() BETWEEN StartYmd AND EndYmd "
        tmp_sql = tmp_sql & " 		) "
        tmp_sql = tmp_sql & " WHERE kojo_gakzeikbn = 3 "
        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)
        tmp_sql = ""

        '契約入金項目情報(契約開始日基準)
        tmp_sql = tmp_sql & " UPDATE kydata_nkin "
        tmp_sql = tmp_sql & " SET zei_rit =  MZ.ZeiRit "
        tmp_sql = tmp_sql & " FROM kydata_nkin AS KYN "
        tmp_sql = tmp_sql & " LEFT JOIN kydata_kihon AS KYK "
        tmp_sql = tmp_sql & " ON KYN.ky_guid = KYK.ky_guid "
        tmp_sql = tmp_sql & " LEFT JOIN tmp_m_zei AS MZ "
        tmp_sql = tmp_sql & " ON KYK.kystart_ymd BETWEEN StartYmd AND EndYmd "
        tmp_sql = tmp_sql & " WHERE sq_zeikbn = 3 "
        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)
        tmp_sql = ""

        '契約次回入金項目情報(システム日付基準)
        tmp_sql = tmp_sql & " UPDATE kydata_nkin_nx SET "
        tmp_sql = tmp_sql & " 	zei_rit = "
        tmp_sql = tmp_sql & " 		( "
        tmp_sql = tmp_sql & " 			SELECT ZeiRit FROM tmp_m_zei WHERE GETDATE() BETWEEN StartYmd AND EndYmd "
        tmp_sql = tmp_sql & " 		) "
        tmp_sql = tmp_sql & " WHERE sq_zeikbn = 3 "
        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)
        tmp_sql = ""

        '未収滞納金情報(該当年月基準)
        tmp_sql = tmp_sql & " UPDATE unyotainodata "
        tmp_sql = tmp_sql & " SET zei_rit =  MZ.ZeiRit "
        tmp_sql = tmp_sql & " FROM unyotainodata AS TAINO "
        tmp_sql = tmp_sql & " LEFT JOIN tmp_m_zei MZ "
        tmp_sql = tmp_sql & " ON TAINO.gt_ym BETWEEN StartYmd AND EndYmd "
        tmp_sql = tmp_sql & " WHERE sq_zeikbn = 3 "
        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)
        tmp_sql = ""

        '家主固定控除情報(システム日付基準)
        tmp_sql = tmp_sql & " UPDATE koteirule SET "
        tmp_sql = tmp_sql & " 	zei_rit = "
        tmp_sql = tmp_sql & " 		( "
        tmp_sql = tmp_sql & " 			SELECT ZeiRit FROM tmp_m_zei WHERE GETDATE() BETWEEN StartYmd AND EndYmd "
        tmp_sql = tmp_sql & " 		) "
        tmp_sql = tmp_sql & " WHERE zei_kbn = 3 "
        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql, tmp_cnt)
        tmp_sql = ""

        '-------------------------------------------------------
        '消費税マスタの仮テーブルを削除
        '-------------------------------------------------------
        DBExec.Exec_NonQuery(sqlcnnv10, tmp_sql_drop, tmp_cnt)

    End Sub

    '20161028 メモタイトル一括更新処理 -add sta
    ''' <summary>
    ''' 契約者口座情報更新クエリ
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Get_List_Qry_BikoTitleCond() As List(Of String)

        Dim rtn_list As New List(Of String)

        Dim tmp_sql As String = ""

        'メモ5まで移行可能な項目
        '仲介業者メモ、保険業者メモ、家賃保証業者メモ、修繕業者メモ、自社メモ、契約者メモ、物件メモ、部屋メモ
        tmp_sql = tmp_sql & " UPDATE m_memo SET memo_name = 'メモ' + CONVERT(VARCHAR,memo_no) WHERE memo_kbn IN(1,2,4,5,7,9,10,11) AND memo_no <= 5 AND memo_name = '' "
        rtn_list.Add(tmp_sql)
        tmp_sql = ""

        'メモ10まで移行可能な項目
        '家主メモ、契約メモ
        tmp_sql = tmp_sql & " UPDATE m_memo SET memo_name = 'メモ' + CONVERT(VARCHAR,memo_no) WHERE memo_kbn IN(3,6) AND memo_no <= 10 AND memo_name = '' "
        rtn_list.Add(tmp_sql)
        tmp_sql = ""

        Return rtn_list

    End Function
    '20161028 メモタイトル一括更新処理 -add end

    ''' <summary>
    ''' 物件部屋所有者情報が両方存在する場合、区分所有優先のため物件所有者情報を削除するクエリの作成
    ''' </summary>
    ''' <returns>物件所有者情報を削除するクエリ</returns>
    ''' <remarks>
    ''' 20161124 物件部屋所有者情報有無による所有者情報のデータ調整処理を追加 新規追加
    ''' ①中間ファイルに物件部屋所有者情報いずれにもデータが登録されている場合、区分所有として移行される
    ''' ②物件所有者情報はそのまま移行されるため所有区分が区分所有かつ物件所有者情報にレコードが存在する状態になる
    ''' ③所有区分が区分所有(kubunsyo = 2)かつ物件所有者情報(bkdata_syo)にレコードが存在する場合は削除
    ''' </remarks>
    Public Function Get_UseQry_BkSyoDelete() As String

        Dim tmp_sql As String = ""
        tmp_sql = tmp_sql & " DELETE FROM bkdata_syo "
        tmp_sql = tmp_sql & " WHERE bk_guid IN "
        tmp_sql = tmp_sql & " 	( "
        tmp_sql = tmp_sql & " 		SELECT bk_guid FROM bkdata_detail WHERE kubunsyo = 2 "
        tmp_sql = tmp_sql & " 	) "

        Return tmp_sql

    End Function

    ''' <summary>
    ''' 物件部屋所有者情報が存在しない場合の仮レコード挿入用クエリの作成
    ''' </summary>
    ''' <returns>物件部屋所有者情報が存在しない場合の仮レコード作成用クエリ</returns>
    ''' <remarks>
    ''' 20161124 物件部屋所有者情報有無による所有者情報のデータ調整処理を追加 新規追加
    ''' </remarks>
    Public Function Get_List_UseQry_BkSyoInsert() As List(Of String)

        Dim rtn_list As New List(Of String)
        Dim tmp_sql As String = ""

        '一棟所有の仮レコード挿入
        tmp_sql = tmp_sql & " INSERT INTO bkdata_syo "
        tmp_sql = tmp_sql & " SELECT "
        tmp_sql = tmp_sql & " 	 bk_guid "
        tmp_sql = tmp_sql & " 	,1 AS kn_no "
        tmp_sql = tmp_sql & " 	,NEWID() AS sorule_guid "
        tmp_sql = tmp_sql & " 	,NULL AS kasi1_ow_no "
        tmp_sql = tmp_sql & " 	,NULL AS syo1_ow_no "
        tmp_sql = tmp_sql & " 	,NULL AS kasi2_ow_no "
        tmp_sql = tmp_sql & " 	,NULL AS syo2_ow_no "
        tmp_sql = tmp_sql & " 	,'19800101' AS syo_startymd "                       '20161124 レビュー結果：ここは19800101で問題ない？V7.適用開始月？ '20161125 レビュー指摘事項対応 19800101がデフォルトで設定されておりました。問題ないと思います。
        tmp_sql = tmp_sql & " 	,NULL AS syo_endymd "
        tmp_sql = tmp_sql & " 	,'<history><c>' + (SELECT REPLACE(CONVERT(VARCHAR,GetDate(),120),'-','/')) + '</c><chost>CONVUSER</chost><cosuser>CONVUSER</cosuser><cappuser>標準ユーザ</cappuser><u>' + (SELECT REPLACE(CONVERT(VARCHAR,GetDate(),120),'-','/')) + '</u><uhost>CONVUSER</uhost><uosuser>CONVUSER</uosuser><uappuser>標準ユーザ</uappuser></history>' AS history "
        tmp_sql = tmp_sql & " FROM bkdata "
        tmp_sql = tmp_sql & " WHERE bk_guid IN (SELECT bk_guid FROM bkdata_detail WHERE kubunsyo = 1) "
        tmp_sql = tmp_sql & " AND   bk_guid NOT IN (SELECT bk_guid FROM bkdata_syo) "
        rtn_list.Add(tmp_sql)
        tmp_sql = ""

        '区分所有の仮レコード挿入
        tmp_sql = tmp_sql & " INSERT INTO hydata_syo "
        tmp_sql = tmp_sql & " SELECT "
        tmp_sql = tmp_sql & " 	 hy_guid "
        tmp_sql = tmp_sql & " 	,1 AS kn_no "
        tmp_sql = tmp_sql & " 	,NEWID() AS sorule_guid "
        tmp_sql = tmp_sql & " 	,NULL AS kasi1_ow_no "
        tmp_sql = tmp_sql & " 	,NULL AS syo1_ow_no "
        tmp_sql = tmp_sql & " 	,NULL AS kasi2_ow_no "
        tmp_sql = tmp_sql & " 	,NULL AS syo2_ow_no "
        tmp_sql = tmp_sql & " 	,'19800101' AS syo_startymd "                       '20161124 レビュー結果：ここは19800101で問題ない？V7.適用開始月？ '20161125 レビュー指摘事項対応 19800101がデフォルトで設定されておりました。問題ないと思います。
        tmp_sql = tmp_sql & " 	,NULL AS syo_endymd "
        tmp_sql = tmp_sql & " 	,'<history><c>' + (SELECT REPLACE(CONVERT(VARCHAR,GetDate(),120),'-','/')) + '</c><chost>CONVUSER</chost><cosuser>CONVUSER</cosuser><cappuser>標準ユーザ</cappuser><u>' + (SELECT REPLACE(CONVERT(VARCHAR,GetDate(),120),'-','/')) + '</u><uhost>CONVUSER</uhost><uosuser>CONVUSER</uosuser><uappuser>標準ユーザ</uappuser></history>' AS history "
        tmp_sql = tmp_sql & " FROM hydata AS HY "
        tmp_sql = tmp_sql & " LEFT JOIN bkdata_detail AS BKD ON HY.bk_guid = BKD.bk_guid "
        tmp_sql = tmp_sql & " WHERE BKD.bk_guid IN (SELECT bk_guid FROM bkdata_detail WHERE kubunsyo = 2) "
        tmp_sql = tmp_sql & " AND   hy_guid NOT IN (SELECT hy_guid FROM hydata_syo) "
        rtn_list.Add(tmp_sql)
        tmp_sql = ""

        Return rtn_list

    End Function

    ''' <summary>
    ''' 契約情報のデータ調整
    ''' </summary>
    Public Sub Set_KydataCond(ByVal sqlcnnv10 As SqlConnection)

        Dim sb As New System.Text.StringBuilder
        Dim cnt As Integer = 0
        sb.AppendLine(" DELETE FROM kydata ")
        sb.AppendLine(" WHERE ky_guid NOT IN ")
        sb.AppendLine(" 	( ")
        sb.AppendLine(" 		SELECT ky_guid FROM kydata_kihon ")
        sb.AppendLine(" 	) ")
        DBExec.Exec_NonQuery(sqlcnnv10, sb.ToString, cnt)

    End Sub

    ''' <summary>
    ''' 契約者口座情報のデータ調整
    ''' </summary>
    Public Sub Set_KysdataKozaCond(ByVal sqlcnnv10 As SqlConnection)

        Dim sb As New System.Text.StringBuilder
        Dim cnt As Integer = 0
        sb.AppendLine(" /*変数・カーソル宣言*/ ")
        sb.AppendLine(" DECLARE @kysno INT ")
        sb.AppendLine(" DECLARE KysInfo CURSOR FOR ")
        sb.AppendLine(" 	SELECT kys_no FROM kysdata ")
        sb.AppendLine(" 	WHERE kys_no NOT IN ")
        sb.AppendLine(" 		( ")
        sb.AppendLine(" 			SELECT kys_no FROM kysdata_koza ")
        sb.AppendLine(" 		) ")
        sb.AppendLine(" 	ORDER BY 1 ")
        sb.AppendLine("  ")
        sb.AppendLine(" SET NOCOUNT ON ")
        sb.AppendLine(" OPEN KysInfo ")
        sb.AppendLine("  ")
        sb.AppendLine(" /*メイン処理*/ ")
        sb.AppendLine(" FETCH NEXT FROM KysInfo INTO @kysno ")
        sb.AppendLine(" WHILE @@FETCH_STATUS = 0 ")
        sb.AppendLine(" BEGIN ")
        sb.AppendLine(" 	INSERT INTO kysdata_koza VALUES(@kysno,1,1,NULL,NULL,1,NULL,'','','','','',NULL,NULL,'','','',1,NULL,2,1,NULL,NULL,'',1,0) ")
        sb.AppendLine(" 	INSERT INTO kysdata_koza VALUES(@kysno,2,2,NULL,NULL,1,NULL,'','','','','',NULL,NULL,'','','',1,NULL,2,1,NULL,NULL,'',1,0) ")
        sb.AppendLine(" 	INSERT INTO kysdata_koza VALUES(@kysno,3,3,NULL,NULL,1,NULL,'','','','','',NULL,NULL,'','','',NULL,NULL,NULL,NULL,NULL,NULL,NULL,1,0) ")
        sb.AppendLine(" 	FETCH NEXT FROM KysInfo INTO @kysno ")
        sb.AppendLine(" END ")
        sb.AppendLine("  ")
        sb.AppendLine(" /*終了処理*/ ")
        sb.AppendLine(" CLOSE KysInfo ")
        sb.AppendLine(" DEALLOCATE KysInfo ")
        DBExec.Exec_NonQuery(sqlcnnv10, sb.ToString, cnt)

    End Sub

    ''' <summary>
    ''' 当月分の毎月入金項目データ調整
    ''' </summary>
    Public Sub Set_KydataNkinCond(ByVal sqlcnnv10 As SqlConnection)

        Dim sb As New System.Text.StringBuilder
        Dim cnt As Integer = 0
        sb.AppendLine(" INSERT INTO kydata_nkin ")
        sb.AppendLine(" SELECT ")
        sb.AppendLine(" 	 ky_guid ")
        sb.AppendLine(" 	,ky_recno ")
        sb.AppendLine(" 	,1 AS tuki_kbn ")
        sb.AppendLine(" 	,nkin_no ")
        sb.AppendLine(" 	,nkin_recno ")
        sb.AppendLine(" 	,nkin_sortorder ")
        sb.AppendLine(" 	,nkin_kbn ")
        sb.AppendLine(" 	,sq_gak ")
        sb.AppendLine(" 	,sq_zeikbn ")
        sb.AppendLine(" 	,sq_zeigak ")
        sb.AppendLine(" 	,calc_kbn ")
        sb.AppendLine(" 	,calc_nkinno ")
        sb.AppendLine(" 	,calc_monthcnt ")
        sb.AppendLine(" 	,sqsaki_no ")
        sb.AppendLine(" 	,nkbn_yotei ")
        sb.AppendLine(" 	,sq_mmkbn ")
        sb.AppendLine(" 	,frstart_ymd ")
        sb.AppendLine(" 	,frend_ymd ")
        sb.AppendLine(" 	,frsq_gak ")
        sb.AppendLine(" 	,sqstart_ymd ")
        sb.AppendLine(" 	,sq_ptn ")
        sb.AppendLine(" 	,sq_interval ")
        sb.AppendLine(" 	,sq_nen ")
        sb.AppendLine(" 	,sq_tuki ")
        sb.AppendLine(" 	,zei_rit ")
        sb.AppendLine(" 	,biko ")
        sb.AppendLine(" 	,NEWID() AS nkin_guid ")
        sb.AppendLine(" 	,history ")
        sb.AppendLine(" 	,fr_kbn ")
        sb.AppendLine(" 	,frtekiyo_gak ")
        sb.AppendLine(" FROM ")
        sb.AppendLine(" ( ")
        sb.AppendLine(" 	SELECT * FROM kydata_nkin AS KYNKIN ")
        sb.AppendLine(" 	WHERE tuki_kbn = 2					/*翌月分*/ ")
        sb.AppendLine(" 	AND   nkin_no BETWEEN 1000 AND 1999	/*毎月*/ ")
        sb.AppendLine(" 	AND   NOT EXISTS ")
        sb.AppendLine(" 		( ")
        sb.AppendLine(" 			/*同一契約情報の当月分入金項目を抽出*/ ")
        sb.AppendLine(" 			SELECT * FROM kydata_nkin AS KYNKINSUB ")
        sb.AppendLine(" 			WHERE KYNKIN.ky_guid = KYNKINSUB.ky_guid ")
        sb.AppendLine(" 			AND   KYNKIN.ky_recno = KYNKINSUB.ky_recno ")
        sb.AppendLine(" 			AND   KYNKIN.nkin_no = KYNKINSUB.nkin_no ")
        sb.AppendLine(" 			AND   tuki_kbn = 1 ")
        sb.AppendLine(" 			AND   nkin_no BETWEEN 1000 AND 1999 ")
        sb.AppendLine(" 		) ")
        sb.AppendLine(" ) AS VW ")
        DBExec.Exec_NonQuery(sqlcnnv10, sb.ToString, cnt)

    End Sub

    ''' <summary>
    ''' 自社web連動に必要なデータを自動作成する
    ''' </summary>
    Public Sub Set_WmpTable(ByVal sqlcnnv10 As SqlConnection)

        '自社web連動時に以下のテーブル・フィールドにデータが存在しない場合、賃貸10で
        '一括掲載チェック等を行うとエラーが発生する
        'そのためプログラム内で自動作成する

        Dim sb As New System.Text.StringBuilder
        Dim cnt As Integer = 0

        '***********************************************************
        '自動作成するデータを取得
        '***********************************************************
        sb.Clear()
        sb.AppendLine(" SELECT ")
        sb.AppendLine(" 	 ROW_NUMBER()OVER(ORDER BY bk_no,hy_no) + (SELECT ISNULL(MAX(CONVERT(INT,id)),0) FROM wmp_hy_code_past_fk) AS id ")
        sb.AppendLine(" 	,bk_no AS bk_no ")
        sb.AppendLine(" 	,hy_no AS hy_no ")
        sb.AppendLine(" FROM hydata AS HY ")
        sb.AppendLine(" LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ")
        sb.AppendLine(" WHERE NOT EXISTS ")
        sb.AppendLine(" 	( ")
        sb.AppendLine(" 		SELECT * FROM wmp_hy_code_past_fk AS WMP_FK ")
        sb.AppendLine(" 		WHERE HY.wmp_id = WMP_FK.id ")
        sb.AppendLine(" 	) ")
        sb.AppendLine(" ORDER BY id ")
        Dim readtbl As New DataTable
        Dim rowcnt As Integer = DBExec.Exec_DataTable(sb.ToString, sqlcnnv10, readtbl)

        '***********************************************************
        'hydata.wmp_id(フィールド)の調整
        '***********************************************************
        sb.Clear()
        sb.AppendLine(" UPDATE hydata SET ")
        sb.AppendLine(" 	wmp_id = id ")
        sb.AppendLine(" FROM hydata AS HY ")
        sb.AppendLine(" LEFT JOIN ")
        sb.AppendLine(" 	( ")
        sb.AppendLine(" 		SELECT * FROM ")
        sb.AppendLine(" 		( ")
        sb.AppendLine(" 			SELECT ")
        sb.AppendLine(" 				 bk_no ")
        sb.AppendLine(" 				,hy_no ")
        sb.AppendLine(" 				,hy_guid ")
        sb.AppendLine(" 				,ROW_NUMBER()OVER(ORDER BY bk_no,hy_no) + (SELECT ISNULL(MAX(CONVERT(INT,id)),0) FROM wmp_hy_code_past_fk) AS id ")
        sb.AppendLine(" 			FROM hydata AS HY ")
        sb.AppendLine(" 			LEFT JOIN bkdata AS BK ON HY.bk_guid = BK.bk_guid ")
        sb.AppendLine(" 			WHERE NOT EXISTS ")
        sb.AppendLine(" 				( ")
        sb.AppendLine(" 					SELECT * FROM wmp_hy_code_past_fk AS WMP_FK ")
        sb.AppendLine(" 					WHERE HY.wmp_id = WMP_FK.id ")
        sb.AppendLine(" 				) ")
        sb.AppendLine(" 		) AS VW ")
        sb.AppendLine(" 	) AS VW ")
        sb.AppendLine(" ON HY.hy_guid = VW.hy_guid ")
        sb.AppendLine(" WHERE VW.hy_guid IS NOT NULL ")
        DBExec.Exec_NonQuery(sqlcnnv10, sb.ToString, cnt)

        '***********************************************************
        'wmp_hy_code_past_fkへの新規レコード作成
        '***********************************************************
        For cntii = 0 To rowcnt - 1

            Dim tmp_id As String = ""
            Dim tmp_bkno As String = ""
            Dim tmp_hyno As String = ""

            For cntjj = 0 To readtbl.Columns.Count - 1

                '項目名取得
                Dim fldname As String = readtbl.Columns(cntjj).ColumnName.Trim
                '登録値取得
                Dim fldvalue As String = Typ.ToStr(readtbl.Rows(cntii).Item(cntjj)).Trim

                '変数へ格納
                Select Case fldname
                    Case "id"
                        tmp_id = fldvalue.Trim
                    Case "bk_no"
                        tmp_bkno = fldvalue.Trim
                    Case "hy_no"
                        tmp_hyno = fldvalue.Trim
                End Select

            Next

            '既存データと照合し、同一データがある場合は部屋Noを変換する
            '※照合と変換処理は賃貸10ソースを引用
            Dim chghyno As String = getNotDuplicateHyCode(tmp_bkno, tmp_hyno, sqlcnnv10)

            'レコードを新規に作成する
            sb.Clear()
            sb.AppendLine(String.Format(" INSERT INTO wmp_hy_code_past_fk VALUES({0},{1},'{2}',{3},'{4}') ", tmp_id, tmp_bkno, chghyno, "GETDATE()", "NJC"))
            DBExec.Exec_NonQuery(sqlcnnv10, sb.ToString, cnt)

        Next

        '***********************************************************
        'hydata_wmpid(テーブル)の調整
        '***********************************************************
        sb.Clear()
        sb.AppendLine(" TRUNCATE TABLE hydata_wmpid ")
        DBExec.Exec_NonQuery(sqlcnnv10, sb.ToString, cnt)

        sb.Clear()
        sb.AppendLine(" INSERT INTO hydata_wmpid SELECT ISNULL(MAX(CONVERT(INT,id)),0) FROM wmp_hy_code_past_fk ")
        DBExec.Exec_NonQuery(sqlcnnv10, sb.ToString, cnt)

    End Sub
    ''' <summary>
    ''' 部屋Noをハッシュ化する処理用(賃貸10から流用)<br/>
    ''' </summary>
    Private Function LenB(ByVal stTarget As String) As Integer
        Return System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(stTarget)
    End Function
    ''' <summary>
    ''' 部屋Noをハッシュ化する処理用(賃貸10から流用して加工)<br/>
    ''' </summary>
    Private Function getNotDuplicateHyCode(bkNo As Integer, hyNo As String, ByVal sqlcnnv10 As SqlConnection) As String
        If (Typ.IsStrMissing(hyNo)) Then Throw New ArgumentException("hyNo")
        Dim result As Integer = 0
        Dim tryHyNo = hyNo
        If (LenB(tryHyNo) > 8) Then
            tryHyNo = Left(sha256_hash(tryHyNo), 8)
        End If
        While (True)
            Dim sb As New System.Text.StringBuilder
            sb.AppendLine(" SELECT COUNT(*) FROM wmp_hy_code_past_fk ")
            sb.AppendLine(String.Format(" WHERE bk_code = {0} ", bkNo))
            sb.AppendLine(String.Format(" AND   hy_code = '{0}' ", tryHyNo))
            Dim existdatacnt As Integer = Typ.ToInt(DBExec.Exec_Scalar(sb.ToString, sqlcnnv10))
            If (existdatacnt = 0) Then Exit While
            tryHyNo = Left(sha256_hash(tryHyNo), 8)
        End While
        Return tryHyNo
    End Function
    ''' <summary>
    ''' 部屋Noをハッシュ化する処理用(賃貸10から流用)<br/>
    ''' </summary>
    Private Function sha256_hash(value As String) As String
        Using hash = SHA256Managed.Create()
            Return String.Join("", hash.
                               ComputeHash(Encoding.UTF8.GetBytes(value)).
                               Select(Function(Item) Item.ToString("x2")))
        End Using
    End Function

End Module
