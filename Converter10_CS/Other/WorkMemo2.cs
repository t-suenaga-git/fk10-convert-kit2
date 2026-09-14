
// SELECT
// VW.*
// ,ISNULL(CVDB.CV用種別,'') AS CV用種別
// ,ISNULL(CVDB.CV用項目名,'') AS CV用項目名
// ,ISNULL(CVDB.CV用最小値,'') AS CV用最小値
// ,ISNULL(CVDB.CV用最大値,'') AS CV用最大値
// ,ISNULL(CVDB.CV用キーNo,'') AS CV用キーNo
// ,ISNULL(CVDB.CV用必須項目フラグ,'') AS CV用必須項目フラグ
// ,ISNULL(CVDB.CV備考,'') AS CV備考
// FROM
// (
// SELECT
// col.column_id AS [行No]
// ,CASE
// WHEN index_column_id IS NULL THEN ''
// ELSE '*'
// END AS [主キー]
// ,jpnname.value AS [種別]
// ,obj.NAME AS [TBL名]
// ,fldjpname.value AS [項目名]
// ,col.NAME AS [フィールド名]
// ,type_name(col.user_type_id) AS [属性]
// ,col.max_length AS [サイズ]
// ,CASE is_nullable
// WHEN '1' THEN '○'
// ELSE ''
// END AS [Null許容]
// ,obj.id
// /*
// ,fldcomment.value AS [説明]
// ,creater.value AS creator
// ,checker.value AS checker
// ,comment.value  AS [TBL説明]
// */
// FROM sys.sysobjects AS obj
// LEFT JOIN sys.extended_properties AS jpnname ON obj.id = jpnname.major_id
// AND jpnname.NAME = 'MS_Description'
// AND jpnname.minor_id = 0
// LEFT JOIN sys.columns AS col ON obj.id = col.object_id
// LEFT JOIN sys.extended_properties AS fldjpname ON col.object_id = fldjpname.major_id
// AND col.column_id = fldjpname.minor_id
// AND fldjpname.NAME = 'Jpfieldname'
// LEFT JOIN sys.indexes AS I ON col.object_id = I.object_id
// AND I.is_primary_key = 1
// LEFT JOIN sys.index_columns AS pkey ON I.object_id = pkey.object_id
// AND col.column_id = pkey.column_id
// AND I.index_id = pkey.index_id
// WHERE obj.xtype = 'U'
// ) AS VW
// LEFT JOIN cv_dbinfo AS CVDB
// ON  VW.TBL名 = CVDB.紐付用TBL名
// AND VW.フィールド名 = CVDB.紐付用フィールド名

// WHERE TBL名 = 'bkdata_szeniji'

// ORDER BY [TBL名],id,[行No]