-- pre_ テーブル作成クエリ(汎用コンバートキット2 に同梱)
-- 対応ツール: 汎用コンバートキット2 ver 1.3.0.0
-- 更新履歴:
--   ver 1.3.0.0: pre_契約情報 に、tuti_ymd(契約更新通知日)を追加
--   ver 1.2.0.0: pre_物件管理情報 に、hyteigaku_hiwariflg(部屋毎定額 - 管理手数料を比率計算する 0/1)、
--           hyteigaku_hiwarikbn(部屋毎定額 - 比率計算の基準 1:金額 2:日数 -1:未定義)を追加
--   ver 1.1.0.0: pre_家主情報、pre_仲介業者情報、pre_修繕業者情報、pre_ライフライン業者情報、pre_施設保守業者情報 に、
--           invoice_bango(インボイス登録番号)、tax_gykbn(課税区分 1:未確認 2:課税事業者 3:免税事業者)を追加
-- 注意: 実行すると、pre_ テーブルを削除して作り直すため、中のデータは消えます。
--       新しい列は NOT NULL DEFAULT '' です(新しい列を指定しない既存のINSERTも、そのまま動きます)。
PRINT N'pre_ テーブル作成: 汎用コンバートキット2 ver 1.3.0.0 対応';

IF OBJECT_ID(N'[dbo].[pre_物件鍵タイトルマスタ]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_物件鍵タイトルマスタ];
End

CREATE TABLE [dbo].[pre_物件鍵タイトルマスタ](
	[ROW_ID] [int] NOT NULL,
	[kagi_no] [varchar](max) NOT NULL,
	[kagi_name] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_部屋鍵タイトルマスタ]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_部屋鍵タイトルマスタ];
End

CREATE TABLE [dbo].[pre_部屋鍵タイトルマスタ](
	[ROW_ID] [int] NOT NULL,
	[kagi_no] [varchar](max) NOT NULL,
	[kagi_name] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_部屋共通セールスポイント情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_部屋共通セールスポイント情報];
End

CREATE TABLE [dbo].[pre_部屋共通セールスポイント情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[hy_no] [varchar](max) NOT NULL,
	[parts_no] [varchar](max) NOT NULL,
	[salespoint_parts] [varchar](max) NOT NULL
);


IF OBJECT_ID(N'[dbo].[pre_バス交通マスタ]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_バス交通マスタ];
End

CREATE TABLE [dbo].[pre_バス交通マスタ](
	[ROW_ID] [int] NOT NULL,
	[buskotu_no] [varchar](max) NOT NULL,
	[buskotu_name] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_バス停マスタ]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_バス停マスタ];
End

CREATE TABLE [dbo].[pre_バス停マスタ](
	[ROW_ID] [int] NOT NULL,
	[buskotu_no] [varchar](max) NOT NULL,
	[sortorder] [varchar](max) NOT NULL,
	[buskotu_name] [varchar](max) NOT NULL,
	[busstop_name] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_学校区マスタ]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_学校区マスタ];
End

CREATE TABLE [dbo].[pre_学校区マスタ](
	[ROW_ID] [int] NOT NULL,
	[ken_no][varchar](max) NOT NULL,
	[si_no][varchar](max) NOT NULL,
	[no][varchar](max) NOT NULL,
	[add_cyo][varchar](max) NOT NULL,
	[add_cyome][varchar](max) NOT NULL,
	[add_banti][varchar](max) NOT NULL,
	[syogaku_name][varchar](max) NOT NULL,
	[cyugaku_name][varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_特約マスタ]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_特約マスタ];
End

CREATE TABLE [dbo].[pre_特約マスタ](
	[ROW_ID] [int] NOT NULL,
	[tokuyaku_grpno][varchar](max) NOT NULL,
	[tokuyaku_no][varchar](max) NOT NULL,
	[tokuyaku_title][varchar](max) NOT NULL,
	[tokuyaku_template][varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_エリアマスタ]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_エリアマスタ];
End

CREATE TABLE [dbo].[pre_エリアマスタ](
	[ROW_ID] [int] NOT NULL,
	[area_no] [varchar](max) NOT NULL,
	[area_name] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_保険種類マスタ]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_保険種類マスタ];
End

CREATE TABLE [dbo].[pre_保険種類マスタ](
	[ROW_ID] [int] NOT NULL,
	[hoken_ruino] [varchar](max) NOT NULL,
	[hoken_ruiname] [varchar](max) NOT NULL,
	[hoken_ruikana] [varchar](max) NOT NULL,
	[biko_kihon] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_自社情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_自社情報];
End

CREATE TABLE [dbo].[pre_自社情報](
	[ROW_ID] [int] NOT NULL,
	[jisya_no] [varchar](max) NOT NULL,
	[jisya_name] [varchar](max) NOT NULL,
	[jisya_kana] [varchar](max) NOT NULL,
	[jisya_name2] [varchar](max) NOT NULL,
	[jisya_name3] [varchar](max) NOT NULL,
	[daihyo_yakusyoku] [varchar](max) NOT NULL,
	[daihyo_name] [varchar](max) NOT NULL,
	[post_code] [varchar](max) NOT NULL,
	[addr1] [varchar](max) NOT NULL,
	[addr2] [varchar](max) NOT NULL,
	[tel1] [varchar](max) NOT NULL,
	[tel2] [varchar](max) NOT NULL,
	[mobiletel1] [varchar](max) NOT NULL,
	[mobiletel2] [varchar](max) NOT NULL,
	[fax] [varchar](max) NOT NULL,
	[mail] [varchar](max) NOT NULL,
	[mobilemail] [varchar](max) NOT NULL,
	[url] [varchar](max) NOT NULL,
	[biko_kihon] [varchar](max) NOT NULL,
	[menkyo_bango] [varchar](max) NOT NULL,
	[menkyo_ymd] [varchar](max) NOT NULL,
	[syunin_bango] [varchar](max) NOT NULL,
	[syunin_name] [varchar](max) NOT NULL,
	[kanrikyokai_bango] [varchar](max) NOT NULL,
	[keieikanrisi_bango] [varchar](max) NOT NULL,
	[keieikanrisi_name] [varchar](max) NOT NULL,
	[kanrigy_bango] [varchar](max) NOT NULL,
	[kanyu_dantai1] [varchar](max) NOT NULL,
	[kanyu_dantai2] [varchar](max) NOT NULL,
	[kanyu_dantai3] [varchar](max) NOT NULL,
	[memo1] [varchar](max) NOT NULL,
	[memo2] [varchar](max) NOT NULL,
	[memo3] [varchar](max) NOT NULL,
	[memo4] [varchar](max) NOT NULL,
	[memo5] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_自社担当者情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_自社担当者情報];
End

CREATE TABLE [dbo].[pre_自社担当者情報](
	[ROW_ID] [int] NOT NULL,
	[logonuser_no] [varchar](max) NOT NULL,
	[logonuser_name] [varchar](max) NOT NULL,
	[logonuser_kana] [varchar](max) NOT NULL,
	[tel1] [varchar](max) NOT NULL,
	[tel2] [varchar](max) NOT NULL,
	[mobiletel1] [varchar](max) NOT NULL,
	[fax1] [varchar](max) NOT NULL,
	[mailaddress1] [varchar](max) NOT NULL,
	[mobilemailaddress1] [varchar](max) NOT NULL,
	[takkenmenkyo] [varchar](max) NOT NULL,
	[biko] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_自社口座情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_自社口座情報];
End

CREATE TABLE [dbo].[pre_自社口座情報](
	[ROW_ID] [int] NOT NULL,
	[jisya_no] [varchar](max) NOT NULL,
	[jisya_kozano] [varchar](max) NOT NULL,
	[kinyu_no] [varchar](max) NOT NULL,
	[kinyu_tenno] [varchar](max) NOT NULL,
	[koza_syubetu] [varchar](max) NOT NULL,
	[koza_bango] [varchar](max) NOT NULL,
	[koza_meigi] [varchar](max) NOT NULL,
	[koza_meigikana] [varchar](max) NOT NULL,
	[yucyokoza_kigo1] [varchar](max) NOT NULL,
	[yucyokoza_kigo2] [varchar](max) NOT NULL,
	[yucyokoza_bango] [varchar](max) NOT NULL,
	[biko_koza] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_振込依頼人情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_振込依頼人情報];
End

CREATE TABLE [dbo].[pre_振込依頼人情報](
	[ROW_ID] [int] NOT NULL,
	[sgfirai_no] [varchar](max) NOT NULL,
	[sgfirai_name] [varchar](max) NOT NULL,
	[sgfirainin_code] [varchar](max) NOT NULL,
	[sgfirainin_kana] [varchar](max) NOT NULL,
	[biko] [varchar](max) NOT NULL,
	[jisya_no] [varchar](max) NOT NULL,
	[jisya_kozano] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_口座振替情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_口座振替情報];
End

CREATE TABLE [dbo].[pre_口座振替情報](
	[ROW_ID] [int] NOT NULL,
	[fkae_no] [varchar](max) NOT NULL,
	[fkae_name] [varchar](max) NOT NULL,
	[fkae_kana] [varchar](max) NOT NULL,
	[fkae_fb_fkomiraino] [varchar](max) NOT NULL,
	[fkae_fb_fkomiraikana] [varchar](max) NOT NULL,
	[kamei_no] [varchar](max) NOT NULL,
	[hikiotosibi] [varchar](max) NOT NULL,
	[tesu_gak] [varchar](max) NOT NULL,
	[biko] [varchar](max) NOT NULL,
	[jisya_no] [varchar](max) NOT NULL,
	[fkae_fb_nkinukekozano] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_家主情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_家主情報];
End

CREATE TABLE [dbo].[pre_家主情報](
	[ROW_ID] [int] NOT NULL,
	[ow_no] [varchar](max) NOT NULL,
	[kojinhojin_flg] [varchar](max) NOT NULL,
	[ow_name] [varchar](max) NOT NULL,
	[ow_kana] [varchar](max) NOT NULL,
	[keisyo] [varchar](max) NOT NULL,
	[post_code] [varchar](max) NOT NULL,
	[addr1] [varchar](max) NOT NULL,
	[addr2] [varchar](max) NOT NULL,
	[toukiaddr1] [varchar](max) NOT NULL,
	[toukiaddr2] [varchar](max) NOT NULL,
	[tel1] [varchar](max) NOT NULL,
	[tel2] [varchar](max) NOT NULL,
	[fax] [varchar](max) NOT NULL,
	[mobiletel1] [varchar](max) NOT NULL,
	[mobiletel2] [varchar](max) NOT NULL,
	[mail] [varchar](max) NOT NULL,
	[mobilemail] [varchar](max) NOT NULL,
	[birthday] [varchar](max) NOT NULL,
	[nensyu] [varchar](max) NOT NULL,
	[biko_kihon] [varchar](max) NOT NULL,
	[gender] [varchar](max) NOT NULL,
	[honseki] [varchar](max) NOT NULL,
	[kinmu_name] [varchar](max) NOT NULL,
	[kinmu_kana] [varchar](max) NOT NULL,
	[kinmu_postcode] [varchar](max) NOT NULL,
	[kinmu_addr1] [varchar](max) NOT NULL,
	[kinmu_addr2] [varchar](max) NOT NULL,
	[kinmu_tel1] [varchar](max) NOT NULL,
	[kinmu_tel2] [varchar](max) NOT NULL,
	[kinmu_fax] [varchar](max) NOT NULL,
	[kinmu_gyosyu] [varchar](max) NOT NULL,
	[kinmu_busyo] [varchar](max) NOT NULL,
	[kinmu_nyuryokuym] [varchar](max) NOT NULL,
	[kinmu_nyusyaym] [varchar](max) NOT NULL,
	[url] [varchar](max) NOT NULL,
	[gyosyu] [varchar](max) NOT NULL,
	[daihyo_name] [varchar](max) NOT NULL,
	[daihyo_kana] [varchar](max) NOT NULL,
	[daihyo_yakusyoku] [varchar](max) NOT NULL,
	[daihyo_atenaflg] [varchar](max) NOT NULL,
	[tanto_name] [varchar](max) NOT NULL,
	[tanto_kana] [varchar](max) NOT NULL,
	[tanto_busyo] [varchar](max) NOT NULL,
	[tanto_atenaflg] [varchar](max) NOT NULL,
	[sihonkin] [varchar](max) NOT NULL,
	[jugyosu] [varchar](max) NOT NULL,
	[nyuryoku_ym] [varchar](max) NOT NULL,
	[torihikisaki] [varchar](max) NOT NULL,
	[renraku_name] [varchar](max) NOT NULL,
	[renraku_kana] [varchar](max) NOT NULL,
	[renraku_keisyo] [varchar](max) NOT NULL,
	[renraku_postcode] [varchar](max) NOT NULL,
	[renraku_addr1] [varchar](max) NOT NULL,
	[renraku_addr2] [varchar](max) NOT NULL,
	[renraku_tel1] [varchar](max) NOT NULL,
	[renraku_tel2] [varchar](max) NOT NULL,
	[renraku_fax] [varchar](max) NOT NULL,
	[renraku_mobiletel1] [varchar](max) NOT NULL,
	[renraku_mobiletel2] [varchar](max) NOT NULL,
	[renraku_aidagara] [varchar](max) NOT NULL,
	[biko_renraku] [varchar](max) NOT NULL,
	[sofu_kbn] [varchar](max) NOT NULL,
	[sofu_name] [varchar](max) NOT NULL,
	[sofu_kana] [varchar](max) NOT NULL,
	[sofu_keisyo] [varchar](max) NOT NULL,
	[sofu_postcode] [varchar](max) NOT NULL,
	[sofu_addr1] [varchar](max) NOT NULL,
	[sofu_addr2] [varchar](max) NOT NULL,
	[sofu_tel1] [varchar](max) NOT NULL,
	[sofu_tel2] [varchar](max) NOT NULL,
	[sofu_fax] [varchar](max) NOT NULL,
	[biko_sofu] [varchar](max) NOT NULL,
	[memo1] [varchar](max) NOT NULL,
	[memo2] [varchar](max) NOT NULL,
	[memo3] [varchar](max) NOT NULL,
	[memo4] [varchar](max) NOT NULL,
	[memo5] [varchar](max) NOT NULL,
	[memo6] [varchar](max) NOT NULL,
	[memo7] [varchar](max) NOT NULL,
	[memo8] [varchar](max) NOT NULL,
	[memo9] [varchar](max) NOT NULL,
	[memo10] [varchar](max) NOT NULL,
	[invoice_bango] [varchar](max) NOT NULL CONSTRAINT [DF_pre_家主情報_invoice_bango] DEFAULT '',
	[tax_gykbn] [varchar](max) NOT NULL CONSTRAINT [DF_pre_家主情報_tax_gykbn] DEFAULT ''
);

IF OBJECT_ID(N'[dbo].[pre_家主口座情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_家主口座情報];
End

CREATE TABLE [dbo].[pre_家主口座情報](
	[ROW_ID] [int] NOT NULL,
	[ow_no] [varchar](max) NOT NULL,
	[ow_kozano] [varchar](max) NOT NULL,
	[kinyu_no] [varchar](max) NOT NULL,
	[kinyu_tenno] [varchar](max) NOT NULL,
	[koza_syubetu] [varchar](max) NOT NULL,
	[koza_bango] [varchar](max) NOT NULL,
	[koza_meigi] [varchar](max) NOT NULL,
	[koza_meigikana] [varchar](max) NOT NULL,
	[yucyokoza_kigo1] [varchar](max) NOT NULL,
	[yucyokoza_kigo2] [varchar](max) NOT NULL,
	[yucyokoza_bango] [varchar](max) NOT NULL,
	[biko_koza] [varchar](max) NOT NULL,
	[biko_furikomi] [varchar](max) NOT NULL,
	[biko_sgfirai] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_家賃入金口座情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_家賃入金口座情報];
End

CREATE TABLE [dbo].[pre_家賃入金口座情報](
	[ROW_ID] [int] NOT NULL,
	[yatin_kozano] [varchar](max) NOT NULL,
	[yatin_kozaname] [varchar](max) NOT NULL,
	[yatin_kozakana] [varchar](max) NOT NULL,
	[jisya_no] [varchar](max) NOT NULL,
	[jisya_kozano] [varchar](max) NOT NULL,
	[ow_no] [varchar](max) NOT NULL,
	[ow_kozano] [varchar](max) NOT NULL,
	[yatin_kozabiko] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_契約者情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_契約者情報];
End

CREATE TABLE [dbo].[pre_契約者情報](
	[ROW_ID] [int] NOT NULL,
	[kys_no] [varchar](max) NOT NULL,
	[kys_name] [varchar](max) NOT NULL,
	[kys_kana] [varchar](max) NOT NULL,
	[kojinhojin_flg] [varchar](max) NOT NULL,
	[keisyo] [varchar](max) NOT NULL,
	[post_code] [varchar](max) NOT NULL,
	[addr1] [varchar](max) NOT NULL,
	[addr2] [varchar](max) NOT NULL,
	[tel1] [varchar](max) NOT NULL,
	[tel2] [varchar](max) NOT NULL,
	[fax] [varchar](max) NOT NULL,
	[mobiletel1] [varchar](max) NOT NULL,
	[mobiletel2] [varchar](max) NOT NULL,
	[mail] [varchar](max) NOT NULL,
	[mobilemail] [varchar](max) NOT NULL,
	[birthday] [varchar](max) NOT NULL,
	[nensyu] [varchar](max) NOT NULL,
	[biko_kihon] [varchar](max) NOT NULL,
	[gender] [varchar](max) NOT NULL,
	[honseki] [varchar](max) NOT NULL,
	[kinmu_name] [varchar](max) NOT NULL,
	[kinmu_kana] [varchar](max) NOT NULL,
	[kinmu_postcode] [varchar](max) NOT NULL,
	[kinmu_addr1] [varchar](max) NOT NULL,
	[kinmu_addr2] [varchar](max) NOT NULL,
	[kinmu_tel1] [varchar](max) NOT NULL,
	[kinmu_tel2] [varchar](max) NOT NULL,
	[kinmu_fax] [varchar](max) NOT NULL,
	[kinmu_gyosyu] [varchar](max) NOT NULL,
	[kinmu_busyo] [varchar](max) NOT NULL,
	[url] [varchar](max) NOT NULL,
	[gyosyu] [varchar](max) NOT NULL,
	[daihyo_name] [varchar](max) NOT NULL,
	[daihyo_kana] [varchar](max) NOT NULL,
	[daihyo_yakusyoku] [varchar](max) NOT NULL,
	[tanto_name] [varchar](max) NOT NULL,
	[tanto_kana] [varchar](max) NOT NULL,
	[tanto_busyo] [varchar](max) NOT NULL,
	[tanto_yakusyoku] [varchar](max) NOT NULL,
	[sihonkin] [varchar](max) NOT NULL,
	[jugyosu] [varchar](max) NOT NULL,
	[nyuryoku_ym] [varchar](max) NOT NULL,
	[torihikisaki] [varchar](max) NOT NULL,
	[renraku_name] [varchar](max) NOT NULL,
	[renraku_kana] [varchar](max) NOT NULL,
	[renraku_keisyo] [varchar](max) NOT NULL,
	[renraku_postcode] [varchar](max) NOT NULL,
	[renraku_addr1] [varchar](max) NOT NULL,
	[renraku_addr2] [varchar](max) NOT NULL,
	[renraku_tel1] [varchar](max) NOT NULL,
	[renraku_tel2] [varchar](max) NOT NULL,
	[renraku_fax] [varchar](max) NOT NULL,
	[renraku_mobiletel1] [varchar](max) NOT NULL,
	[renraku_mobiletel2] [varchar](max) NOT NULL,
	[renraku_aidagara] [varchar](max) NOT NULL,
	[biko_renraku] [varchar](max) NOT NULL,
	[sofu_kbn] [varchar](max) NOT NULL,
	[sofu_name] [varchar](max) NOT NULL,
	[sofu_kana] [varchar](max) NOT NULL,
	[sofu_keisyo] [varchar](max) NOT NULL,
	[sofu_postcode] [varchar](max) NOT NULL,
	[sofu_addr1] [varchar](max) NOT NULL,
	[sofu_addr2] [varchar](max) NOT NULL,
	[sofu_tel1] [varchar](max) NOT NULL,
	[sofu_tel2] [varchar](max) NOT NULL,
	[sofu_fax] [varchar](max) NOT NULL,
	[biko_sofu] [varchar](max) NOT NULL,
	[fkom_syogomoji1] [varchar](max) NOT NULL,
	[fkom_syogomoji2] [varchar](max) NOT NULL,
	[fkom_syogomoji3] [varchar](max) NOT NULL,
	[fkom_syogomoji4] [varchar](max) NOT NULL,
	[fkom_syogomoji5] [varchar](max) NOT NULL,
	[fkom_syogomoji6] [varchar](max) NOT NULL,
	[fkom_syogomoji7] [varchar](max) NOT NULL,
	[fkom_syogomoji8] [varchar](max) NOT NULL,
	[fkom_syogomoji9] [varchar](max) NOT NULL,
	[fkom_syogomoji10] [varchar](max) NOT NULL,
	[memo1] [varchar](max) NOT NULL,
	[memo2] [varchar](max) NOT NULL,
	[memo3] [varchar](max) NOT NULL,
	[memo4] [varchar](max) NOT NULL,
	[memo5] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_契約者口座情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_契約者口座情報];
End

CREATE TABLE [dbo].[pre_契約者口座情報](
	[ROW_ID] [int] NOT NULL,
	[kys_no] [varchar](max) NOT NULL,
	[kys_kozano] [varchar](max) NOT NULL,
	[kinyu_no] [varchar](max) NOT NULL,
	[kinyu_tenno] [varchar](max) NOT NULL,
	[koza_syubetu] [varchar](max) NOT NULL,
	[koza_bango] [varchar](max) NOT NULL,
	[koza_meigi] [varchar](max) NOT NULL,
	[koza_meigikana] [varchar](max) NOT NULL,
	[yucyokoza_kigo1] [varchar](max) NOT NULL,
	[yucyokoza_kigo2] [varchar](max) NOT NULL,
	[yucyokoza_bango] [varchar](max) NOT NULL,
	[fkae_no] [varchar](max) NOT NULL,
	[fkae_tesugak] [varchar](max) NOT NULL,
	[fkae_kysbango] [varchar](max) NOT NULL,
	[biko_koza] [varchar](max) NOT NULL,
	[koza_yucyoflg] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_契約者保証人情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_契約者保証人情報];
End

CREATE TABLE [dbo].[pre_契約者保証人情報](
	[ROW_ID] [int] NOT NULL,
	[kys_no] [varchar](max) NOT NULL,
	[hosyo_no] [varchar](max) NOT NULL,
	[hosyo_name] [varchar](max) NOT NULL,
	[hosyo_kana] [varchar](max) NOT NULL,
	[keisyo] [varchar](max) NOT NULL,
	[post_code] [varchar](max) NOT NULL,
	[addr1] [varchar](max) NOT NULL,
	[addr2] [varchar](max) NOT NULL,
	[tel1] [varchar](max) NOT NULL,
	[tel2] [varchar](max) NOT NULL,
	[fax] [varchar](max) NOT NULL,
	[mobiletel1] [varchar](max) NOT NULL,
	[mobiletel2] [varchar](max) NOT NULL,
	[birthday] [varchar](max) NOT NULL,
	[nensyu] [varchar](max) NOT NULL,
	[biko_hosyo] [varchar](max) NOT NULL,
	[aidagara] [varchar](max) NOT NULL,
	[kinmu_name] [varchar](max) NOT NULL,
	[kinmu_kana] [varchar](max) NOT NULL,
	[kinmu_postcode] [varchar](max) NOT NULL,
	[kinmu_addr1] [varchar](max) NOT NULL,
	[kinmu_addr2] [varchar](max) NOT NULL,
	[kinmu_tel1] [varchar](max) NOT NULL,
	[kinmu_tel2] [varchar](max) NOT NULL,
	[kinmu_fax] [varchar](max) NOT NULL,
	[kinmu_gyosyu] [varchar](max) NOT NULL,
	[kinmu_busyo] [varchar](max) NOT NULL,
	[kinmu_nyuryokuym] [varchar](max) NOT NULL,
	[biko_kinmu] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_仲介業者情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_仲介業者情報];
End

CREATE TABLE [dbo].[pre_仲介業者情報](
	[ROW_ID] [int] NOT NULL,
	[gy_fudono] [varchar](max) NOT NULL,
	[gy_fudoname] [varchar](max) NOT NULL,
	[gy_fudokana] [varchar](max) NOT NULL,
	[post_code] [varchar](max) NOT NULL,
	[addr1] [varchar](max) NOT NULL,
	[addr2] [varchar](max) NOT NULL,
	[tel1] [varchar](max) NOT NULL,
	[tel2] [varchar](max) NOT NULL,
	[Fax] [varchar](max) NOT NULL,
	[mobiletel1] [varchar](max) NOT NULL,
	[mobiletel2] [varchar](max) NOT NULL,
	[daihyo_yakusyoku] [varchar](max) NOT NULL,
	[daihyo_name] [varchar](max) NOT NULL,
	[tanto_busyoyakusyoku] [varchar](max) NOT NULL,
	[tanto_name] [varchar](max) NOT NULL,
	[tanto_kana] [varchar](max) NOT NULL,
	[mail] [varchar](max) NOT NULL,
	[mobilemail] [varchar](max) NOT NULL,
	[url] [varchar](max) NOT NULL,
	[biko_kihon] [varchar](max) NOT NULL,
	[sofu_hakkofurikomi] [varchar](max) NOT NULL,
	[menkyo_bango] [varchar](max) NOT NULL,
	[menkyo_ymd] [varchar](max) NOT NULL,
	[syunin_bango] [varchar](max) NOT NULL,
	[syunin_name] [varchar](max) NOT NULL,
	[kanrikyokai_bango] [varchar](max) NOT NULL,
	[keieikanrisi_bango] [varchar](max) NOT NULL,
	[keieikanrisi_name] [varchar](max) NOT NULL,
	[kanrigy_bango] [varchar](max) NOT NULL,
	[kinyu_no] [varchar](max) NOT NULL,
	[kinyu_tenno] [varchar](max) NOT NULL,
	[koza_syubetu] [varchar](max) NOT NULL,
	[koza_bango] [varchar](max) NOT NULL,
	[koza_meigi] [varchar](max) NOT NULL,
	[koza_meigikana] [varchar](max) NOT NULL,
	[yucyokoza_kigo1] [varchar](max) NOT NULL,
	[yucyokoza_kigo2] [varchar](max) NOT NULL,
	[yucyokoza_bango] [varchar](max) NOT NULL,
	[memo1] [varchar](max) NOT NULL,
	[memo2] [varchar](max) NOT NULL,
	[memo3] [varchar](max) NOT NULL,
	[memo4] [varchar](max) NOT NULL,
	[memo5] [varchar](max) NOT NULL,
	[invoice_bango] [varchar](max) NOT NULL CONSTRAINT [DF_pre_仲介業者情報_invoice_bango] DEFAULT '',
	[tax_gykbn] [varchar](max) NOT NULL CONSTRAINT [DF_pre_仲介業者情報_tax_gykbn] DEFAULT ''
);

IF OBJECT_ID(N'[dbo].[pre_修繕業者情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_修繕業者情報];
End

CREATE TABLE [dbo].[pre_修繕業者情報](
	[ROW_ID] [int] NOT NULL,
	[gy_szenno] [varchar](max) NOT NULL,
	[gy_szenname] [varchar](max) NOT NULL,
	[gy_szenkana] [varchar](max) NOT NULL,
	[post_code] [varchar](max) NOT NULL,
	[addr1] [varchar](max) NOT NULL,
	[addr2] [varchar](max) NOT NULL,
	[tel1] [varchar](max) NOT NULL,
	[tel2] [varchar](max) NOT NULL,
	[Fax] [varchar](max) NOT NULL,
	[mobiletel1] [varchar](max) NOT NULL,
	[mobiletel2] [varchar](max) NOT NULL,
	[daihyo_yakusyoku] [varchar](max) NOT NULL,
	[daihyo_name] [varchar](max) NOT NULL,
	[tanto_busyoyakusyoku] [varchar](max) NOT NULL,
	[tanto_name] [varchar](max) NOT NULL,
	[tanto_kana] [varchar](max) NOT NULL,
	[mail] [varchar](max) NOT NULL,
	[mobilemail] [varchar](max) NOT NULL,
	[url] [varchar](max) NOT NULL,
	[biko_kihon] [varchar](max) NOT NULL,
	[kinyu_no] [varchar](max) NOT NULL,
	[kinyu_tenno] [varchar](max) NOT NULL,
	[koza_syubetu] [varchar](max) NOT NULL,
	[koza_bango] [varchar](max) NOT NULL,
	[koza_meigi] [varchar](max) NOT NULL,
	[koza_meigikana] [varchar](max) NOT NULL,
	[yucyokoza_kigo1] [varchar](max) NOT NULL,
	[yucyokoza_kigo2] [varchar](max) NOT NULL,
	[yucyokoza_bango] [varchar](max) NOT NULL,
	[biko_koza] [varchar](max) NOT NULL,
	[biko_furikomi] [varchar](max) NOT NULL,
	[biko_sgfirai] [varchar](max) NOT NULL,
	[memo1] [varchar](max) NOT NULL,
	[memo2] [varchar](max) NOT NULL,
	[memo3] [varchar](max) NOT NULL,
	[memo4] [varchar](max) NOT NULL,
	[memo5] [varchar](max) NOT NULL,
	[invoice_bango] [varchar](max) NOT NULL CONSTRAINT [DF_pre_修繕業者情報_invoice_bango] DEFAULT '',
	[tax_gykbn] [varchar](max) NOT NULL CONSTRAINT [DF_pre_修繕業者情報_tax_gykbn] DEFAULT ''
);

IF OBJECT_ID(N'[dbo].[pre_ライフライン業者情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_ライフライン業者情報];
End

CREATE TABLE [dbo].[pre_ライフライン業者情報](
	[ROW_ID] [int] NOT NULL,
	[gy_lifelineno] [varchar](max) NOT NULL,
	[gy_lifelinename] [varchar](max) NOT NULL,
	[gy_lifelinekana] [varchar](max) NOT NULL,
	[eigyosyo] [varchar](max) NOT NULL,
	[post_code] [varchar](max) NOT NULL,
	[addr1] [varchar](max) NOT NULL,
	[addr2] [varchar](max) NOT NULL,
	[tel1] [varchar](max) NOT NULL,
	[tel2] [varchar](max) NOT NULL,
	[Fax] [varchar](max) NOT NULL,
	[mobiletel1] [varchar](max) NOT NULL,
	[mobiletel2] [varchar](max) NOT NULL,
	[biko_kihon] [varchar](max) NOT NULL,
	[lifeline_denkiflg] [varchar](max) NOT NULL,
	[lifeline_josuidoflg] [varchar](max) NOT NULL,
	[lifeline_gasflg] [varchar](max) NOT NULL,
	[lifeline_toyuflg] [varchar](max) NOT NULL,
	[lifeline_other1flg] [varchar](max) NOT NULL,
	[lifeline_haisuiflg] [varchar](max) NOT NULL,
	[invoice_bango] [varchar](max) NOT NULL CONSTRAINT [DF_pre_ライフライン業者情報_invoice_bango] DEFAULT '',
	[tax_gykbn] [varchar](max) NOT NULL CONSTRAINT [DF_pre_ライフライン業者情報_tax_gykbn] DEFAULT ''
);

IF OBJECT_ID(N'[dbo].[pre_保険業者情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_保険業者情報];
End

CREATE TABLE [dbo].[pre_保険業者情報](
	[ROW_ID] [int] NOT NULL,
	[gy_hokenno] [varchar](max) NOT NULL,
	[gy_hokenname] [varchar](max) NOT NULL,
	[gy_hokenkana] [varchar](max) NOT NULL,
	[post_code] [varchar](max) NOT NULL,
	[addr1] [varchar](max) NOT NULL,
	[addr2] [varchar](max) NOT NULL,
	[tel1] [varchar](max) NOT NULL,
	[tel2] [varchar](max) NOT NULL,
	[Fax] [varchar](max) NOT NULL,
	[mobiletel1] [varchar](max) NOT NULL,
	[mobiletel2] [varchar](max) NOT NULL,
	[tanto_busyoyakusyoku] [varchar](max) NOT NULL,
	[tanto_name] [varchar](max) NOT NULL,
	[tanto_kana] [varchar](max) NOT NULL,
	[mail] [varchar](max) NOT NULL,
	[mobilemail] [varchar](max) NOT NULL,
	[url] [varchar](max) NOT NULL,
	[biko_kihon] [varchar](max) NOT NULL,
	[kinyu_no] [varchar](max) NOT NULL,
	[kinyu_tenno] [varchar](max) NOT NULL,
	[koza_syubetu] [varchar](max) NOT NULL,
	[koza_bango] [varchar](max) NOT NULL,
	[koza_meigi] [varchar](max) NOT NULL,
	[koza_meigikana] [varchar](max) NOT NULL,
	[memo1] [varchar](max) NOT NULL,
	[memo2] [varchar](max) NOT NULL,
	[memo3] [varchar](max) NOT NULL,
	[memo4] [varchar](max) NOT NULL,
	[memo5] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_家賃保証業者情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_家賃保証業者情報];
End

CREATE TABLE [dbo].[pre_家賃保証業者情報](
	[ROW_ID] [int] NOT NULL,
	[gy_hosyono] [varchar](max) NOT NULL,
	[gy_hosyoname] [varchar](max) NOT NULL,
	[gy_hosyokana] [varchar](max) NOT NULL,
	[post_code] [varchar](max) NOT NULL,
	[addr1] [varchar](max) NOT NULL,
	[addr2] [varchar](max) NOT NULL,
	[tel1] [varchar](max) NOT NULL,
	[tel2] [varchar](max) NOT NULL,
	[Fax] [varchar](max) NOT NULL,
	[mobiletel1] [varchar](max) NOT NULL,
	[mobiletel2] [varchar](max) NOT NULL,
	[tanto_busyoyakusyoku] [varchar](max) NOT NULL,
	[tanto_name] [varchar](max) NOT NULL,
	[tanto_kana] [varchar](max) NOT NULL,
	[mail] [varchar](max) NOT NULL,
	[mobilemail] [varchar](max) NOT NULL,
	[url] [varchar](max) NOT NULL,
	[biko_kihon] [varchar](max) NOT NULL,
	[memo1] [varchar](max) NOT NULL,
	[memo2] [varchar](max) NOT NULL,
	[memo3] [varchar](max) NOT NULL,
	[memo4] [varchar](max) NOT NULL,
	[memo5] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_施設保守業者情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_施設保守業者情報];
End

CREATE TABLE [dbo].[pre_施設保守業者情報](
	[ROW_ID] [int] NOT NULL,
	[gy_sisetuno] [varchar](max) NOT NULL,
	[gy_sisetuname] [varchar](max) NOT NULL,
	[gy_sisetukana] [varchar](max) NOT NULL,
	[eigyosyo] [varchar](max) NOT NULL,
	[post_code] [varchar](max) NOT NULL,
	[addr1] [varchar](max) NOT NULL,
	[addr2] [varchar](max) NOT NULL,
	[tel1] [varchar](max) NOT NULL,
	[tel2] [varchar](max) NOT NULL,
	[Fax] [varchar](max) NOT NULL,
	[mobiletel1] [varchar](max) NOT NULL,
	[mobiletel2] [varchar](max) NOT NULL,
	[tanto_busyoyakusyoku] [varchar](max) NOT NULL,
	[tanto_name] [varchar](max) NOT NULL,
	[tanto_kana] [varchar](max) NOT NULL,
	[mail] [varchar](max) NOT NULL,
	[mobilemail] [varchar](max) NOT NULL,
	[url] [varchar](max) NOT NULL,
	[biko_kihon] [varchar](max) NOT NULL,
	[invoice_bango] [varchar](max) NOT NULL CONSTRAINT [DF_pre_施設保守業者情報_invoice_bango] DEFAULT '',
	[tax_gykbn] [varchar](max) NOT NULL CONSTRAINT [DF_pre_施設保守業者情報_tax_gykbn] DEFAULT ''
);

IF OBJECT_ID(N'[dbo].[pre_施工業者情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_施工業者情報];
End

CREATE TABLE [dbo].[pre_施工業者情報](
	[ROW_ID] [int] NOT NULL,
	[gy_sekono] [varchar](max) NOT NULL,
	[gy_sekoname] [varchar](max) NOT NULL,
	[gy_sekokana] [varchar](max) NOT NULL,
	[eigyosyo] [varchar](max) NOT NULL,
	[post_code] [varchar](max) NOT NULL,
	[addr1] [varchar](max) NOT NULL,
	[addr2] [varchar](max) NOT NULL,
	[tel1] [varchar](max) NOT NULL,
	[tel2] [varchar](max) NOT NULL,
	[Fax] [varchar](max) NOT NULL,
	[mobiletel1] [varchar](max) NOT NULL,
	[mobiletel2] [varchar](max) NOT NULL,
	[tanto_busyoyakusyoku] [varchar](max) NOT NULL,
	[tanto_name] [varchar](max) NOT NULL,
	[tanto_kana] [varchar](max) NOT NULL,
	[mail] [varchar](max) NOT NULL,
	[mobilemail] [varchar](max) NOT NULL,
	[url] [varchar](max) NOT NULL,
	[biko_kihon] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_物件情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_物件情報];
End

CREATE TABLE [dbo].[pre_物件情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[bk_name] [varchar](max) NOT NULL,
	[bk_kana] [varchar](max) NOT NULL,
	[tatemono_sikibetu] [varchar](max) NOT NULL,
	[post_code] [varchar](max) NOT NULL,
	[addr_kenno] [varchar](max) NOT NULL,
	[addr_sino] [varchar](max) NOT NULL,
	[addr_cyo] [varchar](max) NOT NULL,
	[addr_cyome] [varchar](max) NOT NULL,
	[addr_banti] [varchar](max) NOT NULL,
	[addr_etc] [varchar](max) NOT NULL,
	[bk_ruinokbn] [varchar](max) NOT NULL,
	[gps_wgsido] [varchar](max) NOT NULL,
	[gps_wgskeido] [varchar](max) NOT NULL,
	[syunko_ymd] [varchar](max) NOT NULL,
	[kaidate] [varchar](max) NOT NULL,
	[tika] [varchar](max) NOT NULL,
	[elevator_flg] [varchar](max) NOT NULL,
	[elevator_number] [varchar](max) NOT NULL,
	[kozo_nokbn] [varchar](max) NOT NULL,
	[moto_gy_fudono] [varchar](max) NOT NULL,
	[biko_kihon] [varchar](max) NOT NULL,
	[kosu_total] [varchar](max) NOT NULL,
	[men_nobeyuka] [varchar](max) NOT NULL,
	[men_sikiti] [varchar](max) NOT NULL,
	[men_parking] [varchar](max) NOT NULL,
	[men_yukatoki] [varchar](max) NOT NULL,
	[men_sikititoki] [varchar](max) NOT NULL,
	[men_kentiku] [varchar](max) NOT NULL,
	[men_kentikutoki] [varchar](max) NOT NULL,
	[yatin_jisansaki] [varchar](max) NOT NULL,
	[yatin_kozano] [varchar](max) NOT NULL,
	[yatin_kykozaflg] [varchar](max) NOT NULL,
	[yatin_kykozano] [varchar](max) NOT NULL,
	[kozo_taikakbn] [varchar](max) NOT NULL,
	[gy_sekono] [varchar](max) NOT NULL,
	[gy_hosyuno] [varchar](max) NOT NULL,
	[kanri_hosiki] [varchar](max) NOT NULL,
	[kanri_gyname] [varchar](max) NOT NULL,
	[kanri_gytanto] [varchar](max) NOT NULL,
	[kanri_gytel] [varchar](max) NOT NULL,
	[kanrinin_gyomukeitai] [varchar](max) NOT NULL,
	[kanrinin_tel] [varchar](max) NOT NULL,
	[jisya_no] [varchar](max) NOT NULL,
	[kotu_sonota1] [varchar](max) NOT NULL,
	[kotu_sonota1kyori] [varchar](max) NOT NULL,
	[kotu_sonota2] [varchar](max) NOT NULL,
	[kotu_sonota2kyori] [varchar](max) NOT NULL,
	[denki_gy_lifeno] [varchar](max) NOT NULL,
	[water_gy_lifeno] [varchar](max) NOT NULL,
	[gas_gy_lifeno] [varchar](max) NOT NULL,
	[haisui_gy_lifeno] [varchar](max) NOT NULL,
	[toyu_gy_lifeno] [varchar](max) NOT NULL,
	[lifeline1_gy_lineno] [varchar](max) NOT NULL,
	[lifeline2_gy_lineno] [varchar](max) NOT NULL,
	[lifeline3_gy_lineno] [varchar](max) NOT NULL,
	[kensingyomu_umu] [varchar](max) NOT NULL,
	[parking_kanriflg] [varchar](max) NOT NULL,
	[parking_car] [varchar](max) NOT NULL,
	[parking_bike] [varchar](max) NOT NULL,
	[parking_bicycle] [varchar](max) NOT NULL,
	[parking_bicyclefreeflg] [varchar](max) NOT NULL,
	[bk_logonuser_no] [varchar](max) NOT NULL,
	[kanri_gyfax] [varchar](max) NOT NULL,
	[homeelevator_flg] [varchar](max) NOT NULL,
	[kanrinin_name] [varchar](max) NOT NULL,
	[parking_caraki] [varchar](max) NOT NULL,
	[parking_bikeaki] [varchar](max) NOT NULL,
	[parking_bicycleaki] [varchar](max) NOT NULL,
	[syogaku_name] [varchar](max) NOT NULL,
	[syogaku_kyori] [varchar](max) NOT NULL,
	[cyugaku_name] [varchar](max) NOT NULL,
	[cyugaku_kyori] [varchar](max) NOT NULL,
	[bk_area_no] [varchar](max) NOT NULL,
	[emergencyelevator_flg] [varchar](max) NOT NULL,
	[memo1] [varchar](max) NOT NULL,
	[memo2] [varchar](max) NOT NULL,
	[memo3] [varchar](max) NOT NULL,
	[memo4] [varchar](max) NOT NULL,
	[memo5] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_物件所有者情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_物件所有者情報];
End

CREATE TABLE [dbo].[pre_物件所有者情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[kasi1_ow_no] [varchar](max) NOT NULL,
	[syo1_ow_no] [varchar](max) NOT NULL,
	[kasi2_ow_no] [varchar](max) NOT NULL,
	[syo2_ow_no] [varchar](max) NOT NULL,
	[syo_startymd] [varchar](max) NOT NULL,
	[syo_endymd] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_物件近隣駐車場情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_物件近隣駐車場情報];
End

CREATE TABLE [dbo].[pre_物件近隣駐車場情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[parking_no] [varchar](max) NOT NULL,
	[naiyo] [varchar](max) NOT NULL,
	[kyori] [varchar](max) NOT NULL,
	[gak] [varchar](max) NOT NULL,
	[zeikbn] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_物件鍵情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_物件鍵情報];
End

CREATE TABLE [dbo].[pre_物件鍵情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[kagi_no] [varchar](max) NOT NULL,
	[honsu] [varchar](max) NOT NULL,
	[biko] [varchar](max) NOT NULL,
	[hokan] [varchar](max) NOT NULL,
	[gyshare] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_物件交通情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_物件交通情報];
End

CREATE TABLE [dbo].[pre_物件交通情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[ensen_cnt] [varchar](max) NOT NULL,
	[setting_kotutype] [varchar](max) NOT NULL,
	[ensen_no] [varchar](max) NOT NULL,
	[eki_no] [varchar](max) NOT NULL,
	[firsttrain_flg] [varchar](max) NOT NULL,
	[kyori] [varchar](max) NOT NULL,
	[toho_min] [varchar](max) NOT NULL,
	[bus_companytoeki] [varchar](max) NOT NULL,
	[bus_stationtoeki] [varchar](max) NOT NULL,
	[bus_min] [varchar](max) NOT NULL,
	[bus_kyoritoeki] [varchar](max) NOT NULL,
	[bus_tohomintoeki] [varchar](max) NOT NULL,
	[car_kyori] [varchar](max) NOT NULL,
	[car_min] [varchar](max) NOT NULL,
	[bus_company] [varchar](max) NOT NULL,
	[bus_station] [varchar](max) NOT NULL,
	[bus_kyori] [varchar](max) NOT NULL,
	[bus_tohomin] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_部屋情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_部屋情報];
End

CREATE TABLE [dbo].[pre_部屋情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[hy_no] [varchar](max) NOT NULL,
	[hy_ruinokbn] [varchar](max) NOT NULL,
	[madori] [varchar](max) NOT NULL,
	[madori_biko] [varchar](max) NOT NULL,
	[men_senyujitu] [varchar](max) NOT NULL,
	[men_yuka] [varchar](max) NOT NULL,
	[men_senyutouki] [varchar](max) NOT NULL,
	[men_toki] [varchar](max) NOT NULL,
	[men_balcony] [varchar](max) NOT NULL,
	[kaiyaku_months] [varchar](max) NOT NULL,
	[syozai_kaisu1] [varchar](max) NOT NULL,
	[syozai_tikaflg1] [varchar](max) NOT NULL,
	[mukikbn] [varchar](max) NOT NULL,
	[balcony_mukikbn] [varchar](max) NOT NULL,
	[nyukyo_jokyokbn] [varchar](max) NOT NULL,
	[nyukyo_jokyomemo] [varchar](max) NOT NULL,
	[nyukyo_syunflg] [varchar](max) NOT NULL,
	[nyukyo_ym] [varchar](max) NOT NULL,
	[nyukyo_syunkbn] [varchar](max) NOT NULL,
	[nyukyo_joken] [varchar](max) NOT NULL,
	[kakunin_ymd] [varchar](max) NOT NULL,
	[bosyu_startflg] [varchar](max) NOT NULL,
	[torihiki_taiyokbn] [varchar](max) NOT NULL,
	[torihiki_kyakutuke] [varchar](max) NOT NULL,
	[moto_gy_fudono] [varchar](max) NOT NULL,
	[biko] [varchar](max) NOT NULL,
	[hosyo_gyno] [varchar](max) NOT NULL,
	[hosyo_naiyo] [varchar](max) NOT NULL,
	[kyrui_nokbn] [varchar](max) NOT NULL,
	[keiyaku_kikan] [varchar](max) NOT NULL,
	[keiyaku_kijitu] [varchar](max) NOT NULL,
	[parking_biko] [varchar](max) NOT NULL,
	[parking_bikebiko] [varchar](max) NOT NULL,
	[parking_cyurinbiko] [varchar](max) NOT NULL,
	[shared_salespoint] [varchar](max) NOT NULL,
	[hosyo_kbn] [varchar](max) NOT NULL,
	[jisya_no] [varchar](max) NOT NULL,
	[jisya_tanto] [varchar](max) NOT NULL,
	[nextnkin_kosindefault] [varchar](max) NOT NULL,
	[toki_ymd] [varchar](max) NOT NULL,
	[siyo_mokuteki] [varchar](max) NOT NULL,
	[nyukyo_jikikbn] [varchar](max) NOT NULL,
	[torihiki_jisyakbn] [varchar](max) NOT NULL,
	[parking_gakkbn] [varchar](max) NOT NULL,
	[parking_gak] [varchar](max) NOT NULL,
	[parking_gakzei] [varchar](max) NOT NULL,
	[parking_tintaisu] [varchar](max) NOT NULL,
	[genjyo_naiyo] [varchar](max) NOT NULL,
	[nyukyo_naiyo] [varchar](max) NOT NULL,
	[etc_naiyo] [varchar](max) NOT NULL,
	[memo1] [varchar](max) NOT NULL,
	[memo2] [varchar](max) NOT NULL,
	[memo3] [varchar](max) NOT NULL,
	[memo4] [varchar](max) NOT NULL,
	[memo5] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_部屋所有者情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_部屋所有者情報];
End

CREATE TABLE [dbo].[pre_部屋所有者情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[hy_no] [varchar](max) NOT NULL,
	[kasi1_ow_no] [varchar](max) NOT NULL,
	[syo1_ow_no] [varchar](max) NOT NULL,
	[kasi2_ow_no] [varchar](max) NOT NULL,
	[syo2_ow_no] [varchar](max) NOT NULL,
	[syo_startymd] [varchar](max) NOT NULL,
	[syo_endymd] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_部屋入金項目情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_部屋入金項目情報];
End

CREATE TABLE [dbo].[pre_部屋入金項目情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[hy_no] [varchar](max) NOT NULL,
	[tuki_kbn] [varchar](max) NOT NULL,
	[nkin_name] [varchar](max) NOT NULL,
	[nkin_recno] [varchar](max) NOT NULL,
	[nkin_kbn] [varchar](max) NOT NULL,
	[sq_gak] [varchar](max) NOT NULL,
	[sq_zeikbn] [varchar](max) NOT NULL,
	[calc_kbn] [varchar](max) NOT NULL,
	[calc_nkinno] [varchar](max) NOT NULL,
	[calc_monthcnt] [varchar](max) NOT NULL,
	[sqsaki_no] [varchar](max) NOT NULL,
	[sq_mmkbn] [varchar](max) NOT NULL,
	[sqstart_ymd] [varchar](max) NOT NULL,
	[sq_interval] [varchar](max) NOT NULL,
	[biko] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_部屋設備情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_部屋設備情報];
End

CREATE TABLE [dbo].[pre_部屋設備情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[hy_no] [varchar](max) NOT NULL,
	[S001] [varchar](2) NOT NULL,	[S002] [varchar](2) NOT NULL,	[S003] [varchar](2) NOT NULL,	[S004] [varchar](2) NOT NULL,	[S005] [varchar](2) NOT NULL,	[S006] [varchar](2) NOT NULL,	[S007] [varchar](2) NOT NULL,	[S008] [varchar](2) NOT NULL,	[S009] [varchar](2) NOT NULL,	[S010] [varchar](2) NOT NULL,
	[S011] [varchar](2) NOT NULL,	[S012] [varchar](2) NOT NULL,	[S013] [varchar](2) NOT NULL,	[S014] [varchar](2) NOT NULL,	[S015] [varchar](2) NOT NULL,	[S016] [varchar](2) NOT NULL,	[S017] [varchar](2) NOT NULL,	[S018] [varchar](2) NOT NULL,	[S019] [varchar](2) NOT NULL,	[S020] [varchar](2) NOT NULL,
	[S021] [varchar](2) NOT NULL,	[S022] [varchar](2) NOT NULL,	[S023] [varchar](2) NOT NULL,	[S024] [varchar](2) NOT NULL,	[S025] [varchar](2) NOT NULL,	[S026] [varchar](2) NOT NULL,	[S027] [varchar](2) NOT NULL,	[S028] [varchar](2) NOT NULL,	[S029] [varchar](2) NOT NULL,	[S030] [varchar](2) NOT NULL,
	[S031] [varchar](2) NOT NULL,	[S032] [varchar](2) NOT NULL,	[S033] [varchar](2) NOT NULL,	[S034] [varchar](2) NOT NULL,	[S035] [varchar](2) NOT NULL,	[S036] [varchar](2) NOT NULL,	[S037] [varchar](2) NOT NULL,	[S038] [varchar](2) NOT NULL,	[S039] [varchar](2) NOT NULL,	[S040] [varchar](2) NOT NULL,
	[S041] [varchar](2) NOT NULL,	[S042] [varchar](2) NOT NULL,	[S043] [varchar](2) NOT NULL,	[S044] [varchar](2) NOT NULL,	[S045] [varchar](2) NOT NULL,	[S046] [varchar](2) NOT NULL,	[S047] [varchar](2) NOT NULL,	[S048] [varchar](2) NOT NULL,	[S049] [varchar](2) NOT NULL,	[S050] [varchar](2) NOT NULL,
	[S051] [varchar](2) NOT NULL,	[S052] [varchar](2) NOT NULL,	[S053] [varchar](2) NOT NULL,	[S054] [varchar](2) NOT NULL,	[S055] [varchar](2) NOT NULL,	[S056] [varchar](2) NOT NULL,	[S057] [varchar](2) NOT NULL,	[S058] [varchar](2) NOT NULL,	[S059] [varchar](2) NOT NULL,	[S060] [varchar](2) NOT NULL,
	[S061] [varchar](2) NOT NULL,	[S062] [varchar](2) NOT NULL,	[S063] [varchar](2) NOT NULL,	[S064] [varchar](2) NOT NULL,	[S065] [varchar](2) NOT NULL,	[S066] [varchar](2) NOT NULL,	[S067] [varchar](2) NOT NULL,	[S068] [varchar](2) NOT NULL,	[S069] [varchar](2) NOT NULL,	[S070] [varchar](2) NOT NULL,
	[S071] [varchar](2) NOT NULL,	[S072] [varchar](2) NOT NULL,	[S073] [varchar](2) NOT NULL,	[S074] [varchar](2) NOT NULL,	[S075] [varchar](2) NOT NULL,	[S076] [varchar](2) NOT NULL,	[S077] [varchar](2) NOT NULL,	[S078] [varchar](2) NOT NULL,	[S079] [varchar](2) NOT NULL,	[S080] [varchar](2) NOT NULL,
	[S081] [varchar](2) NOT NULL,	[S082] [varchar](2) NOT NULL,	[S083] [varchar](2) NOT NULL,	[S084] [varchar](2) NOT NULL,	[S085] [varchar](2) NOT NULL,	[S086] [varchar](2) NOT NULL,	[S087] [varchar](2) NOT NULL,	[S088] [varchar](2) NOT NULL,	[S089] [varchar](2) NOT NULL,	[S090] [varchar](2) NOT NULL,
	[S091] [varchar](2) NOT NULL,	[S092] [varchar](2) NOT NULL,	[S093] [varchar](2) NOT NULL,	[S094] [varchar](2) NOT NULL,	[S095] [varchar](2) NOT NULL,	[S096] [varchar](2) NOT NULL,	[S097] [varchar](2) NOT NULL,	[S098] [varchar](2) NOT NULL,	[S099] [varchar](2) NOT NULL,	[S100] [varchar](2) NOT NULL,
	[S101] [varchar](2) NOT NULL,	[S102] [varchar](2) NOT NULL,	[S103] [varchar](2) NOT NULL,	[S104] [varchar](2) NOT NULL,	[S105] [varchar](2) NOT NULL,	[S106] [varchar](2) NOT NULL,	[S107] [varchar](2) NOT NULL,	[S108] [varchar](2) NOT NULL,	[S109] [varchar](2) NOT NULL,	[S110] [varchar](2) NOT NULL,
	[S111] [varchar](2) NOT NULL,	[S112] [varchar](2) NOT NULL,	[S113] [varchar](2) NOT NULL,	[S114] [varchar](2) NOT NULL,	[S115] [varchar](2) NOT NULL,	[S116] [varchar](2) NOT NULL,	[S117] [varchar](2) NOT NULL,	[S118] [varchar](2) NOT NULL,	[S119] [varchar](2) NOT NULL,	[S120] [varchar](2) NOT NULL,
	[S121] [varchar](2) NOT NULL,	[S122] [varchar](2) NOT NULL,	[S123] [varchar](2) NOT NULL,	[S124] [varchar](2) NOT NULL,	[S125] [varchar](2) NOT NULL,	[S126] [varchar](2) NOT NULL,	[S127] [varchar](2) NOT NULL,	[S128] [varchar](2) NOT NULL,	[S129] [varchar](2) NOT NULL,	[S130] [varchar](2) NOT NULL,
	[S131] [varchar](2) NOT NULL,	[S132] [varchar](2) NOT NULL,	[S133] [varchar](2) NOT NULL,	[S134] [varchar](2) NOT NULL,	[S135] [varchar](2) NOT NULL,	[S136] [varchar](2) NOT NULL,	[S137] [varchar](2) NOT NULL,	[S138] [varchar](2) NOT NULL,	[S139] [varchar](2) NOT NULL,	[S140] [varchar](2) NOT NULL,
	[S141] [varchar](2) NOT NULL,	[S142] [varchar](2) NOT NULL,	[S143] [varchar](2) NOT NULL,	[S144] [varchar](2) NOT NULL,	[S145] [varchar](2) NOT NULL,	[S146] [varchar](2) NOT NULL,	[S147] [varchar](2) NOT NULL,	[S148] [varchar](2) NOT NULL,	[S149] [varchar](2) NOT NULL,	[S150] [varchar](2) NOT NULL,
	[S151] [varchar](2) NOT NULL,	[S152] [varchar](2) NOT NULL,	[S153] [varchar](2) NOT NULL,	[S154] [varchar](2) NOT NULL,	[S155] [varchar](2) NOT NULL,	[S156] [varchar](2) NOT NULL,	[S157] [varchar](2) NOT NULL,	[S158] [varchar](2) NOT NULL,	[S159] [varchar](2) NOT NULL,	[S160] [varchar](2) NOT NULL,
	[S161] [varchar](2) NOT NULL,	[S162] [varchar](2) NOT NULL,	[S163] [varchar](2) NOT NULL,	[S164] [varchar](2) NOT NULL,	[S165] [varchar](2) NOT NULL,	[S166] [varchar](2) NOT NULL,	[S167] [varchar](2) NOT NULL,	[S168] [varchar](2) NOT NULL,	[S169] [varchar](2) NOT NULL,	[S170] [varchar](2) NOT NULL,
	[S171] [varchar](2) NOT NULL,	[S172] [varchar](2) NOT NULL,	[S173] [varchar](2) NOT NULL,	[S174] [varchar](2) NOT NULL,	[S175] [varchar](2) NOT NULL,	[S176] [varchar](2) NOT NULL,	[S177] [varchar](2) NOT NULL,	[S178] [varchar](2) NOT NULL,	[S179] [varchar](2) NOT NULL,	[S180] [varchar](2) NOT NULL,
	[S181] [varchar](2) NOT NULL,	[S182] [varchar](2) NOT NULL,	[S183] [varchar](2) NOT NULL,	[S184] [varchar](2) NOT NULL,	[S185] [varchar](2) NOT NULL,	[S186] [varchar](2) NOT NULL,	[S187] [varchar](2) NOT NULL,	[S188] [varchar](2) NOT NULL,	[S189] [varchar](2) NOT NULL,	[S190] [varchar](2) NOT NULL,
	[S191] [varchar](2) NOT NULL,	[S192] [varchar](2) NOT NULL,	[S193] [varchar](2) NOT NULL,	[S194] [varchar](2) NOT NULL,	[S195] [varchar](2) NOT NULL,	[S196] [varchar](2) NOT NULL,	[S197] [varchar](2) NOT NULL,	[S198] [varchar](2) NOT NULL,	[S199] [varchar](2) NOT NULL,	[S200] [varchar](2) NOT NULL,
	[S201] [varchar](2) NOT NULL,	[S202] [varchar](2) NOT NULL,	[S203] [varchar](2) NOT NULL,	[S204] [varchar](2) NOT NULL,	[S205] [varchar](2) NOT NULL,	[S206] [varchar](2) NOT NULL,	[S207] [varchar](2) NOT NULL,	[S208] [varchar](2) NOT NULL,	[S209] [varchar](2) NOT NULL,	[S210] [varchar](2) NOT NULL,
	[S211] [varchar](2) NOT NULL,	[S212] [varchar](2) NOT NULL,	[S213] [varchar](2) NOT NULL,	[S214] [varchar](2) NOT NULL,	[S215] [varchar](2) NOT NULL,	[S216] [varchar](2) NOT NULL,	[S217] [varchar](2) NOT NULL,	[S218] [varchar](2) NOT NULL,	[S219] [varchar](2) NOT NULL,	[S220] [varchar](2) NOT NULL,
	[S221] [varchar](2) NOT NULL,	[S222] [varchar](2) NOT NULL,	[S223] [varchar](2) NOT NULL,	[S224] [varchar](2) NOT NULL,	[S225] [varchar](2) NOT NULL,	[S226] [varchar](2) NOT NULL,	[S227] [varchar](2) NOT NULL,	[S228] [varchar](2) NOT NULL,	[S229] [varchar](2) NOT NULL,	[S230] [varchar](2) NOT NULL,
	[S231] [varchar](2) NOT NULL,	[S232] [varchar](2) NOT NULL,	[S233] [varchar](2) NOT NULL,	[S234] [varchar](2) NOT NULL,	[S235] [varchar](2) NOT NULL,	[S236] [varchar](2) NOT NULL,	[S237] [varchar](2) NOT NULL,	[S238] [varchar](2) NOT NULL,	[S239] [varchar](2) NOT NULL,	[S240] [varchar](2) NOT NULL,
	[S241] [varchar](2) NOT NULL,	[S242] [varchar](2) NOT NULL,	[S243] [varchar](2) NOT NULL,	[S244] [varchar](2) NOT NULL,	[S245] [varchar](2) NOT NULL,	[S246] [varchar](2) NOT NULL,	[S247] [varchar](2) NOT NULL,	[S248] [varchar](2) NOT NULL,	[S249] [varchar](2) NOT NULL,	[S250] [varchar](2) NOT NULL,
	[S251] [varchar](2) NOT NULL,	[S252] [varchar](2) NOT NULL,	[S253] [varchar](2) NOT NULL,	[S254] [varchar](2) NOT NULL,	[S255] [varchar](2) NOT NULL,	[S256] [varchar](2) NOT NULL,	[S257] [varchar](2) NOT NULL,	[S258] [varchar](2) NOT NULL,	[S259] [varchar](2) NOT NULL,	[S260] [varchar](2) NOT NULL,
	[S261] [varchar](2) NOT NULL,	[S262] [varchar](2) NOT NULL,	[S263] [varchar](2) NOT NULL,	[S264] [varchar](2) NOT NULL,	[S265] [varchar](2) NOT NULL,	[S266] [varchar](2) NOT NULL,	[S267] [varchar](2) NOT NULL,	[S268] [varchar](2) NOT NULL,	[S269] [varchar](2) NOT NULL,	[S270] [varchar](2) NOT NULL,
	[S271] [varchar](2) NOT NULL,	[S272] [varchar](2) NOT NULL,	[S273] [varchar](2) NOT NULL,	[S274] [varchar](2) NOT NULL,	[S275] [varchar](2) NOT NULL,	[S276] [varchar](2) NOT NULL,	[S277] [varchar](2) NOT NULL,	[S278] [varchar](2) NOT NULL,	[S279] [varchar](2) NOT NULL,	[S280] [varchar](2) NOT NULL,
	[S281] [varchar](2) NOT NULL,	[S282] [varchar](2) NOT NULL,	[S283] [varchar](2) NOT NULL,	[S284] [varchar](2) NOT NULL,	[S285] [varchar](2) NOT NULL,	[S286] [varchar](2) NOT NULL,	[S287] [varchar](2) NOT NULL,	[S288] [varchar](2) NOT NULL,	[S289] [varchar](2) NOT NULL,	[S290] [varchar](2) NOT NULL,
	[S291] [varchar](2) NOT NULL,	[S292] [varchar](2) NOT NULL,	[S293] [varchar](2) NOT NULL,	[S294] [varchar](2) NOT NULL,	[S295] [varchar](2) NOT NULL,	[S296] [varchar](2) NOT NULL,	[S297] [varchar](2) NOT NULL,	[S298] [varchar](2) NOT NULL,	[S299] [varchar](2) NOT NULL,	[S300] [varchar](2) NOT NULL,
	[S301] [varchar](2) NOT NULL,	[S302] [varchar](2) NOT NULL,	[S303] [varchar](2) NOT NULL,	[S304] [varchar](2) NOT NULL,	[S305] [varchar](2) NOT NULL,	[S306] [varchar](2) NOT NULL,	[S307] [varchar](2) NOT NULL,	[S308] [varchar](2) NOT NULL,	[S309] [varchar](2) NOT NULL,	[S310] [varchar](2) NOT NULL,
	[S311] [varchar](2) NOT NULL,	[S312] [varchar](2) NOT NULL,	[S313] [varchar](2) NOT NULL,	[S314] [varchar](2) NOT NULL,	[S315] [varchar](2) NOT NULL,	[S316] [varchar](2) NOT NULL,	[S317] [varchar](2) NOT NULL,	[S318] [varchar](2) NOT NULL,	[S319] [varchar](2) NOT NULL,	[S320] [varchar](2) NOT NULL,
	[S321] [varchar](2) NOT NULL,	[S322] [varchar](2) NOT NULL,	[S323] [varchar](2) NOT NULL,	[S324] [varchar](2) NOT NULL,	[S325] [varchar](2) NOT NULL,	[S326] [varchar](2) NOT NULL,	[S327] [varchar](2) NOT NULL,	[S328] [varchar](2) NOT NULL,	[S329] [varchar](2) NOT NULL,	[S330] [varchar](2) NOT NULL,
	[S331] [varchar](2) NOT NULL,	[S332] [varchar](2) NOT NULL,	[S333] [varchar](2) NOT NULL,	[S334] [varchar](2) NOT NULL,	[S335] [varchar](2) NOT NULL,	[S336] [varchar](2) NOT NULL,	[S337] [varchar](2) NOT NULL,	[S338] [varchar](2) NOT NULL,	[S339] [varchar](2) NOT NULL,	[S340] [varchar](2) NOT NULL,
	[S341] [varchar](2) NOT NULL,	[S342] [varchar](2) NOT NULL,	[S343] [varchar](2) NOT NULL,	[S344] [varchar](2) NOT NULL,	[S345] [varchar](2) NOT NULL,	[S346] [varchar](2) NOT NULL,	[S347] [varchar](2) NOT NULL,	[S348] [varchar](2) NOT NULL,	[S349] [varchar](2) NOT NULL,	[S350] [varchar](2) NOT NULL,
	[S351] [varchar](2) NOT NULL,	[S352] [varchar](2) NOT NULL,	[S353] [varchar](2) NOT NULL,	[S354] [varchar](2) NOT NULL,	[S355] [varchar](2) NOT NULL,	[S356] [varchar](2) NOT NULL,	[S357] [varchar](2) NOT NULL,	[S358] [varchar](2) NOT NULL,	[S359] [varchar](2) NOT NULL,	[S360] [varchar](2) NOT NULL,
	[S361] [varchar](2) NOT NULL,	[S362] [varchar](2) NOT NULL,	[S363] [varchar](2) NOT NULL,	[S364] [varchar](2) NOT NULL,	[S365] [varchar](2) NOT NULL,	[S366] [varchar](2) NOT NULL,	[S367] [varchar](2) NOT NULL,	[S368] [varchar](2) NOT NULL,	[S369] [varchar](2) NOT NULL,	[S370] [varchar](2) NOT NULL,
	[S371] [varchar](2) NOT NULL,	[S372] [varchar](2) NOT NULL,	[S373] [varchar](2) NOT NULL,	[S374] [varchar](2) NOT NULL,	[S375] [varchar](2) NOT NULL,	[S376] [varchar](2) NOT NULL,	[S377] [varchar](2) NOT NULL,	[S378] [varchar](2) NOT NULL,	[S379] [varchar](2) NOT NULL,	[S380] [varchar](2) NOT NULL,
	[S381] [varchar](2) NOT NULL,	[S382] [varchar](2) NOT NULL,	[S383] [varchar](2) NOT NULL,	[S384] [varchar](2) NOT NULL,	[S385] [varchar](2) NOT NULL,	[S386] [varchar](2) NOT NULL,	[S387] [varchar](2) NOT NULL,	[S388] [varchar](2) NOT NULL,	[S389] [varchar](2) NOT NULL,	[S390] [varchar](2) NOT NULL,
	[S391] [varchar](2) NOT NULL,	[S392] [varchar](2) NOT NULL,	[S393] [varchar](2) NOT NULL,	[S394] [varchar](2) NOT NULL,	[S395] [varchar](2) NOT NULL,	[S396] [varchar](2) NOT NULL,	[S397] [varchar](2) NOT NULL,	[S398] [varchar](2) NOT NULL,	[S399] [varchar](2) NOT NULL,	[S400] [varchar](2) NOT NULL,
	[S401] [varchar](2) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_部屋間取内訳情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_部屋間取内訳情報];
End

CREATE TABLE [dbo].[pre_部屋間取内訳情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[hy_no] [varchar](max) NOT NULL,
	[no] [varchar](max) NOT NULL,
	[hykbn] [varchar](max) NOT NULL,
	[jo] [varchar](max) NOT NULL,
	[syozaikai] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_部屋鍵情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_部屋鍵情報];
End

CREATE TABLE [dbo].[pre_部屋鍵情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[hy_no] [varchar](max) NOT NULL,
	[kagi_no] [varchar](max) NOT NULL,
	[honsu] [varchar](max) NOT NULL,
	[biko] [varchar](max) NOT NULL,
	[hokan] [varchar](max) NOT NULL,
	[gyshare] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_物件管理情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_物件管理情報];
End

CREATE TABLE [dbo].[pre_物件管理情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[hy_no] [varchar](max) NOT NULL,
	[kanri_keitaikbn] [varchar](max) NOT NULL,
	[ikkatu_kanriflg] [varchar](max) NOT NULL,
	[cyukai_kyflg] [varchar](max) NOT NULL,
	[cyukai_koflg] [varchar](max) NOT NULL,
	[cyukai_kaiflg] [varchar](max) NOT NULL,
	[soymd1_gaitoukbn] [varchar](max) NOT NULL,
	[soymd1_simekbn] [varchar](max) NOT NULL,
	[soymd1_sokintukikbn] [varchar](max) NOT NULL,
	[soymd1_sokinsimekbn] [varchar](max) NOT NULL,
	[ikkatu_kbn] [varchar](max) NOT NULL,
	[ikkatu_bkgak] [varchar](max) NOT NULL,
	[kanritesu_kbn] [varchar](max) NOT NULL,
	[kanritesu_cyosyukbn] [varchar](max) NOT NULL,
	[kanri_reigaiky] [varchar](max) NOT NULL,
	[kanri_reigaikai] [varchar](max) NOT NULL,
	[hyteigaku_kbn] [varchar](max) NOT NULL,
	[hyteigaku_kanrigak] [varchar](max) NOT NULL,
	[bkteigaku_kanrigak] [varchar](max) NOT NULL,
	[kanri_taxflg] [varchar](max) NOT NULL,
	[biko_basic] [varchar](max) NOT NULL,
	[kanritesu_zeiumukbn] [varchar](max) NOT NULL,
	[no_soruleflg] [varchar](max) NOT NULL,
	[cyukai_zuijisokin] [varchar](max) NOT NULL,
	[hyteigaku_hiwariflg] [varchar](max) NOT NULL CONSTRAINT [DF_pre_物件管理情報_hyteigaku_hiwariflg] DEFAULT '',
	[hyteigaku_hiwarikbn] [varchar](max) NOT NULL CONSTRAINT [DF_pre_物件管理情報_hyteigaku_hiwarikbn] DEFAULT ''
);

IF OBJECT_ID(N'[dbo].[pre_物件管理送金先情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_物件管理送金先情報];
End

CREATE TABLE [dbo].[pre_物件管理送金先情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[hy_no] [varchar](max) NOT NULL,
	[sosaki_recno] [varchar](max) NOT NULL,
	[so_ow_no] [varchar](max) NOT NULL,
	[so_ow_kozano] [varchar](max) NOT NULL,
	[so_rit] [varchar](max) NOT NULL,
	[so_fixsogak] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_物件管理入金項目情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_物件管理入金項目情報];
End

CREATE TABLE [dbo].[pre_物件管理入金項目情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[hy_no] [varchar](max) NOT NULL,
	[taisyokbn] [varchar](max) NOT NULL,
	[nkin_sortorder] [varchar](max) NOT NULL,
	[nkin_name] [varchar](max) NOT NULL,
	[sokin_rit] [varchar](max) NOT NULL,
	[kanrigak_rit] [varchar](max) NOT NULL,
	[hosyo_flg] [varchar](max) NOT NULL,
	[so_no] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_物件管理控除項目情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_物件管理控除項目情報];
End

CREATE TABLE [dbo].[pre_物件管理控除項目情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[hy_no] [varchar](max) NOT NULL,
	[taisyokbn] [varchar](max) NOT NULL,
	[nkin_sortorder] [varchar](max) NOT NULL,
	[nkin_name] [varchar](max) NOT NULL,
	[sosai_flg] [varchar](max) NOT NULL,
	[kojo_gak] [varchar](max) NOT NULL,
	[kojo_gakzeikbn] [varchar](max) NOT NULL,
	[tateazu_flg] [varchar](max) NOT NULL,
	[kojo_zeigak] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_契約情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_契約情報];
End

CREATE TABLE [dbo].[pre_契約情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[hy_no] [varchar](max) NOT NULL,
	[status] [varchar](max) NOT NULL,
	[status_ymd] [varchar](max) NOT NULL,
	[cyukai_gy_fudono] [varchar](max) NOT NULL,
	[syunin_no] [varchar](max) NOT NULL,
	[kaiyaku_months] [varchar](max) NOT NULL,
	[cyukai_tantoname] [varchar](max) NOT NULL,
	[kystart_ymd] [varchar](max) NOT NULL,
	[kyend_ymd] [varchar](max) NOT NULL,
	[ky_bango] [varchar](max) NOT NULL,
	[ky_ymd] [varchar](max) NOT NULL,
	[kanri_gy_fudono] [varchar](max) NOT NULL,
	[gy_hosyono] [varchar](max) NOT NULL,
	[hosyo_naiyo] [varchar](max) NOT NULL,
	[kokyaku_bango] [varchar](max) NOT NULL,
	[kyrui_no] [varchar](max) NOT NULL,
	[siyo_mokuteki] [varchar](max) NOT NULL,
	[kosin_umu] [varchar](max) NOT NULL,
	[yatin_kbn] [varchar](max) NOT NULL,
	[fkae_startym] [varchar](max) NOT NULL,
	[yatin_kozakbn] [varchar](max) NOT NULL,
	[yatin_kozano] [varchar](max) NOT NULL,
	[maitukiyatin_kozano] [varchar](max) NOT NULL,
	[biko] [varchar](max) NOT NULL,
	[nextky_startymd] [varchar](max) NOT NULL,
	[nextky_endymd] [varchar](max) NOT NULL,
	[soyotei_ymd] [varchar](max) NOT NULL,
	[kanritesu_flg] [varchar](max) NOT NULL,
	[kanritesu_gak] [varchar](max) NOT NULL,
	[nextnkinset_kbn] [varchar](max) NOT NULL,
	[biko2] [varchar](max) NOT NULL,
	[hoken_biko] [varchar](max) NOT NULL,
	[nkinsime_ymd] [varchar](max) NOT NULL,
	[yatin_jisansaki] [varchar](max) NOT NULL,
	[yokugetu_uketoriumu] [varchar](max) NOT NULL,
	[yokugetu_uketorimonth] [varchar](max) NOT NULL,
	[ky_logonuser_no] [varchar](max) NOT NULL,
	[sq_logonuser_no] [varchar](max) NOT NULL,
	[tuti_ymd] [varchar](max) NOT NULL CONSTRAINT [DF_pre_契約情報_tuti_ymd] DEFAULT ''
);

IF OBJECT_ID(N'[dbo].[pre_契約契約者保証人情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_契約契約者保証人情報];
End

CREATE TABLE [dbo].[pre_契約契約者保証人情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[hy_no] [varchar](max) NOT NULL,
	[kys1_no] [varchar](max) NOT NULL,
	[nyukyo1_flg] [varchar](max) NOT NULL,
	[kys2_no] [varchar](max) NOT NULL,
	[nyukyo2_flg] [varchar](max) NOT NULL,
	[kys3_no] [varchar](max) NOT NULL,
	[nyukyo3_flg] [varchar](max) NOT NULL,
	[kys_hosyo_no] [varchar](max) NOT NULL,
	[hosyo1_no] [varchar](max) NOT NULL,
	[hosyo2_no] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_契約特約およびメモ情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_契約特約およびメモ情報];
End

CREATE TABLE [dbo].[pre_契約特約およびメモ情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[hy_no] [varchar](max) NOT NULL,
	[genjo_tokuyaku] [varchar](max) NOT NULL,
	[syuzen_tokuyaku] [varchar](max) NOT NULL,
	[sonota_tokuyaku] [varchar](max) NOT NULL,
	[memo1] [varchar](max) NOT NULL,
	[memo2] [varchar](max) NOT NULL,
	[memo3] [varchar](max) NOT NULL,
	[memo4] [varchar](max) NOT NULL,
	[memo5] [varchar](max) NOT NULL,
	[memo6] [varchar](max) NOT NULL,
	[memo7] [varchar](max) NOT NULL,
	[memo8] [varchar](max) NOT NULL,
	[memo9] [varchar](max) NOT NULL,
	[memo10] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_契約保険情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_契約保険情報];
End

CREATE TABLE [dbo].[pre_契約保険情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[hy_no] [varchar](max) NOT NULL,
	[hoken_no] [varchar](max) NOT NULL,
	[gy_hokenno] [varchar](max) NOT NULL,
	[ky_ymd] [varchar](max) NOT NULL,
	[hoken_gak] [varchar](max) NOT NULL,
	[biko] [varchar](max) NOT NULL,
	[syoken_bango][varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_契約入金項目情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_契約入金項目情報];
End

CREATE TABLE [dbo].[pre_契約入金項目情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[hy_no] [varchar](max) NOT NULL,
	[tuki_kbn] [varchar](max) NOT NULL,
	[nkin_name] [varchar](max) NOT NULL,
	[nkin_recno] [varchar](max) NOT NULL,
	[nkin_sortorder] [varchar](max) NOT NULL,
	[nkin_kbn] [varchar](max) NOT NULL,
	[sq_gak] [varchar](max) NOT NULL,
	[sq_zeikbn] [varchar](max) NOT NULL,
	[sq_zeigak] [varchar](max) NOT NULL,
	[calc_kbn] [varchar](max) NOT NULL,
	[calc_nkinno] [varchar](max) NOT NULL,
	[calc_monthcnt] [varchar](max) NOT NULL,
	[sqsaki_no] [varchar](max) NOT NULL,
	[nkbn_yotei] [varchar](max) NOT NULL,
	[sq_mmkbn] [varchar](max) NOT NULL,
	[sqstart_ymd] [varchar](max) NOT NULL,
	[sq_interval] [varchar](max) NOT NULL,
	[biko] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_契約次回入金項目情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_契約次回入金項目情報];
End

CREATE TABLE [dbo].[pre_契約次回入金項目情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[hy_no] [varchar](max) NOT NULL,
	[tuki_kbn] [varchar](max) NOT NULL,
	[nkin_name] [varchar](max) NOT NULL,
	[nkin_recno] [varchar](max) NOT NULL,
	[nkin_sortorder] [varchar](max) NOT NULL,
	[nkin_kbn] [varchar](max) NOT NULL,
	[sq_gak] [varchar](max) NOT NULL,
	[sq_zeikbn] [varchar](max) NOT NULL,
	[sq_zeigak] [varchar](max) NOT NULL,
	[calc_kbn] [varchar](max) NOT NULL,
	[calc_nkinno] [varchar](max) NOT NULL,
	[calc_monthcnt] [varchar](max) NOT NULL,
	[sqsaki_no] [varchar](max) NOT NULL,
	[nkbn_yotei] [varchar](max) NOT NULL,
	[sq_mmkbn] [varchar](max) NOT NULL,
	[sqstart_ymd] [varchar](max) NOT NULL,
	[sq_interval] [varchar](max) NOT NULL,
	[biko] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_契約入居者情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_契約入居者情報];
End

CREATE TABLE [dbo].[pre_契約入居者情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[hy_no] [varchar](max) NOT NULL,
	[nyukyosya_cnt] [varchar](max) NOT NULL,
	[kys_no] [varchar](max) NOT NULL,
	[name] [varchar](max) NOT NULL,
	[kana] [varchar](max) NOT NULL,
	[gender] [varchar](max) NOT NULL,
	[aidagara] [varchar](max) NOT NULL,
	[birthday] [varchar](max) NOT NULL,
	[kinmusaki] [varchar](max) NOT NULL,
	[tel] [varchar](max) NOT NULL,
	[mobiletel] [varchar](max) NOT NULL,
	[biko] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_契約車情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_契約車情報];
End

CREATE TABLE [dbo].[pre_契約車情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[hy_no] [varchar](max) NOT NULL,
	[car_cnt] [varchar](max) NOT NULL,
	[carmaker] [varchar](max) NOT NULL,
	[carname] [varchar](max) NOT NULL,
	[carcolor] [varchar](max) NOT NULL,
	[carnumber] [varchar](max) NOT NULL,
	[biko] [varchar](max) NOT NULL,
	[car_kukaku] [varchar](max) NOT NULL
);

IF OBJECT_ID(N'[dbo].[pre_家主固定控除情報]', N'U') IS NOT NULL
Begin
	Drop Table [dbo].[pre_家主固定控除情報];
End

CREATE TABLE [dbo].[pre_家主固定控除情報](
	[ROW_ID] [int] NOT NULL,
	[bk_no] [varchar](max) NOT NULL,
	[nkin_sortorder] [varchar](max) NOT NULL,
	[nkin_name] [varchar](max) NOT NULL,
	[hy_no] [varchar](max) NOT NULL,
	[sosai_flg] [varchar](max) NOT NULL,
	[gak] [varchar](max) NOT NULL,
	[zei_kbn] [varchar](max) NOT NULL,
	[zei_gak] [varchar](max) NOT NULL,
	[nkbn_yotei] [varchar](max) NOT NULL,
	[yatin_kozano] [varchar](max) NOT NULL,
	[gtstart_ym] [varchar](max) NOT NULL,
	[gtend_ym] [varchar](max) NOT NULL,
	[tateazu_flg] [varchar](max) NOT NULL,
	[sq_mmkbn] [varchar](max) NOT NULL,
	[sq_interval] [varchar](max) NOT NULL,
	[biko] [varchar](max) NOT NULL
);
