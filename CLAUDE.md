# fk10-convert-kit2 作業メモ（Claude 向け）

## 作業の進め方

- 流れ：**修正 → ビルド → ユーザーが動作確認 → OK の連絡 → コミット**。
  - Claude は画面（GUI）を操作できない。画面の動作はユーザーに確認してもらい、確認前にコミットしない。
  - 実際のコードを動かして確認したいときは、ビルドした exe を PowerShell のリフレクションで読み込んで呼び出す
    （開発用に作ったスクラッチ DB・一時フォルダの中だけで行い、確認後は元に戻す）。
- 修正後は、変更点・確認してほしい手順を、箇条書きで短く伝える。
- ユーザーへの返答は「です・ます」調にする。

## Git の運用ルール

- **コミット・push は、ユーザーに頼まれたときだけ Claude が実行する。** 頼まれていないのに、コミットや push をしない。
  - Claude のシェルは Windows 上のユーザー環境で動くため、Git の設定・認証情報はそのまま使える。
- コミットのルール：
  1. 論理的なまとまりごとに分ける（機能追加・不具合修正・リリースノートは別コミットにするのが基本）
  2. `git add` はファイルを指定する（`git add -A` は使わない）。`git status` で、関係ないファイルが入っていないか確認する
  3. メッセージは日本語。1 行目に要約、必要なら空行のあとに箇条書きで補足する
  4. ユーザーに無断で、過去のコミットの書き換え（amend・rebase・force push）をしない
- **push は `origin` と `njc` の両方**に行う。
  ```powershell
  git push origin main
  git push njc main
  ```
- `.vscode/` は個人設定なのでコミットしない（`git status` では未追跡のまま残る）。
- 改行コードの警告（`LF will be replaced by CRLF`）は、無害なので無視してよい。
- Git を読むだけの操作（status / diff / log）には、`git --no-optional-locks` を付けるとよい。

### リモート
- `origin` … https://github.com/t-suenaga-git/fk10-convert-kit2.git
- `njc` … https://git.njc-web.info/t-suenaga/fk10-convert-kit2.git

## 構成

```
Converter10.sln
├─ Converter10_CS\   汎用コンバートキット2 本体（C#, SDK形式, net48）→ Converter10v2.exe
└─ RelationSetting\  紐付設定ツール（VB.NET, SDK形式, net48）→ RelationSetting.exe
```

- RelationSetting は kit2 から `Process.Start` で呼ばれる別 exe。
  kit2 の exe と同じ場所の `RelationSetting\RelationSetting.exe` を起動し、引数（`cmdline(1)`〜`cmdline(9)`）で接続情報やパスを渡す。
  **引数を変えるときは、kit2 の `MainFrm.cs`（紐付設定画面の呼び出し処理）と `RelationFrm.vb`（`Set_CmdlineInfoToObj`）を必ずセットで直す。**
- kit2 をビルドすると、RelationSetting も先にビルドされ、kit2 の出力先の `RelationSetting\` に exe と依存 DLL がコピーされる（`Converter10_CS.csproj` の `CopyRelationSetting` ターゲット）。
  紐付設定ファイルのひな形（`RelationSetting\relationfile\*.xlsx`）は、出力先に無いときだけコピーする。
- `MainFrm` は部分クラスで、機能ごとにファイルを分けている：
  - `MainFrm.cs` … 本体（画面制御・コンバート処理）
  - `MainFrm.PreTableConnection.cs` … Pre テーブル接続情報の入力・保存（`ini\coninfo_10.xml`）
  - `MainFrm.CvResult.cs` … 移行結果 CSV の出力、紐付設定ファイルの複製、出力先フォルダを開く処理
- 移行のしくみ：Pre テーブル（EF）→ `CVTBL_*` → 移行先の本番テーブル。移行先は**新スキーマ（`fk8db_20260917`）のみ**に対応（旧スキーマには対応しない）。

### 実行時の出力・設定

- ログなどの成果物は、実行ごとに `deliverables\<日時>\` へ保存する（日時は 1 回の実行で共通）。
  - コンバートログ `cv_log_<日時>.csv`、中間ファイルチェックログ `midchk_log_<日時>.csv`
  - 移行結果 `移行結果_<日時>.csv`
  - 紐付設定ファイルの複製 `各マスタ紐付け情報_<日時>.xlsx`（元ファイルは、前回の紐付けの復元・結果保存に使うので変更しない）
- 接続設定は `ini\coninfo_10.xml` に保存する（パスワードは暗号化）。
- `middlefile` と `list` フォルダは使わない（存在しなくてよい）。

## ビルド

- VS Code の Ctrl+Shift+B（`.vscode\tasks.json` の「MSBuild Direct」）
  - Build Tools for Visual Studio 18 の MSBuild で `Converter10.sln /restore /p:Configuration=Debug`
- コマンドで、本体だけビルドする場合：
  ```
  "C:\Program Files (x86)\Microsoft Visual Studio\18\BuildTools\MSBuild\Current\Bin\MSBuild.exe" Converter10_CS\Converter10_CS.csproj /p:Configuration=Debug /nologo /v:minimal
  ```
  - COM 参照があるため、`dotnet build` は使えない（MSBuild を使う）。
  - NuGet パッケージを変えたときは `/restore` を付ける。
  - アプリを起動したままだと exe がロックされてビルドできない（MSB3027）。先に閉じてもらう。
- kit2 の出力先：`Converter10_CS\bin\Debug\net48\`

## リリース

バージョンは `Converter10_CS\My Project\AssemblyInfo.cs` で設定する（画面のタイトルにも表示される）。

1. リリースノート（`リリースノート.md`）の先頭に、`## ver X.Y.Z.W(日付)` の見出しを入れ、変更内容を追記する
2. **Release 構成**でビルドする（`/p:Configuration=Release`）
3. 配布用フォルダ `ConvertKit2-X.Y.Z.W\` を作り、`Converter10_CS\bin\Release\net48\` の中身をコピーする。次を差し替え・確認する：
   - `ConnectionStrings.config` は、**`ConnectionStrings.config.example`（ひな形）に置き換える**（DB のログイン情報入りの実ファイルを、zip・Git に入れない）
   - `RelationSetting\relationfile\各マスタ紐付け情報.xlsx` は、リポジトリのひな形（データなし）のままであること
   - 起動時に自動で作られる `deliverables`・`ini`・`RelationSetting\rellog` を、配布物に含めない
4. zip にして `ReleaseBin\ConvertKit2-X.Y.Z.W.zip` に置く（zip 内のトップフォルダも `ConvertKit2-X.Y.Z.W`）
   - zip のエントリ名は `/` 区切りにする（PowerShell 5.1 の `ZipFile.CreateFromDirectory` は `\` 区切りになるので使わず、`ZipArchive` でエントリを 1 つずつ作る）
   - 作成後に展開し、exe のバージョンと、`ConnectionStrings.config` がひな形であることを確認する。可能なら起動も確認する
5. コミットし、注釈付きタグ `vX.Y.Z.W` を付けて、`origin` と `njc` の両方へ push する（ブランチとタグの両方）
   ```powershell
   git tag -a vX.Y.Z.W -m "汎用コンバートキット2 verX.Y.Z.W"
   git push origin main; git push origin vX.Y.Z.W
   git push njc main;    git push njc vX.Y.Z.W
   ```
   - 公開済みのタグを、あとから付け替え（force）ない。直したいときは、新しいバージョンにする

## 前提・方針

- **64ビット対応**：両プロジェクトとも `PlatformTarget=AnyCPU`、`Prefer32Bit=false`。
- **ACE.OLEDB（Access Database Engine）は使わない。** Excel ファイルの読み書きに ACE を使う処理を新しく追加しない。
  - RelationSetting の紐付設定ファイルの読み書きは ClosedXML（`RelationSetting\Com\RelationExcelFile.vb`）。
  - 本体（Converter10_CS）には、ClosedXML を入れない（Excel の出力が必要になっても、CSV にする）。
- Excel の COM 操作（Interop）は今も多く使っているので、実行には Excel 本体が必要。
- **DB の接続情報（パスワードなど）は、リポジトリに入れない。** `Converter10_CS\ConnectionStrings.config` は Git 管理外（雛形は `ConnectionStrings.config.example`）。
- ソースファイルは **UTF-8（BOM付き）・CRLF**。編集するときはこれを崩さない。
- 修正箇所には、既存の書き方に合わせて `'YYYYMMDD 内容 -chg sta` / `-chg end`（C# は `//`）のコメントを付ける。
- SQL の `INSERT` は**列リストを書く**（列名を省略すると、移行先に列が増えたときに無言で失敗する）。
- `DBExec.Exec_NonQuery` の失敗は `cv_log`（小分類「SQL実行」）に出る。ログに出ないエラーを、握りつぶさない。
