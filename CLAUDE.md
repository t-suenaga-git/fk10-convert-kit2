# fk10-convert-kit2 作業メモ（Claude 向け）

## Git の運用ルール（必ず守る）

- **コミット・push は Claude が実行しない。** コマンドを提示して、ユーザー（たけしさん）に実行してもらう。
  - 理由：Claude のシェルは隔離された Linux 環境で、Git のユーザー設定・認証情報が無い。
    Linux 側から Git を書き込み操作すると `index.lock` が残ったり、改行コードの差分が出たりする。
  - 社内 GitLab（njc）には Claude の環境からは接続できない。
- 「コミットして」「push して」と頼まれたら、次の流れにする：
  1. Claude が差分を確認し、`git add` の対象とコミットメッセージ（日本語）を提案する
  2. 作業前からの未コミット変更など、今回と関係ないファイルは含めないよう注意する
  3. ユーザーに次を実行してもらう（push は **origin と njc の両方**）
     ```powershell
     git add <対象ファイル>
     git status
     git commit -m "<メッセージ>"
     git push origin main
     git push njc main
     git log --oneline -1
     ```
  4. 結果を貼ってもらい、`(HEAD -> main, origin/main, njc/main)` になっているか確認する
- Claude が Git を読むだけの操作（status / diff / log）をするときは `git --no-optional-locks` を付ける。

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

## ビルド

- VS Code の Ctrl+Shift+B（`.vscode\tasks.json` の「MSBuild Direct」）
  - Build Tools for Visual Studio 18 の MSBuild で `Converter10.sln /restore /p:Configuration=Debug`
- kit2 の出力先：`Converter10_CS\bin\Debug\net48\`

## 前提・方針

- **64ビット対応**：両プロジェクトとも `PlatformTarget=AnyCPU`、`Prefer32Bit=false`。
- **ACE.OLEDB（Access Database Engine）は使わない。** Excel ファイルの読み書きに ACE を使う処理を新しく追加しない。
  - RelationSetting の紐付設定ファイルの読み書きは ClosedXML（`RelationSetting\Com\RelationExcelFile.vb`）。
- Excel の COM 操作（Interop）は今も多く使っているので、実行には Excel 本体が必要。
- ソースファイルは **UTF-8（BOM付き）・CRLF**。編集するときはこれを崩さない。
- 修正箇所には、既存の書き方に合わせて `'YYYYMMDD 内容 -chg sta` / `-chg end`（C# は `//`）のコメントを付ける。
