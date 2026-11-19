# pre_ テーブルの定義が、3か所で一致しているかを確認する(開発用)
#   1. 同梱の Sql\01_AdjustDataQuery作成準備.Sql の CREATE TABLE
#   2. EFエンティティ(EntityFramework\Entity\pre_*.cs)
#   3. 追加列の一覧(Model\PreTableSchema.cs の AddedColumns)
# 使い方: powershell -ExecutionPolicy Bypass -File tools\Check-PreSchema.ps1
# 不一致があれば、内容を表示して終了コード 1 で終わる。

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$root = Split-Path -Parent $PSScriptRoot
$sqlPath = Join-Path $root 'Converter10_CS\Sql\01_AdjustDataQuery作成準備.Sql'
$entityDir = Join-Path $root 'Converter10_CS\EntityFramework\Entity'
$schemaPath = Join-Path $root 'Converter10_CS\Model\PreTableSchema.cs'

# ---- 1. SQL の CREATE TABLE ----
$sql = [IO.File]::ReadAllText($sqlPath, [Text.Encoding]::UTF8)
$sqlTables = @{}
foreach ($m in [regex]::Matches($sql, '(?s)CREATE TABLE \[dbo\]\.\[(pre_[^\]]+)\]\((.*?)\r?\n\);')) {
    $cols = @()
    # 1行に複数の列が並ぶことや、列名と型の間に空白がないこと([列名][型])があるため、「[列名]」+「[型]」の並びで拾う
    foreach ($c in [regex]::Matches($m.Groups[2].Value, '\[([^\]]+)\]\s*\[[A-Za-z]+\]')) { $cols += $c.Groups[1].Value }
    $sqlTables[$m.Groups[1].Value] = $cols
}

# ---- 2. EF エンティティ ----
$entTables = @{}
foreach ($f in Get-ChildItem $entityDir -Filter 'pre_*.cs') {
    $text = [IO.File]::ReadAllText($f.FullName, [Text.Encoding]::UTF8)
    $cols = @()
    foreach ($c in [regex]::Matches($text, '(?m)^\s*public \w+\??\s+(\w+)\s*\{\s*get;\s*set;\s*\}')) { $cols += $c.Groups[1].Value }
    $entTables[[IO.Path]::GetFileNameWithoutExtension($f.Name)] = $cols
}

# ---- 3. 追加列の一覧 ----
$schemaText = [IO.File]::ReadAllText($schemaPath, [Text.Encoding]::UTF8)
$added = @()
foreach ($m in [regex]::Matches($schemaText, 'new AddedColumn\("([^"]+)",\s*"([^"]+)"')) {
    $added += [pscustomobject]@{ Table = $m.Groups[1].Value; Column = $m.Groups[2].Value }
}

$ng = New-Object System.Collections.Generic.List[string]


# SQL と エンティティ の突き合わせ(テーブルごとにまとめて表示する)
foreach ($t in ($sqlTables.Keys + $entTables.Keys | Sort-Object -Unique)) {
    if (-not $sqlTables.ContainsKey($t)) { $ng.Add("SQLにないテーブル: $t"); continue }
    if (-not $entTables.ContainsKey($t)) { $ng.Add("エンティティにないテーブル: $t"); continue }
    $onlySql = @($sqlTables[$t] | Where-Object { $entTables[$t] -notcontains $_ })
    $onlyEnt = @($entTables[$t] | Where-Object { $sqlTables[$t] -notcontains $_ })
    if ($onlySql.Count -gt 0) { $ng.Add("$t : SQLだけにある列 $($onlySql.Count) 件 ($((($onlySql | Select-Object -First 5) -join ', '))...)") }
    if ($onlyEnt.Count -gt 0) { $ng.Add("$t : エンティティだけにある列 $($onlyEnt.Count) 件 ($((($onlyEnt | Select-Object -First 5) -join ', '))...)") }
}

# 追加列の一覧と、SQL・エンティティの突き合わせ
foreach ($a in $added) {
    if (-not $sqlTables.ContainsKey($a.Table) -or $sqlTables[$a.Table] -notcontains $a.Column) { $ng.Add("追加列の一覧にあるが、SQLにない: $($a.Table).$($a.Column)") }
    if (-not $entTables.ContainsKey($a.Table) -or $entTables[$a.Table] -notcontains $a.Column) { $ng.Add("追加列の一覧にあるが、エンティティにない: $($a.Table).$($a.Column)") }
}

Write-Host ("SQLのテーブル数: {0} / エンティティ数: {1} / 追加列の一覧: {2} 列" -f $sqlTables.Count, $entTables.Count, $added.Count)
if ($ng.Count -eq 0) {
    Write-Host 'OK: 一致しています。'
    exit 0
}
Write-Host ("NG: {0} 件の不一致があります。" -f $ng.Count)
$ng | ForEach-Object { Write-Host (' - ' + $_) }
exit 1
