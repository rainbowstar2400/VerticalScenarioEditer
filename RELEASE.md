# リリース手順（M13）

## 1. ビルド（ポータブル出力）

```powershell
dotnet publish VerticalScenarioEditer/VerticalScenarioEditer.csproj -c Release -r win-x64 -o publish
```

## 2. インストーラ作成（Inno Setup）

1. Inno Setup をインストールします。
2. `Installer/VerticalScenarioEditer.iss` を Inno Setup で開き、"Compile" を実行します。
3. `Installer/output` にセットアップファイルが生成されます。

## 3. テスト

- `QA.md` のチェックリストを実行します。
