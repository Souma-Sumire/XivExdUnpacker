# XivExdUnpacker

[Lumina](https://github.com/NotAdam/Lumina) に基づくファイナルファンタジーXIV EXD データ解凍ツール。`SaintCoinach.Cmd` の `rawexd` 機能の代替として設計されています。

[English](../README.md) | [日本語](./README.ja.md) | [Deutsch](./README.de.md) | [Français](./README.fr.md) | [简体中文](./README.cn.md) | [한국어](./README.ko.md) | [繁體中文](./README.tc.md)

## 利用方法

### 必要条件

- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- FFXIV のローカルインストール

### インストール

1. [Releases](https://github.com/Souma-Sumire/XivExdUnpacker/releases) から最新のリリースをダウンロードします。
2. アーカイブを解凍します。

### 設定

`config.yml.example` を `config.yml` にコピーし、各サーバーの `path` と `outputDir` を編集します。

### 実行

```bash
# ヘルプを表示
XivExdUnpacker.exe --help

# 中国語のすべてのテーブルをエクスポート
XivExdUnpacker.exe --language cn

# 英語と日本語の Action と Item テーブルをエクスポート
XivExdUnpacker.exe --language en ja --sheets Action Item

# すべての言語をエクスポートし、hexcode を出力し、エクスポート前に出力ディレクトリをクリア
XivExdUnpacker.exe --language all --hexcode --clear
```

### コマンドライン引数

| 引数 | 短縮形 | 説明 | デフォルト |
| ---- | ---- | ---- | ------ |
| `--language` | `-l` | エクスポートする言語を指定 (必須) | - |
| `--sheets` | `-s` | エクスポートするシート名を指定 | すべて |
| `--hexcode` | `-x` | 生データを保持 | false |
| `--clear` | `-c` | エクスポート前に出力ディレクトリをクリア | false |
| `--skip-offset` | - | CSV の offset 行をスキップ | false |
| `--help` | `-h` | ヘルプ情報を表示 | - |

## 開発

### プロジェクトの準備

```bash
# リポジトリをクローン
git clone --recursive https://github.com/Souma-Sumire/XivExdUnpacker.git
cd XivExdUnpacker
# サブモジュールの初期化
git submodule update --init --recursive
# Schema サブモジュールを最新に更新
git submodule update --remote
```

### 実行

```bash
# ヘルプを表示
dotnet run -- --help
# 以下省略、上記と同様
```
