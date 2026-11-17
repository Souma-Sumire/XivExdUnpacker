# XivExdUnpacker

[Lumina](https://github.com/NotAdam/Lumina) ベースの『ファイナルファンタジーXIV』EXD データ展開ツールです。`SaintCoinach.Cmd` の `rawexd` 機能の代替として設計されています。

[English](../README.md) | [日本語](./README.ja.md) | [Deutsch](./README.de.md) | [Français](./README.fr.md) | [简体中文](./README.cn.md) | [한국어](./README.ko.md) | [繁體中文](./README.tc.md)

## クイックスタート

1. [Releases](https://github.com/Souma-Sumire/XivExdUnpacker/releases) からアーカイブをダウンロードして展開します（[.NET 10](https://dotnet.microsoft.com/download/dotnet/10.0) が必要です）。
2. `config.yml.example` を `config.yml` にコピーし、必要に応じてパスを設定します。
3. 実行します：

```bash
# 日本語の全シートを出力
XivExdUnpacker --language ja
# 全言語の Action と Item シートを <hex:...> 形式で出力し、出力前にディレクトリを空にする
XivExdUnpacker --language all --sheets Action Item --hexcode --clear
```

## コマンドライン引数

| 引数 | 短縮 | 説明 | デフォルト |
| --- | --- | --- | --- |
| `--language` | `-l` | 出力言語（`en`, `ja`, `de`, `fr`, `chs`, `ko`, `tc`, `all`） | 必須 |
| `--sheets` | `-s` | 出力するシート名 | すべて |
| `--hexcode` | `-x` | 制御タグを `<hex:...>` として保持（[フォーマット仕様](./STRING_FORMAT.md)） | `false` |
| `--clear` | `-c` | 出力前に出力ディレクトリを空にする | `false` |
| `--skip-offset` | `-o` | CSV の `offset` 行を省略する | `false` |
| `--help` | `-h` | ヘルプを表示 | - |

> 展開済みデータのみ必要な場合は [ffxiv-datamining-hexcode-mixed](https://github.com/Souma-Sumire/ffxiv-datamining-hexcode-mixed) をご覧ください。

## 開発

```bash
git clone https://github.com/Souma-Sumire/XivExdUnpacker.git
cd XivExdUnpacker
dotnet run -- --language ja
```
