# XivExdUnpacker

基於 [Lumina](https://github.com/NotAdam/Lumina) 的《Final Fantasy XIV》EXD 資料解包工具，用於替代 `SaintCoinach.Cmd` 的 `rawexd` 匯出功能。

[English](../README.md) | [日本語](./README.ja.md) | [Deutsch](./README.de.md) | [Français](./README.fr.md) | [简体中文](./README.cn.md) | [한국어](./README.ko.md) | [繁體中文](./README.tc.md)

## 快速開始

1. 從 [Releases](https://github.com/Souma-Sumire/XivExdUnpacker/releases) 下載對應平台的壓縮檔並解壓（需安裝 [.NET 10](https://dotnet.microsoft.com/download/dotnet/10.0)）。
2. 將 `config.yml.example` 複製為 `config.yml`，依需求調整路徑與 `outputRoot`。
3. 執行命令：

```bash
# 匯出繁體中文全部資料表
XivExdUnpacker --language tc
# 匯出所有語言的 Action 與 Item 表，保留 <hex:...> 控制碼並在匯出前清空輸出
XivExdUnpacker --language all --sheets Action Item --hexcode --clear
```

## 命令列參數

| 參數 | 簡寫 | 說明 | 預設值 |
| --- | --- | --- | --- |
| `--language` | `-l` | 目標語言（`en`、`ja`、`de`、`fr`、`chs`、`ko`、`tc`、`all`） | 必填 |
| `--sheets` | `-s` | 指定匯出的表名 | 全部 |
| `--hexcode` | `-x` | 保留控制碼為 `<hex:...>`（見[標籤格式說明](./STRING_FORMAT.md)） | `false` |
| `--clear` | `-c` | 匯出前清空輸出目錄 | `false` |
| `--skip-offset` | `-o` | 略過 CSV 中的 `offset` 列 | `false` |
| `--help` | `-h` | 顯示說明 | - |

> 只需現成的解包 CSV 資料？可直接瀏覽 [ffxiv-datamining-hexcode-mixed](https://github.com/Souma-Sumire/ffxiv-datamining-hexcode-mixed)。

## 本地開發

```bash
git clone https://github.com/Souma-Sumire/XivExdUnpacker.git
cd XivExdUnpacker
dotnet run -- --language tc
```
