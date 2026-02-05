# XivExdUnpacker

一個基於 [Lumina](https://github.com/NotAdam/Lumina) 的最終幻想 XIV EXD 資料解包工具，用於平替 `SaintCoinach.Cmd` 的 `rawexd` 功能。

[English](../README.md) | [日本語](./README.ja.md) | [Deutsch](./README.de.md) | [Français](./README.fr.md) | [简体中文](./README.cn.md) | [한국어](./README.ko.md) | [繁體中文](./README.tc.md)

## 使用

### 環境要求

- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- FFXIV 的本地安裝

### 安裝

1. 從 [Releases](https://github.com/Souma-Sumire/XivExdUnpacker/releases) 下載最新版本。
2. 解壓縮檔案。

### 配置

將 `config.yml.example` 複製為 `config.yml`，並編輯設置各服的 `path` 和 `outputDir`

### 運行

```bash
# 顯示幫助信息
XivExdUnpacker.exe --help

# 匯出中文的所有表
XivExdUnpacker.exe --language cn

# 匯出英文和日文的 Action 和 Item 表
XivExdUnpacker.exe --language en ja --sheets Action Item

# 匯出所有語言，輸出 hexcode，匯出前清空輸出目錄
XivExdUnpacker.exe --language all --hexcode --clear
```

### 命令行參數

| 參數 | 簡寫 | 說明 | 默認值 |
| ---- | ---- | ---- | ------ |
| `--language` | `-l` | 指定要匯出的語言 (必需) | - |
| `--sheets` | `-s` | 指定要匯出的表名 | 全部 |
| `--hexcode` | `-x` | 保留原始資料 | false |
| `--clear` | `-c` | 匯出前清空輸出目錄 | false |
| `--skip-offset` | - | 跳過 CSV 的 offset 行 | false |
| `--help` | `-h` | 顯示幫助信息 | - |

## 開發

```bash
git clone https://github.com/Souma-Sumire/XivExdUnpacker.git
cd XivExdUnpacker
dotnet run -- --language tc
```
