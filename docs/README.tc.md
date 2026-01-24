# XivExdUnpacker

一個基於 [Lumina](https://github.com/NotAdam/Lumina) 的最終幻想 XIV EXD 資料解包工具，用於平替 `SaintCoinach.Cmd` 的 `rawexd` 功能。

[English](../README.md) | [日本語](./README.ja.md) | [Deutsch](./README.de.md) | [Français](./README.fr.md) | [简体中文](./README.cn.md) | [한국어](./README.ko.md) | [繁體中文](./README.tc.md)

## 環境要求

- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- FFXIV 的本地安裝

## 快速開始

### 準備項目

```bash
# 複製項目
git clone --recursive https://github.com/Souma-Sumire/XivExdUnpacker.git
cd XivExdUnpacker
# 初始化子模塊
git submodule update --init --recursive
# 更新 Schema 子模塊到最新版本
git submodule update --remote
```

### 配置

將 `config.yml.example` 複製為 `config.yml`，並編輯設置各服的 `path` 和 `outputDir`

### 運行

```bash
# 顯示幫助信息
dotnet run -- --help

# 匯出中文的所有表 (默認解碼字串)
dotnet run -- --language cn

# 匯出英文和日文的 Action 和 Item 表
dotnet run -- --language en ja --sheets Action Item

# 匯出所有語言，保留原始資料，清空輸出目錄
dotnet run -- --language all --hexcode --clear

# 使用簡寫，匯出中文，清空輸出目錄，使用原始 HEX
dotnet run -- -l cn -c -x
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
