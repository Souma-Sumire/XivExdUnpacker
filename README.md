# XivExd Unpacker

基于 [Lumina](https://github.com/NotAdam/Lumina) 和 [EXDSchema](https://github.com/vi-xiv/EXDSchema) 的 FFXIV EXD 解包工具，用于代替 `SaintCoinach.Cmd`。

## 环境要求

- [.NET 10.0 SDK](https://dotnet.microsoft.com/download) 或更高版本
- FFXIV 的本地安装

## 快速开始

```bash
# 1. 克隆项目
git clone --recursive https://github.com/Souma-Sumire/ExdDump.git
cd ExdDump

# 如果克隆时漏掉了 --recursive 子模块，请执行：
git submodule update --init --recursive

# 更新 Schema 子模块到最新版本：
git submodule update --remote

# 2. 配置 (编辑 config.yml)
# 设置各服的 path 和 outputDir

# 3. 运行
dotnet run
```

---

FINAL FANTASY is a registered trademark of Square Enix Holdings Co., Ltd.

FINAL FANTASY XIV © SQUARE ENIX CO., LTD.
