# XivExd Unpacker

基于 [Lumina](https://github.com/NotAdam/Lumina) 和 [EXDSchema](https://github.com/vi-xiv/EXDSchema) 的 FFXIV EXD 解包工具，用于代替 `SaintCoinach.Cmd` 导出 CSV 文件。

暂时只有hexcode。

## 环境要求

- [.NET 10.0 SDK](https://dotnet.microsoft.com/download) 或更高版本
- FFXIV 的本地安装

## 快速开始

```bash
# 1. 克隆项目
git clone https://github.com/Souma-Sumire/ExdDump.git
cd ExdDump

# 初始化子模块
git submodule update --init --recursive

# 更新 Schema 子模块到最新版本：
git submodule update --remote

# 2. 配置 (编辑 config.yml)
# 将 config.yml.example 复制为 config.yml，并编辑设置各服的 path 和 outputDir

# 3. 运行
dotnet run
```

## 命令行用法

```bash
# 进入交互式菜单
dotnet run

# 处理所有定义的客户端
dotnet run all

# 仅处理特定客户端
dotnet run cn
```

---

FINAL FANTASY is a registered trademark of Square Enix Holdings Co., Ltd.

FINAL FANTASY XIV © SQUARE ENIX CO., LTD.
