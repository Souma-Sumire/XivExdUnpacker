# XivExdUnpacker

[Lumina](https://github.com/NotAdam/Lumina) 기반의 파이널 판타지 XIV EXD 데이터 언패킹 도구로, `SaintCoinach.Cmd`의 `rawexd` 기능을 대체합니다.

[English](../README.md) | [日本語](./README.ja.md) | [Deutsch](./README.de.md) | [Français](./README.fr.md) | [简体中文](./README.cn.md) | [한국어](./README.ko.md) | [繁體中文](./README.tc.md)

## 빠른 시작

1. [Releases](https://github.com/Souma-Sumire/XivExdUnpacker/releases)에서 압축 파일을 받아 해제합니다([.NET 10](https://dotnet.microsoft.com/download/dotnet/10.0) 필요).
2. `config.yml.example`을 `config.yml`로 복사하고 필요 시 경로를 수정합니다.
3. 실행:

```bash
# 한국어 전체 시트 내보내기
XivExdUnpacker --language ko
# 모든 언어의 Action 및 Item 시트를 <hex:...> 형식으로 내보내고, 출력 전 폴더 비우기
XivExdUnpacker --language all --sheets Action Item --hexcode --clear
```

## CLI 옵션

| 옵션 | 단축 | 설명 | 기본값 |
| --- | --- | --- | --- |
| `--language` | `-l` | 내보낼 언어(`en`, `ja`, `de`, `fr`, `chs`, `ko`, `tc`, `all`) | 필수 |
| `--sheets` | `-s` | 내보낼 시트 이름 | 전체 |
| `--hexcode` | `-x` | 제어 태그를 `<hex:...>` 형태로 유지([형식 안내](./STRING_FORMAT.md)) | `false` |
| `--clear` | `-c` | 내보내기 전 출력 디렉터리 비우기 | `false` |
| `--skip-offset` | `-o` | CSV의 `offset` 행 생략 | `false` |
| `--help` | `-h` | 도움말 표시 | - |

> 추출된 데이터만 필요하신가요? [ffxiv-datamining-hexcode-mixed](https://github.com/Souma-Sumire/ffxiv-datamining-hexcode-mixed)를 참고하세요.

## 개발

```bash
git clone https://github.com/Souma-Sumire/XivExdUnpacker.git
cd XivExdUnpacker
dotnet run -- --language ko
```
