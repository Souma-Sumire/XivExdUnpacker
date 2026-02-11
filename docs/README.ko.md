# XivExdUnpacker

[Lumina](https://github.com/NotAdam/Lumina)를 기반으로 한 파이널 판타지 XIV EXD 데이터 언패킹 도구로, `SaintCoinach.Cmd`의 `rawexd` 기능을 대체하기 위해 설계되었습니다.

[English](../README.md) | [日本語](./README.ja.md) | [Deutsch](./README.de.md) | [Français](./README.fr.md) | [简体中文](./README.cn.md) | [한국어](./README.ko.md) | [繁體中文](./README.tc.md)

## 사용

### 요구 사항

- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- FFXIV 로컬 설치

### 설치

1. [Releases](https://github.com/Souma-Sumire/XivExdUnpacker/releases)에서 최신 릴리스를 다운로드합니다.
2. 아카이브 압축을 풉니다.

### 설정

`config.yml.example`을 `config.yml`로 복사하고 각 서버의 `path` 및 `outputDir`을 편집합니다.

### 실행

```bash
# 도움말 표시
XivExdUnpacker.exe --help

# 중국어용 모든 테이블 내보내기
XivExdUnpacker.exe --language cn

# 영어 및 일본어용 Action 및 Item 테이블 내보내기
XivExdUnpacker.exe --language en ja --sheets Action Item

# 모든 언어 내보내기, hexcode 출력, 내보내기 전 출력 디렉토리 비우기
XivExdUnpacker.exe --language all --hexcode --clear
```

### 명령줄 인수

| 인수 | 약어 | 설명 | 기본값 |
| ---- | ---- | ---- | ------ |
| `--language` | `-l` | 내보낼 언어 지정 (필수) | - |
| `--sheets` | `-s` | 내보낼 시트 이름 지정 | 전체 |
| `--hexcode` | `-x` | 원본 데이터 유지 | false |
| `--clear` | `-c` | 내보내기 전 출력 디렉토리 비우기 | false |
| `--skip-offset` | `-o` | CSV 오프셋 행 건너뛰기 | false |
| `--help` | `-h` | 도움말 정보 표시 | - |

## 개발

```bash
git clone https://github.com/Souma-Sumire/XivExdUnpacker.git
cd XivExdUnpacker
dotnet run -- --language ko
```
