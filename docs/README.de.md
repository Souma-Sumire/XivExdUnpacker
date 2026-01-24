# XivExdUnpacker

Ein auf [Lumina](https://github.com/NotAdam/Lumina) basierendes Tool zum Entpacken von Final Fantasy XIV EXD-Daten, konzipiert als Ersatz für die `rawexd`-Funktionalität von `SaintCoinach.Cmd`.

[English](../README.md) | [日本語](./README.ja.md) | [Deutsch](./README.de.md) | [Français](./README.fr.md) | [简体中文](./README.cn.md) | [한국어](./README.ko.md) | [繁體中文](./README.tc.md)

## Anforderungen

- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- Lokale Installation von FFXIV

## Schnellstart

### Projekt vorbereiten

```bash
# Repository klonen
git clone --recursive https://github.com/Souma-Sumire/XivExdUnpacker.git
cd XivExdUnpacker
# Submodule initialisieren
git submodule update --init --recursive
# Schema-Submodul auf die neueste Version aktualisieren
git submodule update --remote
```

### Konfiguration

Kopieren Sie `config.yml.example` nach `config.yml` und bearbeiten Sie `path` und `outputDir` für jeden Server.

### Ausführen

```bash
# Hilfe anzeigen
dotnet run -- --help

# Alle Tabellen für Chinesisch exportieren (Standard-String-Dekodierung)
dotnet run -- --language cn

# Tabellen Action und Item für Englisch und Japanisch exportieren
dotnet run -- --language en ja --sheets Action Item

# Alle Sprachen exportieren, Rohdaten beibehalten, Ausgabeverzeichnis leeren
dotnet run -- --language all --hexcode --clear

# Kurzbeispiele: Chinesisch exportieren, Verzeichnis leeren, Roh-HEX verwenden
dotnet run -- -l cn -c -x
```

### Befehlszeilenargumente

| Argument | Kurz | Beschreibung | Standard |
| ---- | ---- | ---- | ------ |
| `--language` | `-l` | Zu exportierende Sprachen angeben (Erforderlich) | - |
| `--sheets` | `-s` | Namen der zu exportierenden Tabellen angeben | Alle |
| `--hexcode` | `-x` | Rohdaten beibehalten | false |
| `--clear` | `-c` | Ausgabeverzeichnis vor dem Export leeren | false |
| `--skip-offset` | - | CSV-Offset-Zeilen überspringen | false |
| `--help` | `-h` | Hilfeinformationen anzeigen | - |
