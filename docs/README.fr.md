# XivExdUnpacker

Un outil d'extraction des données EXD de Final Fantasy XIV basé sur [Lumina](https://github.com/NotAdam/Lumina), conçu pour remplacer la commande `rawexd` de `SaintCoinach.Cmd`.

[English](../README.md) | [日本語](./README.ja.md) | [Deutsch](./README.de.md) | [Français](./README.fr.md) | [简体中文](./README.cn.md) | [한국어](./README.ko.md) | [繁體中文](./README.tc.md)

## Démarrage rapide

1. Téléchargez l'archive depuis [Releases](https://github.com/Souma-Sumire/XivExdUnpacker/releases) (nécessite [.NET 10](https://dotnet.microsoft.com/download/dotnet/10.0)).
2. Copiez `config.yml.example` vers `config.yml` et ajustez les chemins si besoin.
3. Exécutez :

```bash
# Exporter toutes les tables en français
XivExdUnpacker --language fr
# Exporter pour toutes les langues les tables Action et Item au format <hex:...>, en vidant d'abord la sortie
XivExdUnpacker --language all --sheets Action Item --hexcode --clear
```

## Options CLI

| Option | Court | Description | Défaut |
| --- | --- | --- | --- |
| `--language` | `-l` | Langues cibles (`en`, `ja`, `de`, `fr`, `chs`, `ko`, `tc`, `all`) | Requis |
| `--sheets` | `-s` | Tables à exporter | Toutes |
| `--hexcode` | `-x` | Conserver les balises de contrôle en `<hex:...>` ([format](./STRING_FORMAT.md)) | `false` |
| `--clear` | `-c` | Vider le dossier de sortie avant l'export | `false` |
| `--skip-offset` | `-o` | Ignorer la ligne `offset` dans les CSV | `false` |
| `--help` | `-h` | Afficher l'aide | - |

> Vous cherchez uniquement les données CSV ? Consultez [ffxiv-datamining-hexcode-mixed](https://github.com/Souma-Sumire/ffxiv-datamining-hexcode-mixed).

## Développement

```bash
git clone https://github.com/Souma-Sumire/XivExdUnpacker.git
cd XivExdUnpacker
dotnet run -- --language fr
```
