# CopilotStudio-Sample-Apps

Copilot Studio + Power Platform で構築したソリューションのサンプル集です。  
各ソリューションには Dataverse テーブル定義・モデル駆動型アプリ・Power Automate フロー・Copilot Studio エージェント・導入手順書が含まれます。

[![GitHub Copilot](https://img.shields.io/badge/GitHub%20Copilot-対応-blueviolet?style=for-the-badge&logo=github)](https://github.com/features/copilot)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=for-the-badge)](./LICENSE)

> [!NOTE]
> 本リポジトリは [ギークフジワラ](https://twitter.com/geekfujiwara) の実務経験・検証に基づき継続的に更新されています。

---

## ソリューション一覧

| ソリューション                                                     | 説明                                                                                                | フォルダ                         |
| ------------------------------------------------------------------ | --------------------------------------------------------------------------------------------------- | -------------------------------- |
| **[汎用ヘルプデスク](./HelpDeskGeneric/)**                         | チケット管理・ナレッジ管理・AI 回答案生成を統合したモデル駆動型アプリ + Copilot Studio エージェント | `HelpDeskGeneric/`               |
| **[請求書処理](./InvoiceProcessorSampleDemo/)**                    | PDF 請求書の AI 分析・自動登録 + モデル駆動型アプリによる請求書管理                                 | `InvoiceProcessorSampleDemo/`    |

---

## リポジトリ構成

```
CopilotStudio-Sample-Apps/
├── HelpDeskGeneric/                    # 汎用ヘルプデスク
│   ├── README.md
│   └── HelpDeskGeneric_1_1_0_3.zip
├── InvoiceProcessorSampleDemo/         # 請求書処理
│   ├── README.md
│   ├── InvoiceProcessorSmpleDemo_1_0_0_4.zip
│   └── InvoiceProcessorSmpleDemo_1_0_0_3_managed.zip
├── LICENSE
├── .gitignore
└── README.md                           # ← このファイル
```

---

## クイックスタート

### 1. リポジトリをクローン

```bash
git clone https://github.com/dayamashita1223/CopilotStudio-Sample-Apps.git
cd CopilotStudio-Sample-Apps
```

### 2. ソリューションを選択して導入

```bash
cd HelpDeskGeneric
```

各ソリューションの `README.md` と `docs/` 内の導入手順書に従ってセットアップしてください。

### 3. GitHub Copilot で新しいソリューションを開発

VS Code で開くと `.github/agents/` と `.github/skills/` が自動認識され、**GeekPowerCode** エージェントが GitHub Copilot Chat で使えるようになります。

```
@GeekPowerCode IT資産管理アプリを作成してください。
テーブル: Asset, AssetCategory, Location
フィールド: 資産名, シリアル番号, ステータス, 担当者
```

---

## GitHub Copilot カスタムエージェント

推奨モデル: Claude Opus 4.6

**GeekPowerCode** エージェントは [開発標準](docs/POWER_PLATFORM_DEVELOPMENT_STANDARD.md) に従い、以下を自動で考慮します:

- 英語スキーマ名でテーブル設計
- `createdby` システム列を報告者として活用
- `systemuser` テーブルへの Lookup でユーザー参照
- 先にデプロイ → Dataverse 接続確立 → 開発の順序
- 設計 → ユーザー承認 → 実装の順序

### 使い方

| 方法               | 手順                                              |
| ------------------ | ------------------------------------------------- |
| エージェントモード | VS Code Copilot Chat で **@GeekPowerCode** を選択 |
| スキルとして       | チャットで `/power-platform-standard` と入力      |
| URL を渡して       | `このリポジトリを参照して...` + リポジトリ URL    |

---

## 開発標準ドキュメント

| ドキュメント                                                                            | 内容                                                |
| --------------------------------------------------------------------------------------- | --------------------------------------------------- |
| [POWER_PLATFORM_DEVELOPMENT_STANDARD.md](./docs/POWER_PLATFORM_DEVELOPMENT_STANDARD.md) | コードファースト開発標準 — 設計原則・全フェーズ詳細 |
| [DATAVERSE_GUIDE.md](./docs/DATAVERSE_GUIDE.md)                                         | Dataverse 統合ガイド                                |
| [CONNECTOR_REFERENCE.md](./docs/CONNECTOR_REFERENCE.md)                                 | コネクタ設定リファレンス                            |
| [ADVANCED_PATTERNS.md](./docs/ADVANCED_PATTERNS.md)                                     | 高度な実装パターン                                  |

### 公式ドキュメント

- [Power Apps Code Apps](https://learn.microsoft.com/ja-jp/power-apps/developer/code-apps/)
- [Power Automate クラウドフロー](https://learn.microsoft.com/ja-jp/power-automate/overview-cloud)
- [Copilot Studio](https://learn.microsoft.com/ja-jp/microsoft-copilot-studio/)
- [Dataverse Web API](https://learn.microsoft.com/ja-jp/power-apps/developer/data-platform/webapi/overview)
- [Power Platform CLI](https://learn.microsoft.com/ja-jp/power-platform/developer/cli/introduction)
- [開発標準 検証記事](https://www.geekfujiwara.com/tech/powerplatform/8082/)

---

## 前提条件

| 項目                | 詳細                                                                                                                            |
| ------------------- | ------------------------------------------------------------------------------------------------------------------------------- |
| Visual Studio Code  | [Power Platform Tools 拡張機能](https://marketplace.visualstudio.com/items?itemName=microsoft-IsvExpTools.powerplatform-vscode) |
| GitHub Copilot      | VS Code 拡張機能（推奨モデル: Claude Opus 4.6）                                                                                 |
| Node.js             | LTS v18.x / v20.x                                                                                                               |
| Python 3.10+        | Dataverse 自動化スクリプト用                                                                                                    |
| PAC CLI             | 最新バージョン                                                                                                                  |
| Power Platform 環境 | Code Apps が有効化されていること                                                                                                |
| ライセンス          | Power Apps Premium / Microsoft Copilot Studio                                                                                   |

---

## ソリューションの追加方法

新しいソリューションを追加する場合は、リポジトリルートに新しいフォルダを作成してください:

```
CopilotStudio-Sample-Apps/
├── HelpDeskGeneric/          # 既存
├── NewSolution/              # 新規追加
│   ├── docs/
│   ├── scripts/
│   ├── src/                  # Code Apps を使う場合
│   ├── package.json          # Code Apps を使う場合
│   ├── .env.example
│   └── README.md
└── ...
```

**GeekPowerCode** エージェントにソリューションフォルダ内で作業を依頼すれば、開発標準に従って自動構築します。

---

## ライセンス

MIT License - 詳細は [LICENSE](./LICENSE) を参照してください。

## フィードバック

- 問題報告: [GitHub Issues](https://github.com/dayamashita1223/CopilotStudio-Sample-Apps/issues)
- X: [@geekfujiwara](https://twitter.com/geekfujiwara)
