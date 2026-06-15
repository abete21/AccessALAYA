# Access ALAYA Version 2.03  
日本語版 ReadMe

---

## 基本情報

**ソフト名**: Access ALAYA Version 2.03  
**制作者名**: abete21  
**動作環境**:  
- Windows 10 以降  
- Microsoft Access（2013 以降で動作確認済み、32bit/64bit どちらも対応）  

**種別**: ビジネス / データベース関連（Access）  
**連絡先**: AccessALAYA@gmail.com  
**URL**: https://accessalaya.jp  

---

## 改訂履歴
Alaya2.03  2026/06/14  Fix English version module and manuals. (The First release)
Alaya2.00  2026/05/29  Start to develop English version
Alaya1.20  2025/04/29  フォームコントロール用クラスモジュールの追加と複合キー対応
Alaya1.12  2024/03/30  SQLにループ定義とマクロ外出しコマンドを追加
Alaya1.10  2023/09/02  標準Formの生成機能追加
Alaya1.09  2023/08/26  インストーラ導入
---

## 概要

Access ALAYA は、Microsoft Access を **マルチユーザで安全かつ安定して運用するためのフレームワーク**です。

Access は Excel のように手軽にデータベースを扱える反面、  
スタンドアロンまたは 2～3 名程度の利用を前提としているため、  
**デフォルト設定のまま多数のユーザが同時利用すると破損しやすい脆弱性**があります。

Access ALAYA はこの問題を解決し、  
**数十名規模の同時利用でも安定して稼働できる環境**を提供します。

実際の運用例では：

- 登録ユーザ：100 名  
- 同時ログイン：30 名  
- ノーメンテで 1 年以上安定稼働  

という実績があります。

---

## Access ALAYA の特徴

### ● マルチユーザ対応  
最初から多人数利用を前提に設計されています。  
後からマルチユーザ向けチューニングを行う必要はありません。

### ● ローコード開発  
共通関数をライブラリ化し、できるだけコードを書かずに開発できます。  
ユーザロジックに集中でき、メンテナンス性も向上します。

### ● バージョンアップと配布管理  
頻繁な機能追加や修正を前提とし、  
- モジュール更新  
- バージョン管理  
- エンドユーザへの自動配布  

を行う仕組みを備えています。

---

## 動作環境・必要ソフト

- Microsoft Access が動作する Windows  
- Access 2013 以降（32bit/64bit どちらも可）

※ Mac 版 Office には Access が含まれないため、Mac では動作しません。

---

## 📦 ZIP アーカイブの内容

リリースパッケージ `AccessALAYA2.03.zip` には以下が含まれます。

- **AccessALAYA/**
  - README.md — 英語版説明  
  - README_JP.md — 日本語版説明  
  - **docs/**
    - Starting AccessALAYA.pptx — はじめての利用ガイド  
  - **sample/**
    - ReadMe.md — サンプルアプリ説明  
  - **src/**
    - **ALAYA Installer/** — インストーラモジュール  
      - ALAYA Installer1.01.accdb  
      - ALAYA Installer1.02.accdb  
      - ALAYA Installer1.03.accdb（最新版）  
    - **ALAYADB/** — サンプル共有テーブル  
      - Membership.accdb  
      - SchedIndex.accdb  
      - Schedule.accdb  
      - Tasks.accdb  
    - **ALAYAMenu/** — メイン GUI モジュール  
      - ALAYA Menu**.accdb — ユーザインタフェース  
      - ALAYA Menu.rel — リリース定義ファイル  
      - StartALAYA.vbs — 起動スクリプト（バージョン管理・配布管理）  
      - PrivateDB/ — ローカル DB（v2.03 は空）  
      - xlTemplate/ — Excel テンプレート（v2.03 は空）  
    - **DBBackup/** — バックアップ用（インストール時は空）

ZIP を展開後、`sample` フォルダ内のアプリを Access で開くと  
ALAYA フレームワークの動作を確認できます。

---

## インストール方法

1. ZIP を展開し、ファイルサーバに配置します。  
2. `src/ALAYA Installer/` 内の **最新の ALAYA Installer***.accdb** を起動します。  
3. 指示に従ってサーバへインストールしてください。

---

## アンインストール方法

レジストリは使用していません。  
ファイルサーバ上の「プロジェクト名」フォルダを削除するだけでアンインストールできます。

---

## 使用方法

- ALAYA の起動は **VBS スクリプト経由**で行います。  
- メニューモジュールフォルダに生成される `StartALAYA.vbs` をダブルクリックしてください。  
- クライアントには **GUI モジュール (.accdb) を配布せず**、  
  サーバ上のスクリプトへのショートカットを配布します。

---

## ライセンス

- 個人利用・学術利用など **非営利目的は無償**  
- 商用利用・再配布・二次利用はライセンス契約が必要  
- 試用は無償ですが、**3 か月以上継続利用する場合は連絡をお願いします**

ライセンス契約いただいた場合は手厚いサポートを提供します。

---

## ご質問・お問い合わせ

Q&A フォーラム（推奨）：  
https://accessalaya.jp/alaya-qa/

その他の問い合わせ：  
AccessALAYA@gmail.com

---

以上
