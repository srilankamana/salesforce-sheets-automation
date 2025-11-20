# 🚀 Salesforce Opportunity URL Generator via Google Sheets

![Google Sheets](https://img.shields.io/badge/Google%20Sheets-34A853?style=for-the-badge&logo=google-sheets&logoColor=white) ![Salesforce](https://img.shields.io/badge/Salesforce-00A1E0?style=for-the-badge&logo=salesforce&logoColor=white)

API連携が制限された環境下で、Googleスプレッドシート関数のみを使用してSalesforceの案件作成を半自動化する業務効率化ツール。
**Salesforce URL Hacking**（URLパラメータによる値の受け渡し）技術を応用し、Lightning Experience環境での動的なフォーム入力を実現しました。

## 📝 背景と課題 (Context)
経理・請求業務において、毎月スプレッドシートの管理表からSalesforceへ数十件の案件データを手動で転記していました。

* **課題:** 1件あたり約5分の入力工数がかかり、手入力による金額や日付の転記ミス、項目選択ミスのリスクが常にある状態。
* **制約:** 社内のセキュリティポリシーにより、ZapierなどのiPaaSや外部API連携ツールの導入が許可されていなかった。

## 💡 解決策 (Solution)
Salesforceの新規作成画面URLに `defaultFieldValues` パラメータを付与することで、スプレッドシート上のデータを保持した状態で入力フォームを起動する仕組みを構築しました。

### 主な機能
1.  **動的なURL生成:** 企業名、金額、入社日などの変数データを元に、SalesforceのURLを関数で動的に生成。
2.  **データクレンジング:** `REGEXREPLACE` 関数を使用し、元データに含まれる不要な文字列（管理番号など）を自動除去。
3.  **日付の自動計算:** 入社日から「契約開始日（当月1日）」や「完了予定日（当月末日）」を自動算出。
4.  **Lightning対応:** Classic環境とは異なるパラメータ仕様に対応し、カンマ区切りの値を適切に処理。

## 🛠️ 技術スタック (Tech Stack)
* **Platform:** Google Sheets, Salesforce Lightning Experience
* **Functions:** `HYPERLINK`, `ENCODEURL`, `REGEXREPLACE`, `TEXT`, `EDATE`, `EOMONTH`, `SPLIT/INDEX`

## 💻 実装コード (Implementation)

スプレッドシートのセル文字数制限や解析エラーを回避するため、**「URL生成セル」**と**「リンク表示セル」**を分離する設計を採用しました。

### 1. URL生成用セル (Backend Cell)
URLパラメータを組み立てるロジック部分です。
※ IDやドメインはダミーに変更しています。

```excel
="https://" & YOUR_DOMAIN & "[.lightning.force.com/lightning/o/Opportunity/new](https://.lightning.force.com/lightning/o/Opportunity/new)?" &
"recordTypeId=" & YOUR_RECORD_TYPE_ID &
"&defaultFieldValues=" &
  "Pricebook2Id=" & YOUR_PRICEBOOK_ID & "," &
  "AccountId=" & AM4 & "," &
  "Name=" & ENCODEURL("案件_" & TEXT(EDATE(G4, 1), "yymm") & "_" & REGEXREPLACE(B4, "^[0-9]+ ", "")) & "," &
  "Type=" & ENCODEURL("New Business") & "," &
  "StageName=" & ENCODEURL("Prospecting") & "," &
  "AmountAlt__c=" & J4 & "," &
  "CloseDate=" & TEXT(EOMONTH(G4, 0), "YYYY-MM-DD") & "," &
  "ContractStartDate__c=" & TEXT(EOMONTH(G4, -1) + 1, "YYYY-MM-DD") & "," &
  "CMRR_Flag__c=true"
```

### 2. リンク表示用セル (Frontend Cell)
ユーザーがクリックするリンクを生成します。URL生成セル（AO4）を参照します。
```excel
=HYPERLINK(AO4, "Salesforceで案件作成")
```
## 🧩 工夫したポイント (Key Challenges)

### 1. スプレッドシートの「数式解析エラー」の回避
当初は1つのセルですべての処理を行おうとしましたが、URLが長くなりすぎるとスプレッドシートが数式として認識しない（テキスト扱いになる）問題が発生しました。これに対し、**URL生成ロジックとHYPERLINK関数を別のセルに分割**することで解決しました。

### 2. Salesforce Lightningのパラメータ仕様への対応
Lightning Experienceでは `defaultFieldValues` パラメータを使用しますが、値の中にカンマが含まれると区切り文字として誤認されます。これに対し、**商談名などを `ENCODEURL` でエンコード**し、かつデータ整形を行うことで誤作動を防ぎました。

### 3. 表記ゆれの吸収
元データの企業名に「1 株式会社〇〇」のように管理番号が含まれており、そのままVLOOKUP等で使用できませんでした。
`REGEXREPLACE(B4, "^[0-9]+ ", "")` という正規表現を使用し、**先頭の数字とスペースのみを動的に削除**するロジックを組み込みました。

### 4. Salesforce IDの桁数仕様への対応（15桁 vs 18桁）
SalesforceのレコードIDには「18桁」と「15桁」の2種類が存在しますが、URLパラメータで値を渡す際、18桁のIDを使用すると内部エラーが発生するケースを確認しました。
これに対し、**`LEFT(AM4, 15)` 関数を使用して強制的に15桁に変換するロジック**を実装し、運用ミスによるエラーをシステム側で防いでいます。

## 🚀 成果 (Impact)
* **工数削減:** 1件あたりの作業時間を **5分 → 30秒** に短縮（約90%削減）。
* **品質向上:** 手入力による転記ミスが **0件** に。
* **属人化の解消:** 複雑な入力ルールを数式内に隠蔽したことで、マニュアルなしでも誰でも正確に案件作成が可能になった。
