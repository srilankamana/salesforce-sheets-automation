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
