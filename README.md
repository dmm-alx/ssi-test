# dmm-ssi-Glogs-Python

## Overview

株式会社DMM少額短期保険のGoogleドライブの監査ログの抽出するスクリプト

## Description

* Cloud Functions (Python)環境
* Cloud Scheduler設定： */10 9,10 * * *
* ソースコードをGithubで管理して実行はGoole Cloudで行う
* コードをGoogle Cloud RepositoryでGoogle Cloud Functionsにpushして実行する　

## Requirements

* Cloud Functions (Python)環境
* 抽出対象のドライブのリストシートを用意する（
    - main.pyの’SPREADSHEET_ID’にログ抽出対象ドライブのリストが載ったSpreadsheetのIDをいれる
    - main.pyの’SHEET_NAME’にいれる
    - Spreadsheetのサンプルは　`src/ShareDriveList.xlsx`　のしたにあります。
        * A列　にドライブID（必須）
        * B列　にドライブ名（自動生成）
        * C列　にログのデータ記入シートID（自動生成）
        * D列　に実行時間（自動生成）
        * E列　に実行ステータス（自動生成）
        * F列　に実行ステータス　Numeric（自動生成）
        * H列　にエラーメッセージ（自動生成）

* ログ抽出ごに格納するフォルダーを’FOLDER_ID’に設定する

## Usage

# Setup
Google側で毎日実行を行うトリガーを設定する必要がある。

'''
https://asia-northeast1-be-alexis.cloudfunctions.net/main
'''

現在の実行時間：朝９と10時の10分おき

# Coding
1. GitリポジトリをローカルにCloneする
2. 作成させれましたリポジトリに対してサービスアカウントのCredentials（自身が取得したもの）を設定する
3. 各テスト用のドライブ、シート（SPREADSHEET_ID、SHEET_NAME、FOLDER_ID）
4. 問題解決したら、Githubにpushしてソースコードを管理する。
5. GithubとGoogle reposityミラリングされているのでGoogle reposityが更新される。
6. Google functionsでGoogle reposityからソースコードを読み込んでデプロイする。
7. 本番で問題ないことモニタリングする。

# Caution
実行結果をDrive list SpreadSheetに書き込んでいる　（F列）
* 1　の場合処理が正常に実行が最後まで行った。ループが回ってきた時に次のドライブに移る。
* 2　の場合実行が途中。次の実行でそのシートの最終行の時間を元にデータを取得する。
* 現在の設定で2時間回ってもステータス２があればコード確認する

## Contribution

## License

## Author
Nanou Alexis
