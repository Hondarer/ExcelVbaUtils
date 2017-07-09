# ExcelChildProcessController

Microsoft Excel でコマンドラインアプリケーションを呼び出して制御するためのライブラリです。

## どんなことができるの

+ Microsoft Excel から、コマンドラインアプリケーションを起動して制御することができます。
+ 戻り値のみならず、標準入力に入力を行ったり、標準出力や標準エラー出力の文字列を処理することができます。
+ 対話型アプリケーションのプロンプトを待ち合わせながら、処理を進めることができます。
+ モジュール設計のため、Microsoft Excel 特有の、「おのおのが拡張して破綻する VBA 沼」にはまることを避けられる可能性があります。※確実な回避には統制と教育が必要です。
+ 現段階で最も普及していると思われるマクロ実行環境 Microsoft Excel x86 から、.NET のコンソールアプリや Oracle SQL*Plus x64 を起動し制御できるため、自動化の効率化に役立ちます。これは、x64 環境下で Microsoft Excel x86 のためだけに Oracle x86 ODBC driver をインストールしなくてもよいことを意味します。
+ VBA でのクラス志向やコールバック処理、Win32 API に関する P/Invoke のサンプルとしても活用できます。

## 動作環境

- Microsoft Windows(システムロケールが cp932:ja-JP.sjis であることが前提となっています。)
- Microsoft Excel x86

## サンプル説明

### SampleFunction1

Windows コマンドラインプロセッサ(cmd.exe)を起動し、コマンドを一気に流し込むサンプルです。

### SampleFunction2

Windows コマンドラインプロセッサ(cmd.exe)を起動し、プロンプトを確認しながらコマンドを与えていくサンプルです。

### SimpleQueryFunction

Oracle SQL*Plus を起動し、クエリを与え、結果をシートに書き込むサンプルです。

Orcale 11g XE 付属のサンプル DB を参照しています。
