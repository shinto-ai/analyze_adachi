# データ分析プログラム README

このプログラムは、指定されたディレクトリ(dataディレクトリ)内の複数のExcelファイルを読み込み、各ファイルのデータに対してカイ二乗検定を行います。
さらに、指定された試行回数分、各ファイルのシートをランダムに振り分けた場合のカイ二乗値を求め、これらを比較することで統計的有意性を評価します。

以下の「データの準備」と「プログラムの実行方法」を読めば、本プログラムを利用できます。
それ以外の詳細については、適宜参照してください。
何か不明点があれば、下記の連絡先までご連絡ください。


## データの準備

1. 分析対象のExcelファイル(`.xlsx` or `.xls`)を`data`フォルダに配置します。
2. Excelファイルは**1つ以上**配置可能です。ファイル名は任意です。
3. 各Excelファイル内には、以下の条件を満たすデータが含まれている必要があります。
   - 各データがシートごとに分割されていること
   - データは**数値のみ**で構成されていること
   - 比較対象のデータ同士で、**行列の数が等しい**こと
   - **シートの1行目と1列目は削除**されます（ヘッダー行・列として使用されないため）
   - シート内の**空欄のセルは0**として扱われます

`archive`フォルダには、本プログラムの動作確認のために利用したdemoファイルが含まれています。
適宜参考にしてください。

- HU_data : 安達先生から頂いたデータ。
- Otani_data : 安達先生から頂いたデータ。
- demo_data : 大きく偏らせた場合、正確に有意差を捉えられているかを確認。また、ファイルが3つ以上の場合の動作を確認。
- demo_row1 : 行数が1の場合の動作確認。行数が1の場合、各セルの期待度数が0となるので、カイ二乗値も0になる。



## プログラムの実行方法

プログラムの実行方法は非常にシンプルです。
プログラムが配置されているディレクトリに移動し、仮想環境を有効化します。(仮想環境については「作業内容の記録」にて簡単に説明。)
その後、プログラム実行のコマンドを入力します。

### コマンドラインからの実行

1. ターミナルまたはコマンドプロンプトを開き、プログラムが配置されているディレクトリに移動します。

   ```bash
   cd プロジェクトフォルダ
   ```

2. 仮想環境を有効化します。以下のコマンドを入力して下さい。

    ```bash
    # 仮想環境の有効化（Windowsの場合）
    venv\Scripts\activate

    # 仮想環境の有効化（Unix/Linux/macOSの場合）
    source venv/bin/activate
    ```

3. プログラムを実行します。

   ```bash
   python analyze_data.py
   ```

### プログラムの流れ

1. プログラムを起動すると、`data`フォルダ内のExcelファイルが読み込まれます。
2. 各ファイルの有効なシートデータが取得されます。(全て空欄のシートなどはスキップされます。)
3. 各ファイルごとにクロス集計表の作成が行われ、`output`フォルダに保存されます。
4. クロス集計表から各セルの期待度数を求められ、`output`フォルダに保存されます。
5. 元のデータとその期待度数から、行ごとにカイ二乗値を求め、それを全て合計することでそのファイル全体のカイ二乗値とします。
6. さらに、各ファイルごとのカイ二乗値を合計することで、実験データ全体のカイ二乗値とします。

7. 統計的優位性評価のため、全てのファイルのシートを混ぜ、ランダムに振り分けます。(ファイル内のシート数に基づいて振り分けられます。)
8. ランダムに振り分けたシートに対して、上記2~6を行います。
9. 元のデータ全体のカイ二乗値と、ランダムに振り分けたデータ全体のカイ二乗値を比較し、前者が大きい回数を数えます。
10. 入力された試行回数分、ランダム化検定を行い、最終的に元のデータ全体のカイ二乗値が大きかった割合を出力します。

### 試行回数の入力

- プログラム実行中に、ランダム化検定の試行回数を入力するプロンプトが表示されます。
- **正の整数値**を入力してください。
- 試行回数が大きいほど結果の精度は上がりますが、処理時間も長くなります。

## 出力結果

- `output`フォルダ内に、各ファイルの**クロス集計表**がExcel形式で保存されます。
- プログラムの実行結果として、以下の情報が表示されます。
  - **処理時間**
  - **試行回数**
  - **元のカイ二乗値**
  - **元のカイ二乗値の方が大きい回数**
  - **元のカイ二乗値の方が大きい割合**

## エラーハンドリング

- プログラムは、以下のエラーに対して適切なメッセージを表示します。
  - `data`フォルダが存在しない場合
  - Excelファイルが見つからない場合
  - データが不適切な場合（空のシート、数値以外のデータなど）
  - シート間でデータサイズが一致しない場合


## 連絡先

ご質問や不明点がある場合は、開発者までお気軽にお問い合わせください。

- **開発者**: 進藤 稜真
- **作成日**: 2024年11月11日
- **Eメール**: shinto.ryoma@gmail.com


## 作業内容の記録

環境構築にあたって、本PC内で行った内容は以下の通りです。

## 環境設定

### 1. Pythonのインストール

Python 3.10.11 をインストールしました。

- [Python公式サイト](https://www.python.org/downloads/)

### 2. 仮想環境の作成

仮想環境を作成しました。
仮想環境を使用すると、プロジェクトごとにパッケージの依存関係を管理できます。
すなわち、本PCで他のプログラムを実行するためにパッケージのバージョンを変更しても、この仮想環境内のバージョンは固定されたままになります。
よって、仮想環境を有効化してプログラムを実行することで、パッケージのバージョンの不整合による不具合は起きません。

```bash
# 仮想環境の作成
python -m venv venv

# 仮想環境の有効化（Windowsの場合）
venv\Scripts\activate

# 仮想環境の有効化（Unix/Linux/macOSの場合）
source venv/bin/activate
```

### 3. 必要なパッケージのインストール

以下のコマンドを使用して、必要なパッケージをインストールします。

```bash
pip install pandas numpy openpyxl
```

## プログラムの配置

1. プログラムファイル（例: `analyze_data.py`）を任意のディレクトリに配置します。
2. 同じディレクトリ内に`data`フォルダを作成します。

```
プロジェクトフォルダ/
├── analyze_data.py
└── data/
```