# bacteriology

R を用いた細菌学（bacteriology）データ解析プロジェクトです。

## 必要環境

- [R](https://www.r-project.org/) (バージョン 4.0 以上推奨)
- [RStudio](https://posit.co/download/rstudio-desktop/)（任意・推奨）

## セットアップ

### 1. リポジトリのクローン

```bash
git clone <リポジトリURL>
cd bacteriology
```

### 2. 必要パッケージのインストール

R コンソールで以下を実行してください：

```r
# 依存パッケージのインストール（例）
install.packages(c("tidyverse", "ggplot2"))
```

## 使い方

### RStudio を使う場合

1. RStudio を起動する
2. `File` → `Open Project` → プロジェクト内の `.Rproj` ファイルを選択する
3. `Files` ペインからスクリプト（`.R` ファイル）を開いて実行する

### R コンソール / ターミナルを使う場合

```bash
# ターミナルから直接スクリプトを実行
Rscript your_script.R
```

または R コンソール内で：

```r
source("your_script.R")
```

## プロジェクト構成

```
bacteriology/
├── R/            # R スクリプト・関数定義
├── data/         # 入力データ
├── output/       # 解析結果・図表
├── vignettes/    # 使用例・チュートリアル（R Markdown）
└── README.md     # このファイル
```

## ライセンス

このプロジェクトのライセンスについては、リポジトリ管理者にお問い合わせください。

## 問い合わせ

ご不明な点は [Issues](../../issues) からお問い合わせください。
