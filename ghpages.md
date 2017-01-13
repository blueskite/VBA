# GitHub Pagesを利用した静的サイト構築

このサイトはGithub Pages上でMarkdown (GFM)を利用したWikiページとしています。

## ワークフロー

```shell 
## 最初
$ git clone https://github.com/blueskite/VBA.git
$ cd VBA
～編集
$ git add *; git commit -m "edited"; git push

## ファイルを削除した時
$ git rm ファイル名


## 他で編集したものの取込
$ git pull


```


## ghpages のカスタマイズ
設定ファイル _config.yml

```yaml
title: VBA
description: Excel VBAを使いこなそう
# show_downloads: true
google_analytics: XXXXXXX
theme: jekyll-theme-cayman
```





### コードハイライタで装飾できる言語種類
Github Pagesは [Rouge](https://github.com/jneen/rouge)でコードハイライトできる。

* html
* shell
* vb
* yaml

参考) https://github.com/jneen/rouge/wiki/List-of-supported-languages-and-lexers




## ローカルで動作



## gitの利用メモ

