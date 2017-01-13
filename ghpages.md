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
tagline: Excel VBAを使いこなそう
google_analytics: XXXXXXX
theme: jekyll-theme-cayman

markdown: kramdown
kramdown:
  input: GFM
  syntax_highlighter: rouge
```

<https://github.com/pietromenna/jekyll-cayman-theme> より次のファイルをコピー

* _layouts/default.html  レイアウトファイル 
* _includes/head.html  htmlヘッダ、 highlight.cssを追加
* _includes/page-header.html  微調整
* _includes/page-footer.html 微調整
* css/cayman.css 微調整
* css/normalize.css
* css/highlight.css GhPagesのCSSから.highlight をコピー(もっといい方法があるはず)


### コードハイライタで装飾できる言語種類
Github Pagesは [Rouge](https://github.com/jneen/rouge)でコードハイライトできる

* html
* shell
* vb
* yaml

参考) <https://github.com/jneen/rouge/wiki/List-of-supported-languages-and-lexers>

--------------------------

## ローカルで動作




--------------------------

## gitの利用メモ

