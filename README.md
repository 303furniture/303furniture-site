# 303furniture-site
303furniture GitHub Pages site

## 正規URL方針

### ドメイン正規化
- **正規ドメイン**: `https://www.303furniture.net`
- **非www → www**: 301リダイレクトで統一

### URL正規化
- **拡張子あり**: `.html`付きURLを正規とする
- **拡張子なし**: `/product` → `/product.html` へ301リダイレクト
- **正規URL一覧**:
  - ホーム: `https://www.303furniture.net/`
  - 商品: `https://www.303furniture.net/product.html`
  - 家具: `https://www.303furniture.net/furniture.html`

### .htaccess設定
- ドメイン正規化（非www → www）
- URL正規化（拡張子なし → 拡張子あり）
- セキュリティヘッダー設定
- キャッシュ・圧縮設定

### SEO設定
- canonicalタグ: 各ページで正規URLを指定
- sitemap.xml: 正規URLのみ含む
- robots.txt: 適切なクローリング設定
- OGP設定: 正規URLで統一