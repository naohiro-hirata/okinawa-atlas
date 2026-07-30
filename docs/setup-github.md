# Phase 0 — GitHubで公開するまで（GitHub未経験向け）

ターミナルもPythonも使いません。すべてアプリとブラウザのクリック操作です。
所要 30〜40分。うち待ち時間が10分ほど。

終わると `https://（あなたのアカウント名）.github.io/okinawa-atlas/` でサイトが見られる状態になります。
この時点で施策・空き家データは実データ、人口はダミー表示のままです。

---

## 先に知っておくこと

**このリポジトリは「公開（public）」にします。** GitHub Pages の無料枠は公開リポジトリだけが対象で、
非公開にすると有料プランが必要になります。今回のデータに個人情報は含まれないので問題ありませんが、
**この先このフォルダに個人情報を含むファイルを置かないでください。** 置いた瞬間に全世界から見えます。

用語は3つだけ覚えれば足ります。

| 用語 | 意味 |
|---|---|
| リポジトリ | プロジェクト1個ぶんの保管庫。今回は `okinawa-atlas` |
| コミット | 変更に名前を付けて記録すること。手元の作業 |
| プッシュ | コミットをGitHub（クラウド）に送ること。これで公開される |

---

## ステップ1 — フォルダを置く（3分）

ダウンロードした `okinawa-atlas` を、わかりやすい場所に置きます。

- Windows: `C:\Users\（ユーザー名）\Documents\okinawa-atlas`
- Mac: `/Users/（ユーザー名）/Documents/okinawa-atlas`

zipで届いていたら展開してください。

**確認:** フォルダを開いて、`README.md` `CLAUDE.md` `scripts` `app` `public` が
**すぐ目に入る**こと。`okinawa-atlas` の中にもう1つ `okinawa-atlas` があったら、
内側のフォルダの中身を外に出してください。二重になっていると後で動きません。

日本語やスペースを含むフォルダ名の下に置くのは避けてください（`デスクトップ` は問題ありません）。

---

## ステップ2 — GitHubアカウントを作る（5分）

> **すでにアカウントを持っている場合はこのステップを飛ばして、ステップ3へ進んでください。**

1. <https://github.com/signup> を開く
2. メールアドレス、パスワード、ユーザー名を入れる
3. **ユーザー名は公開URLの一部になります。** `okinawa-lab` のようにすると
   `https://okinawa-lab.github.io/okinawa-atlas/` になります。後から変えると面倒なので、
   社名や事業名にしておくのが無難です
4. メールに届いた数字コードを入力して認証
5. プランを聞かれたら **Free** を選ぶ

すでにアカウントがあるならこのステップは飛ばしてください。

---

## ステップ3 — GitHub Desktop を入れる（5分）

1. <https://desktop.github.com/> を開いて、自分のOSのボタンを押す
   （Windowsはインストーラ、Macはzip。Macはzipを展開して、出てきたアプリを「アプリケーション」に移動）
2. 起動して **Sign in to GitHub.com** → ブラウザが開くので、ステップ2のアカウントで許可する
3. 名前とメールの確認画面が出たらそのまま **Finish**

---

## ステップ4 — フォルダをリポジトリにする（5分）

1. GitHub Desktop のメニュー **File → Add local repository…**
2. **Choose…** でステップ1のフォルダを選ぶ
3. 「This directory does not appear to be a Git repository」と赤い文字が出ます。**これは正常です。**
   その文中の **create a repository** というリンクを押す
4. 出てきた画面で以下を確認して **Create repository**
   - Name: `okinawa-atlas`
   - Git ignore: **None**（フォルダに `.gitignore` が同梱されているので、ここで選ぶと上書きされます）
   - License: None のまま

これで手元の記録が始まりました。まだGitHubには何も送られていません。

---

## ステップ5 — GitHubに送る（5分）

1. 画面上部の **Publish repository** を押す
2. ダイアログで **`Keep this code private` のチェックを外す**

   > **ここが最重要です。** チェックが入ったままだと非公開リポジトリになり、
   > 無料プランではGitHub Pagesが使えず、この先の手順が全部失敗します。

3. **Publish repository** を押す

数十秒でアップロードが終わります。ブラウザで
`https://github.com/（アカウント名）/okinawa-atlas` を開くと、ファイルが並んでいるはずです。

---

## ステップ6 — 公開設定をONにする（3分）

ここまでではまだサイトになっていません。ブラウザのリポジトリ画面で操作します。

1. 上のタブの **Settings**（歯車）を押す
2. 左のメニューを下にスクロールして **Pages** を押す
3. **Build and deployment** の **Source** を、`Deploy from a branch` から
   **`GitHub Actions`** に変える

これで、プッシュするたびに自動でビルドして公開する設定になりました。

---

## ステップ7 — ビルドを走らせて確認する（10分）

設定より先にファイルを送ったので、1回だけ手で走らせます。

1. 上のタブの **Actions** を押す
2. 左に **Build and deploy to GitHub Pages** があるので押す
3. 一覧の中の実行を1つ押す
4. 右上の **Re-run all jobs** → **Re-run jobs**

3〜5分ほどで、緑のチェックが2つ（build と deploy）付きます。

**確認:** Settings → Pages に戻ると、上部に公開URLが出ています。
`https://（アカウント名）.github.io/okinawa-atlas/` を開いて、
左に41市町村のリスト、中央に花ブロックの一覧が出れば成功です。

---

## つまずいたときの読み方

Actions で **赤い×** が付いた場合、それは失敗したという意味です。押すと黒い画面にログが出ます。
**そのログを最後の20行くらいコピーして、私に貼ってください。** 原因を特定します。

よくある3つ:

| 症状 | 原因 | 対処 |
|---|---|---|
| Pages の Source に `GitHub Actions` が選べない | リポジトリが非公開になっている | Settings → 一番下の Danger Zone → Change visibility → Public |
| build が緑なのに deploy が赤 | ステップ6より前の実行を見ている | Actions で Re-run all jobs をやり直す |
| `No such file or directory: data/raw/survey/latest.xlsx` | ステップ1のフォルダが二重になっている | 内側のフォルダの中身を1つ上に出して、GitHub Desktop で Commit → Push |

---

## これ以降の更新のしかた

手元のフォルダのファイルを変えたら、毎回この3操作です。

1. GitHub Desktop を開く（変更したファイルが左に一覧で出ます）
2. 左下の **Summary** に何をしたか一言書いて **Commit to main**
3. 上の **Push origin** を押す

3〜5分後に公開サイトが更新されます。来年の調査Excelを差し替えるときも、
`data/raw/survey/latest.xlsx` を新しいものに置き換えて、この3操作だけで反映されます。

---

## Phase 0 が終わったら

ここまでで、GitHubの操作（コミットとプッシュ）は覚えたことになります。
Phase 1 以降はExcelパーサの調整で、試行錯誤の反復が必要になるため Claude Code を使います。
Windows での準備は [setup-windows.md](setup-windows.md)、そのあと [handoff.md](handoff.md) の Phase 1 へ。
