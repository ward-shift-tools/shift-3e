# 3E版 開発引き継ぎドキュメント

Beta版完成時点（2026-04-17 / commit `e410178` / タグ `v1.0-beta`）からの派生開発用。

---

## 🎯 現状サマリー

### Beta版でできること
ICU看護師シフトを以下の制約下で自動生成:
- 5段階Tier制（A/AB/B/C+/C）
- ペア制約（B+C族禁止、C族≤1、AB代替ペナルティ）
- 日勤/夜勤リーダー確保、新人夜勤時A必須
- 日勤最低人数（全体／新人除く）分割管理
- 休暇・研修・委員会・夜勤不可などの希望処理
- 公休日数、連勤、前月繰越の自動処理

### 技術スタック
| 層 | 技術 |
|---|---|
| 最適化 | PuLP + HiGHS (MILP) |
| UI | Streamlit |
| I/O | openpyxl (xlsx) + gspread (Google Sheets) |
| デプロイ | Streamlit Cloud ← GitHub main branch |

---

## 📂 格納場所・アクセス

| 項目 | 場所 |
|---|---|
| ローカル作業ディレクトリ | `/Users/yoshikawayusuke/ShiftManagementTool/shift-icu/` |
| 親フォルダ（複数部署の目次） | `/Users/yoshikawayusuke/ShiftManagementTool/` |
| ICU版 GitHub | `https://github.com/ushichichi1/shift-icu`（旧`shift-scheduler`をリネーム予定） |
| 3E版 GitHub | `https://github.com/ushichichi1/shift-3e`（**新規作成が必要**） |
| GitHubコミッタ | 吉川祐輔 |
| Streamlit Cloud | リポジトリ別にアプリを作成（ICU用・3E用で独立デプロイ） |

### 重要ファイル（ICU版）
```
ShiftManagementTool/                  ← 親フォルダ（部署目次）
├── README.md                         ← 全部署のインデックス
│
└── shift-icu/                        ← ICU版（独立リポ: shift-icu）
    ├── shift_scheduler.py            ... ソルバー本体（約1900行）
    ├── app.py                        ... Streamlit UI（約2300行）
    ├── create_test_data.py           ... テストデータ生成
    ├── SPEC.md                       ... 仕様書（ICU版・Beta完成時点）
    ├── HANDOVER_3E.md                ... ← このファイル
    ├── README.md                     ... ICU版の起動・概要
    ├── requirements.txt              ... 依存関係
    ├── credentials.json              ... GSheet認証（.gitignoreで除外）
    └── .gitignore
```

### 3E版の配置予定
```
ShiftManagementTool/
└── shift-3e/                         ← 3E版（独立リポ: shift-3e、新規作成）
    ├── ... (ICU版をベースにカスタマイズ)
    └── ...
```

---

## 🔀 3E版を派生させる手順（方針決定済: 別リポジトリ独立）

「システムや制約周りが大幅に刷新される」という方針のため、**別リポジトリ完全独立**で開発する。

### ステップ
```bash
cd /Users/yoshikawayusuke/ShiftManagementTool

# 1. ICU版をテンプレートとしてコピー
cp -r shift-icu shift-3e
cd shift-3e

# 2. 独立リポとして git を初期化し直し
rm -rf .git
git init
git add -A
git commit -m "Initial commit from shift-icu v1.0-beta template"

# 3. GitHub で新規リポジトリ "shift-3e" を作成後
git remote add origin https://github.com/ushichichi1/shift-3e.git
git branch -M main
git push -u origin main

# 4. 3E固有の制約に改修開始
# 5. Streamlit Cloud で新規アプリとしてデプロイ
```

### 親フォルダREADME更新
`ShiftManagementTool/README.md` の「部署別バージョン」表に `3E` のエントリを追加して、状態・フォルダ・GitHub・Streamlit Cloud URL を記載。

---

## 🧩 3E版で検討すべき差分ポイント

"3E" が意味する内容に応じて、以下のどれが該当するか次セッションで明確化:

### 候補A: 3交代制（日勤・準夜・深夜）
現在は2交代制ベース。3交代にするなら:
- シフト定数追加: 準夜（既存の`SN`を拡張？）、深夜
- 夜勤時間数の再定義（現状16h/夜勤 → 8h×2交代）
- 72h規制、連勤制約の再計算
- 「明」の扱い（3交代では深夜明け＝早朝終業）
- Tier別の配置ルール

**改修範囲**: 中〜大（定数・制約の再設計、テンプレート大幅修正）

### 候補B: 3E病棟（特定の一般病棟）
ICUと異なる病棟ルールを適用:
- リーダー配置基準の緩和/変更（ICUはA単独可、他病棟はAB以上で可、など）
- 夜勤人数（ICU2名→他病棟Nn）
- Tier定義の調整
- 重症度/看護度の加味

**改修範囲**: 小〜中（定数変更と一部制約の調整）

### 候補C: 3病棟統合管理
複数病棟を1モデルで扱う：
- 病棟間応援シフト
- 病棟ごとの制約プロファイル

**改修範囲**: 大（データモデルから再設計）

---

## ⚠️ 改修時の注意点（Beta版で固まった設計）

### 変更しにくい箇所（理由あり）
- **Tier数は5段階に最適化済み** → 増減するなら`VALID_TIERS`とすべてのペア判定式を見直し
- **夜勤ペア制約は相互関連** → 一部だけ変えると整合性破綻（SPEC.md §2.1 参照）
- **公休(O)・明(A)・休暇(V)の関係** → 公休計上は O のみ、A/V は別枠という運用慣習に依存

### 変更が容易な箇所
- **設定値**（最低人数、夜勤上限、連勤等）: `SETTINGS_DEF` で即変更可
- **ペナルティ重み**: `build_and_solve` の目的関数近辺（line 1600付近）
- **シフト種別の有効/無効**: `enabled_shifts` で選択可
- **色・UI**: `FS`, `FT` 等の dict

### 既知の設計判断
- 明(A)は公休(O)にカウントしない（運用慣習）
- C族同日≤1はハード（経験不足ペア防止）
- 希望日数は月7日まで（入力量制御）
- 委員会は希望指定のみ（ソルバー自由割当しない）

### 既知の落とし穴（Beta版で踏んで対応済み）
| 問題 | 対応 | 教訓 |
|---|---|---|
| `value or ""` で 0 が "" 化 | None判定に全変更 | 0/None を混同しない |
| V=1とO=1の二重強制 → Infeasible | 曜日制限で休暇日スキップ | 複数ハード制約の衝突に注意 |
| pandas.apply の型混入 | iterrows + str()化 | Streamlit Cloud環境での型挙動注意 |

---

## 🔑 新セッション開始時のプロンプトテンプレート

次セッションでは、以下を冒頭で共有するとスムーズ：

```
親フォルダ: /Users/yoshikawayusuke/ShiftManagementTool
ICU版（テンプレート）: /Users/yoshikawayusuke/ShiftManagementTool/shift-icu (git: ushichichi1/shift-icu)
3E版を新規作成: /Users/yoshikawayusuke/ShiftManagementTool/shift-3e （別リポ）
ICU版完成タグ: v1.0-beta

ICU版をテンプレートに、3E版を別リポジトリとして新規開発したい。
まず shift-icu/HANDOVER_3E.md と shift-icu/SPEC.md を読んで現状を把握して。
その上で 3E版でこう変えたい: [具体的な差分を記述]
```

## 🔮 将来のTOPハブ構想

ユーザーが最初にアクセスするTOP画面で、部署を選択して該当アプリへ遷移するハブを将来的に作成予定。
- 1つのURL（例: `shift-management.streamlit.app`）からICU / 3E / ERなどを選択
- 選択した部署の専用アプリへ遷移
- 将来的には **部署横断の応援シフト最適化**（統合モデル）も視野

このハブは `shift-management-hub` などの新規リポジトリとして別途作成する想定。各部署ツールは引き続き独立リポとして維持。

### 最初にやるべきこと（次セッション）
1. `HANDOVER_3E.md`（このファイル）と `SPEC.md` を読む
2. 3E版の「3E」が何を指すか確認（候補A/B/Cのどれか）
3. 差分の大きさを評価 → ブランチ分岐 or 別リポジトリ判断
4. `v1.0-beta` からブランチ切って着手

---

## 📋 Beta版完成時点の状態

### 現在の機能一覧
- [x] 5-Tier制（A/AB/B/C+/C）
- [x] 夜勤ペアリング制約（ハード／ソフト複合）
- [x] 日勤/夜勤リーダー配置
- [x] 新人除き日勤最低人数
- [x] 委員会シフト（A/AB同席必須）
- [x] パートタイム・時短・夜勤専従対応
- [x] 夜勤研修（3人目枠）
- [x] 前月繰越、連勤制限、公休数管理
- [x] 曜日/祝日/土日勤務制限
- [x] 希望記号（日/夜/休/研/委/夜不/休暇/明休）
- [x] 読み込み結果サマリープレビュー
- [x] 論理エラー検出
- [x] Excel/Google Sheetsテンプレート
- [x] Streamlit Cloud デプロイ

### 未対応（Betaの範囲外）
- [ ] 複数月連続生成（月跨ぎ調整）
- [ ] 希望の優先度（絶対休 vs できれば休）
- [ ] スタッフ間相性
- [ ] 看護度・重症度別配置
- [ ] モバイル入力フォーム
- [ ] 3交代制
- [ ] 複数病棟統合

---

## 📞 不明点・困ったとき

- **仕様の詳細**: `SPEC.md` 参照
- **制約の根拠**: `shift_scheduler.py::build_and_solve()` のコメント
- **UI挙動**: `app.py` の該当セクション（機能名でgrep）
- **過去の判断理由**: コミットメッセージ（`git log`）

---

以上。Beta版お疲れ様でした。3E版の方向性が決まり次第、次セッションへ引き継ぎを。
