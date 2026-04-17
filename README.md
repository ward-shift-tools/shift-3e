# Shift Management Tool — ICU版

ICU看護師シフト自動生成システム（2交代制・5段階Tier）

**バージョン**: v1.0-beta（2026-04-17）  
**ステータス**: ✅ Beta完成・本番運用可能

---

## 🚀 起動方法

### ローカル実行
```bash
pip install -r ../requirements.txt  # 依存は親フォルダの共通ファイル（または自フォルダに配置）
streamlit run app.py
```

### Streamlit Cloud
- GitHubリポジトリ: `ushichichi1/shift-icu`（旧 `shift-scheduler`）
- Main file path: `app.py`
- URL: Streamlit Cloud管理画面参照

---

## 📁 ファイル構成

| ファイル | 役割 |
|---|---|
| `app.py` | Streamlit UI（アップロード・プレビュー・生成・結果表示） |
| `shift_scheduler.py` | MILPソルバー本体（PuLP + HiGHS） |
| `create_test_data.py` | テスト用xlsx生成スクリプト |
| `SPEC.md` | **完全仕様書（12章構成）** |
| `HANDOVER_3E.md` | 3E版派生のための引き継ぎドキュメント |
| `requirements.txt` | 依存ライブラリ |

---

## 🏥 ICU版の特徴

### Tier制度（5段階）
- **A**: ベテラン・リーダー格（日勤/夜勤リーダー単独可）
- **AB**: 中堅・リーダー代行可（夜勤リーダー可）
- **B**: 一人立ち済み（B+B, B+C族の夜勤ペア禁止）
- **C+**: C既卒（A/AB/B下で夜勤可）
- **C**: 新人・経験浅い（必ずA/AB/Bと夜勤ペア）

### 主要制約
- 夜勤は2人/日（研修時は3人）
- 日勤最低5人（新人除き4人）
- A夜勤リーダー＋日勤リーダー両方必須
- 新人夜勤時はA必須
- 委員会（委）指定時はA/AB同席必須

詳細は [`SPEC.md`](./SPEC.md) 参照。

---

## 🛠️ 主な機能

- 📋 読み込み結果サマリー（属性別グループ表示）
- 🔍 論理エラー自動検出
- 📊 Excelテンプレート / Googleスプレッドシート入出力
- 🎲 複数パターン生成
- ⚙️ 柔軟な設定（公休日数、夜勤上限、連勤等）
- 🏛️ 委員会バックアップ配置

---

## 🔧 トラブルシューティング

- **Infeasible（解なし）が出る** → [`SPEC.md` §9.1](./SPEC.md) のチェックリスト参照
- **読み込みで値が消える** → 読み込みサマリーパネルで確認
- **Streamlit Cloudでエラー** → Manage app → Logs を確認

---

## 📞 関連リポジトリ

- 親フォルダREADME: [`../README.md`](../README.md)
- 将来: `shift-3e/`（3E版、別リポジトリ）
