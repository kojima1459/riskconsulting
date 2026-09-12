# bench/ - S1抽出品質ベンチ(裁定書37・班F・gate外・任意実行)

PO の最大懸念は「社内 Deep Research の再現性。プロンプトの指示通りに求める情報を
過不足なく正確に抜き出すか。ハルシネーションで変な情報を取ってきたらレポートは
台無し」というものである。ここは **DR品質(社内DRアプリそのものの精度)** と
**S1抽出品質(本製品のS1が、DR出力から正解フィールドを取れるか・捏造しないか)**
を分けて測る(混ぜない。伝書鳩1-5)。**このベンチは `tools/gate.py` の21ゲートには
含まれない**(LLM実行を伴う指標をCIの合否にしないため)。

## 使い方

```bash
# 1) 採点器自体の回帰(mockのS1応答を抽出して算術を検算する。AI不要)
python3 tools/bench_s1.py --mode selfcheck

# 2) gold の嘘を機械で止める(forbiddenが本当に素材に無いか・recallが本当にあるか)
python3 tools/bench_s1.py --gold-check

# 3) 実際のS1出力を採点する(replay。AI不要・保存済みJSONを読むだけ)
python3 tools/bench_s1.py --mode replay --in bench/out/live
```

`--mode replay` は既定で `bench/out/live/` を読む。`--in` に単一ファイルや別ディレクトリ
を渡してもよい。結果は標準出力の表と `bench/out/summary.json` の両方に出る。

## 出す数値(4つだけ。定義はここが正。`tools/bench_s1.py` はこれと一字一句合わせる)

| 指標 | 定義 |
|---|---|
| **一致率(exact)** | gold の `exact` フィールド(例: `company_name`・`financials.source`)の一致数 / exact フィールド総数(複数回replayした場合は回数ぶん積み上げる) |
| **回収率(recall)** | gold の `recall` 断片(必ず拾うべき事実の短い断片)のうち、S1出力のJSON全文に部分一致で現れた数 / recall断片総数 |
| **捏造率(forbidden)** | gold の `forbidden` 断片(素材に無い=書かれていたら捏造と判定する断片)のうち、S1出力に現れてしまった数 / forbidden断片総数。**0でなければ赤** |
| **揺れ(variance)** | 同一companyでS1出力を複数回(2回以上)与えたとき、`exact` フィールドの値が全回で同値だった割合。1回しか出力が無いcompanyはこの計算に入れない |

補足で **出典実在率(sources)**: S1出力に `sources`(班Aが同時並行で追加中のキー。
15章の未確定分)があれば、各 `url` が該当companyの `bench/fixtures/<company>/` 本文に
部分一致で存在するかを見る。`sources` が1件も出てこない出力だけなら `n/a`。

**表記揺れ耐性**: 全角半角・空白・句読点・鍵括弧・中黒・カンマ・ピリオド・長音は
正規化してから比較する。規則は `src/app/modGround.bas` の `NormalizeForMatch` と
**同じもの**(値源はそちら1本。ズレたら `tools/bench_s1.py` の
`normalize_for_match` をそちらに合わせて直す)。断片の一致は先頭何字などの
打ち切りをせず**全文一致**で見る(recall/forbiddenの断片は短い固定文字列なので、
`modGround.QuoteFound` の headChars 切詰めは持ち込まない)。

**採点器は「未知キーを無視」する。** 班Aが `sources[]` と `missing_info[].kind` を
15章へ追加中(裁定書37時点で未確定)なので、採点器はS1出力の知らないキーに触れても
エラーにしない(スキーマ検証はしない。gold の exact/recall/forbidden にある項目
だけを見る)。

## gold の作り方(推測で足さない)

`bench/gold/<company>.json`:

```json
{
  "company": "有限会社春華堂",
  "fixture_dir": "shunkado",
  "exact": {"company_name": "有限会社春華堂", "financials.source": "unknown"},
  "recall": ["春華堂", "うなぎパイ", "1887年", "..."],
  "forbidden": ["浸水深30cm~50cm未満", "..."]
}
```

- `exact`: ドット区切りのJSONパス -> 期待値(スカラーのみ。配列は対象外)。
- `recall`: **fixtures に実在する**短い事実の断片。gold-check が実在を検算する。
- `forbidden`: **fixtures に存在しない**断片。S1が拾ってしまえば捏造。
  docs/08_ドシエ収集プロンプト集.md「実機確認の結果」節が実名で列挙した捏造の
  実例(架空の5区分・矛盾する営業利益・「浸水深30cm~50cm未満」等)は、そのまま
  forbidden の題材として使ってよい(実際に起きた捏造だから)。
- **gold は必ず `bench/fixtures/<company>/` を読んで作る。** 素材に書いていない
  ことを「たぶんこうだろう」で exact/recall に足さない。各社 10〜20 項目
  (exact + recall + forbidden の合計)を目安にする。
- 作ったら **`python3 tools/bench_s1.py --gold-check` を必ず通す**。forbidden が
  実は素材に書いてあった・recall が実は無かった、を機械で止める。

## fixtures(貼付素材)について

`bench/fixtures/<company>/input_*.txt` は 13章§2.2 の `data_key` のうち
`input_` で始まる10本(`input_hp` / `input_yuho` / `input_memo` /
`input_contract` / `input_prev_renewal` / `input_dossier` / `input_field_notes` /
`input_coverage_note` / `input_finance` / `input_hearing_answers`)に対応する。
**無い欄は空ファイルを置く**(0バイトでよい。採点器はそれを「その欄は検査しない」
として扱う)。

現在の5社(ニデック・成英工業・エイベックス・レバレジーズ・春華堂)は
`docs/demo/確認用v2/出力サンプル*.html` の「STEP 1｜入力と充足度診断」欄
(収集した情報の要約)と現場メモ・出典欄から**そのまま書き起こした**もので、
実際にDRへ貼った生の原文そのものではない(そのHTML自体が「手動模擬実行」の
デモ出力であり、貼付前の生素材は本リポジトリに残っていない)。したがって
現状のgoldは「デモHTMLが書いている事実の範囲」を正とした簡易版であり、
**本物の貼付原文が手に入ったら fixtures と gold を作り直すこと**(§「concerns」
参照)。

## live 運用(会社PCでS1を回した案件JSONをここへ集める)

実AIは社外(このリポジトリの開発環境)からは呼べない。実機での抽出精度は、
**会社PCで実際にS1を回して得たS1出力**をここへ運んで採点する。

1. 会社PCの「リスク提案ナビ」で対象案件を開く。
2. HOME(またはナビ画面)の「案件を書き出す」で**案件JSONをエクスポート**する
   (実体は `modNaviStore.ExportCaseJson`。全 `data_key` を含み、`s1_json` も
   その1本として入っている。フォーマットは
   `{"format":"riscon-navi-case","version":1,"case":{...},"data":[{"key":"s1_json","content":"..."}...]}`)。
3. そのファイルを `bench/out/live/<company>.json` として置く
   (`<company>` は `bench/gold/<company>.json` の `fixture_dir` と同じ名前に
   揃える。同一companyを複数回実行した場合は `<company>.1.json` /
   `<company>.2.json` のように連番を付けると「揺れ」も測れる)。
4. `python3 tools/bench_s1.py --mode replay --in bench/out/live` を実行する。
   採点器は案件JSON全体からでも生のS1 JSONそのものからでも読める
   (`load_s1_output` がどちらの形かを自動判定する)。

会社PCでの持ち出しは `docs/24_実機テスト手順書_Windows.md` §7 の観察項目に
1行を追加してある(「S1出力を bench/out/live へ集める」)。数値の解釈は本ベンチが
`gate.py` に入っていない=**参考指標**であることを忘れないこと(月次で見る)。

## 自己テスト・変異注入の記録(裁定書37の「敵対的検証・変異注入を必須とする」規約)

`python3 tools/bench_s1.py --mode selfcheck` は以下を確認する:

1. `src/test/modMockLlm.bas` の `BuildS1NewJson()` から実際にmock応答を
   実行時抽出し(ハードコードで手写ししない)、JSONとしてパースできること。
2. 無傷のmock応答を専用のミニgoldで採点すると 一致率1.0・回収率1.0・捏造率0.0
   になること。
3. `company_name` を1フィールード壊すと一致率が1.0未満に落ちること。
4. forbidden断片を本文へ混ぜると捏造率が0より大きくなること。
5. **変異注入(a)**: `fragment_found` を常時 `True` を返す関数に差し替えると、
   本来0.0のはずの捏造率が1.0(全件誤検知)になること=このテスト自体が
   「常にPASSする恒真テスト」になっていないことの裏取り。
6. **変異注入(b)**: `normalize_for_match` を無効化(素通し)したときの数値を
   出力する(このgoldの文字列は表記揺れを含まないため数値は変わらないが、
   `--gold-check` 側では正規化を無効化すると表記揺れのある断片が拾えなくなり
   一致率・回収率が下がることを確認できる。手順は下記)。

`--gold-check` への変異注入デモ(実演して復元済み。復元後は `git status` に
差分が残らないことを確認すること):

```
# gold に素材内の断片(例: shunkado.json の forbidden へ "うなぎパイ" を追加)
# -> forbiddenは「素材に無いこと」が条件なので gold-check が赤くなる
python3 tools/bench_s1.py --gold-check   # NG: forbidden断片が fixtures に実在します
# 追加した1行を消して復元
python3 tools/bench_s1.py --gold-check   # OK に戻る
```

## 既知の限界(concerns)

- 現5社のfixturesはデモHTMLの要約文からの書き起こしであり、実際にDRへ貼った
  生の貼付原文そのものではない。実際の貼付原文が手に入り次第、fixtures/gold を
  作り直すこと。
- `exact` は現状 `company_name` と `financials.source` の2項目に絞っている
  (5社とも決算数値が非開示・不明のため、数値系exactフィールドは「事実として
  確定できる」ものが乏しい)。有報の数字が取れる会社が増えたら
  `financials.net_assets` 等を追加する。
- `sources[]` の採点は班Aの15章改訂(未確定)を先取りした設計。キー名や形が
  変わったら `load_s1_output` 直後の `sources` 参照だけを直せばよい設計にして
  ある(スキーマ全体には依存しない)。
