# 業種対応表(RiskMap → 業種マスタ industry_code)

D-1: `RiskMap/` 配下72業種を、日本標準産業分類(中分類2桁)に基づき業種マスタへ対応づけた表。

- 既存30行に当たる業種は既存行へ `aliases` を追記(code・名称は変更していない)。
- 当たらない業種は新規行を追加。同一中分類に複数のRiskMap業種が入る場合は `2桁-連番` で細分化。
- 43番(太陽光発電・屋根設置型)は40番(太陽光発電)と同一industry_code(33)のエイリアスとして扱った(業種としては同一、設置形態の違いのため)。

| RiskMapファイル | RiskMap業種名 | industry_code | 業種マスタ名 | 扱い |
|---|---|---|---|---|
| 10denkikikaiseizougyou.docx | 電気機械製造業 | 29 | 電気機械器具製造業 | 既存行+alias |
| 11sangyouyoukikaiseizougyou.docx | 産業用機械製造業 | 26 | 生産用機械器具製造業 | 既存行+alias |
| 12seimitsukikaiseizougyou.docx | 精密機械製造業 | 27 | 業務用機械器具製造業 | 新規行 |
| 13kenchikugyou.docx | 建築業 | 06 | 総合工事業 | 既存行+alias |
| 14denki_setsubikoujigyou.docx | 電気・設備工事業 | 08 | 設備工事業 | 既存行+alias |
| 15unsougyou.docx | 運送業 | 44 | 道路貨物運送業 | 既存行+alias |
| 16iryouhinhanbaigyou.docx | 衣料品販売業 | 57 | 織物・衣服・身の回り品小売業 | 新規行 |
| 17iyakuhin_keshouhinhanbaigyou.docx | 医薬品・化粧品販売業 | 60 | その他の小売業 | 既存行+alias |
| 18kateidenki_kiguhanbaigyou.docx | 家庭電気・器具販売業 | 59 | 機械器具小売業 | 新規行 |
| 19supa_hyakkatengyou.docx | スーパー・百貨店業 | 56 | 各種商品小売業 | 新規行 |
| 1shokuhinseizougyou.docx | 食品製造業 | 09 | 食料品製造業 | 既存行+alias |
| 20oroshiurigyou.docx | 卸売業 | 50 | 各種商品卸売業 | 既存行+alias |
| 21gorufujou.docx | ゴルフ場 | 80 | 娯楽業 | 新規行 |
| 22fittonesukurabu.docx | フィットネスクラブ | 80-2 | 娯楽業 | 新規行 |
| 23hoteru_ryokangyou.docx | ホテル・旅館業 | 75 | 宿泊業 | 既存行+alias |
| 24yuuenchi.docx | 遊園地 | 80-3 | 娯楽業 | 新規行 |
| 25kankonsousaigyou.docx | 冠婚葬祭業 | 95 | その他のサービス業 | 新規行 |
| 26sentaku_riyou_biyougyou.docx | 洗濯・理容・美容業 | 78 | 洗濯・理容・美容・浴場業 | 新規行 |
| 27gakkouhoujin_gakushuujuku.docx | 学校法人・学習塾 | 81 | 学校教育 | 新規行 |
| 28inshokugyou.docx | 飲食業 | 76 | 飲食店 | 既存行+alias |
| 29kyuushoku_inshokutakuhaigyou.docx | 給食・飲食宅配業 | 77 | 持ち帰り・配達飲食サービス業 | 新規行 |
| 2seniseihinseizougyou.docx | 繊維製品製造業 | 11 | 繊維工業 | 既存行+alias |
| 30iryou_fukushigyou.docx | 医療・福祉業 | 83 | 医療業 | 新規行 |
| 31jouhoutsuushingyou.docx | 情報通信業 | 39 | 情報サービス業 | 既存行+alias |
| 32zeirishi_kaikeishi.docx | 税理士・会計士 | 72 | 専門サービス業 | 新規行 |
| 33shuppan_koukokugyou.docx | 出版・広告業 | 73 | 広告業 | 新規行 |
| 34jinzaihakengyou.docx | 人材派遣業 | 91 | 職業紹介・労働者派遣業 | 新規行 |
| 35gasorinsutando.docx | ガソリンスタンド | 60 | その他の小売業 | 既存行+alias |
| 36keibigyou.docx | 警備業 | 92 | その他の事業サービス業 | 新規行 |
| 37birumentenansugyou.docx | ビルメンテナンス業 | 92-2 | その他の事業サービス業 | 新規行 |
| 38housoukyoku.docx | 放送局 | 38 | 放送業 | 新規行 |
| 39jidoushaseibigyou.docx | 自動車整備業 | 89 | 自動車整備業 | 新規行 |
| 3mokuseihin_kaguseizougyou.docx | 木製品・家具製造業 | 12 | 木材・木製品・家具製造業 | 新規行 |
| 40taiyoukouhatsuden.docx | 太陽光発電 | 33 | 電気業 | 新規行 |
| 41shousuiryokuhatsuden.docx | 小水力発電 | 33-2 | 電気業 | 新規行 |
| 42baiomasuhatsudengyou.docx | バイオマス発電業 | 33-3 | 電気業 | 新規行 |
| 43taiyoukouhatsudenyanesetchigata_.docx | 太陽光発電（屋根設置型） | 33 | 電気業 | エイリアス(40と同一) |
| 44chinetsuhatsuden.docx | 地熱発電 | 33-4 | 電気業 | 新規行 |
| 45menruiseizougyou.docx | 麺類製造業 | 09 | 食料品製造業 | 既存行+alias |
| 46seishigyou.docx | 製紙業 | 14 | パルプ・紙・紙加工品製造業 | 新規行 |
| 47dobokukoujigyou.docx | 土木工事業 | 06 | 総合工事業 | 既存行+alias |
| 48baikubingyou.docx | バイク便業 | 44 | 道路貨物運送業 | 既存行+alias |
| 49doraggusutoa.docx | ドラッグストア | 60 | その他の小売業 | 既存行+alias |
| 4insatsugyou.docx | 印刷業 | 15 | 印刷・同関連業 | 新規行 |
| 50chouzaiyakkyoku.docx | 調剤薬局 | 60 | その他の小売業 | 既存行+alias |
| 51tsuushinhanbaigyou.docx | 通信販売業 | 61 | 無店舗小売業 | 既存行+alias |
| 52pachinkogyou.docx | パチンコ業 | 80-4 | 娯楽業 | 新規行 |
| 53bijinesuhoteru.docx | ビジネスホテル | 75 | 宿泊業 | 既存行+alias |
| 54hoikuen_youchien.docx | 保育園・幼稚園 | 85-2 | 社会保険・社会福祉・介護事業 | 新規行 |
| 55jidoushakyoushuujo.docx | 自動車教習所 | 82 | その他の教育,学習支援業 | 新規行 |
| 56kaigosabisujigyou.docx | 介護サービス事業 | 85 | 社会保険・社会福祉・介護事業 | 新規行 |
| 57fudousangyou.docx | 不動産業 | 68 | 不動産取引業 | 既存行+alias |
| 58manshonkanrikumiai.docx | マンション管理組合 | 69 | 不動産賃貸業・管理業 | 新規行 |
| 59rokujisangyoujigyoushanougyou_.docx | 6次産業事業者（農業） | 01 | 農業 | 新規行 |
| 5iyakuhin_kagakuhinseizougyou.docx | 医薬品・化学品製造業 | 16 | 化学工業 | 既存行+alias |
| 60kuukou.docx | 空港 | 48 | 運輸に付帯するサービス業 | 新規行 |
| 61shuukyouhoujin.docx | 宗教法人 | 94 | 宗教 | 新規行 |
| 62jidoushaseizougyou.docx | 自動車製造業 | 31 | 輸送用機械器具製造業 | 既存行+alias |
| 63jidoushabuhinseizougyou.docx | 自動車部品製造業 | 31 | 輸送用機械器具製造業 | 既存行+alias |
| 64fuuryokuhatsuden.docx | 風力発電 | 33-5 | 電気業 | 新規行 |
| 65homusenta.docx | ホームセンター | 60 | その他の小売業 | 既存行+alias |
| 66shougyoushisetsukanrigyou.docx | 商業施設管理業 | 92-3 | その他の事業サービス業 | 新規行 |
| 67soukogyou.docx | 倉庫業 | 47 | 倉庫業 | 既存行+alias |
| 68ryokoudairitengyou.docx | 旅行代理店業 | 79 | その他の生活関連サービス業 | 新規行 |
| 69takushihaiyagyou.docx | タクシー・ハイヤー業 | 43 | 道路旅客運送業 | 新規行 |
| 6purasuchikkuseihinseizougyou.docx | プラスチック製品製造業 | 18 | プラスチック製品製造業 | 既存行+alias |
| 70ibentokaisai.docx | イベント開催 | 79-2 | その他の生活関連サービス業 | 新規行 |
| 71jidoushahanbaigyou.docx | 自動車販売業 | 59-2 | 機械器具小売業 | 新規行 |
| 72shoukibotetsudoujigyousha.docx | （小規模）鉄道事業者 | 42 | 鉄道業 | 新規行 |
| 7garasuseihinseizougyou.docx | ガラス製品製造業 | 21 | 窯業・土石製品製造業 | 既存行+alias |
| 8sonotakagakuhinseizougyou.docx | その他化学品製造業 | 16 | 化学工業 | 既存行+alias |
| 9kinzokuseihinseizougyou.docx | 金属製品製造業 | 24 | 金属製品製造業 | 既存行+alias |
