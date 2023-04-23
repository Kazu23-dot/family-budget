Attribute VB_Name = "mdlCmnConst"

'''*************************************************************************************
'''<summary>
''' 定数一覧(仕入シート[01～12])　※テストコメント
'''</summary>
'''<option>
''' 2021.08.10 K.Hiraoka Created
'''</option>
'''*************************************************************************************
''シートの読取位置を定義
''***********************
'Public Const PURCHASING_DATA_SHEET_START_COL As String = "C"  '仕入シート(0～12)の読取開始列
'Public Const PURCHASING_DATA_SHEET_START_ROW As Integer = 14  '仕入シート(0～12)の読取開始行
'
''シートから取得したデータの格納用テーブル構造を定義
''***************************************************
'Public Const ARRAY_POSTING_TARGET_DATA_MAX_COL_NUM As Integer = 16 '取得データの最大カラム数を定義
'
'Public Const ARRAY_POSTING_TARGET_CHECK_FLAG_COL As String = "B"     '転記チェックフラグ
'Public Const ARRAY_POSTING_TARGET_DATE_COL As String = "C"           '日付
'Public Const ARRAY_POSTING_TARGET_PRODUCT_NM_COL As String = "S"     '商品名
'Public Const ARRAY_POSTING_TARGET_PURCHASING_STORE_COL As String = "E" '仕入店舗
'Public Const ARRAY_POSTING_TARGET_SUPPLIER_COL As String = "F" '仕入先
'Public Const ARRAY_POSTING_TARGET_POSTAGE_COL As String = "I"          '送料
'Public Const ARRAY_POSTING_TARGET_POINT_PAYMENT_AMOUNT_COL As String = "J" 'P払い額
'Public Const ARRAY_POSTING_TARGET_PURCHASE_AMOUNT_COL As String = "M"      '購入額
'Public Const ARRAY_POSTING_TARGET_REDUCTION_RATE_COL As String = "P"       '還元率
'Public Const ARRAY_POSTING_TARGET_SALES_CHANNEL_COL As String = "Q"        '販路
'Public Const ARRAY_POSTING_TARGET_ESTIMATED_SELLING_PRICE_COL As String = "T" '想定販売価格
'Public Const ARRAY_POSTING_TARGET_ASIN_COL As String = "R"                    'ASIN
'Public Const ARRAY_POSTING_TARGET_ESTIMATED_POSTAGE_COL As String = "U"       '想定送料
'Public Const ARRAY_POSTING_TARGET_OTHER_FEES_COL As String = "W"              'その他手数料
'Public Const ARRAY_POSTING_TARGET_QUANTITY_COL As String = "X"                '数量
'Public Const ARRAY_POSTING_TARGET_GROUPING_FLAG_COL As String = "AH"          'グループ化フラグ
'
''取得データ格納用配列のカラムインデックスを定義
''**********************************************
'Public Const ARRAY_POSTING_TARGET_CHECK_FLAG_COL_INDEX As Integer = 0     '転記チェックフラグ
'Public Const ARRAY_POSTING_TARGET_DATE_COL_INDEX As Integer = 1           '日付
'Public Const ARRAY_POSTING_TARGET_PRODUCT_NM_COL_INDEX As Integer = 2     '商品名
'Public Const ARRAY_POSTING_TARGET_PURCHASING_STORE_COL_INDEX As Integer = 3 '仕入店舗
'Public Const ARRAY_POSTING_TARGET_SUPPLIER_COL_INDEX As Integer = 4         '仕入先
'Public Const ARRAY_POSTING_TARGET_POSTAGE_COL_INDEX As Integer = 5         '送料
'Public Const ARRAY_POSTING_TARGET_POINT_PAYMENT_AMOUNT_COL_INDEX As Integer = 6 'P払い額
'Public Const ARRAY_POSTING_TARGET_PURCHASE_AMOUNT_COL_INDEX As Integer = 7      '購入額
'Public Const ARRAY_POSTING_TARGET_REDUCTION_RATE_COL_INDEX As Integer = 8       '還元率
'Public Const ARRAY_POSTING_TARGET_SALES_CHANNEL_COL_INDEX As Integer = 9        '販路
'Public Const ARRAY_POSTING_TARGET_ESTIMATED_SELLING_PRICE_COL_INDEX As Integer = 10 '想定販売価格
'Public Const ARRAY_POSTING_TARGET_ASIN_COL_INDEX As Integer = 11                    'ASIN
'Public Const ARRAY_POSTING_TARGET_ESTIMATED_POSTAGE_COL_INDEX As Integer = 12       '想定送料
'Public Const ARRAY_POSTING_TARGET_OTHER_FEES_COL_INDEX As Integer = 13              'その他手数料
'Public Const ARRAY_POSTING_TARGET_QUANTITY_COL_INDEX As Integer = 14                '数量
'Public Const ARRAY_POSTING_TARGET_GROUPING_FLAG_COL_INDEX As Integer = 15           'グループ化フラグ
'Public Const ARRAY_POSTING_TARGET_SERIAL_NUMBER_COL_INDEX As Integer = 16           'シリアルNo


'''*************************************************************************************
'''<summary>
''' 定数一覧(Amazon売上シート[TSVアップロード])
'''</summary>
'''<option>
''' 2021.08.14 K.Hiraoka Created
'''</option>
'''*************************************************************************************
''取得した売上反映用配列の各カラムを定義
''***************************************************
'Public Const POSTING_UPLOAD_AMAZON_EARNINGS_DATE_COL_INDEX As Integer = 0     '売上日時
'Public Const POSTING_UPLOAD_AMAZON_EARNINGS_ORDER_NO_COL_INDEX As Integer = 1     '注文番号
'Public Const POSTING_UPLOAD_AMAZON_EARNINGS_TRANSACTION_KATEGORY_COL_INDEX As Integer = 2     'トランザクションの種類
'Public Const POSTING_UPLOAD_AMAZON_EARNINGS_SKU_COL_INDEX As Integer = 3     'SKU
'Public Const POSTING_UPLOAD_AMAZON_EARNINGS_PRODUCT_NM_COL_INDEX As Integer = 4     '商品名
'Public Const POSTING_UPLOAD_AMAZON_EARNINGS_SALES_COL_INDEX As Integer = 5     '商品代金
'Public Const POSTING_UPLOAD_AMAZON_EARNINGS_PROMOTIONAL_DISCOUNT_COL_INDEX As Integer = 6     'プロモーション割引
'Public Const POSTING_UPLOAD_AMAZON_EARNINGS_AMAZON_COMMISSION_COL_INDEX As Integer = 7     'Amazon手数料
'Public Const POSTING_UPLOAD_AMAZON_EARNINGS_OTHER_COST_COL_INDEX As Integer = 8     'その他費用
'Public Const POSTING_UPLOAD_AMAZON_EARNINGS_QUANTITY_COL_INDEX As Integer = 9     '数量
'
''取得した売上TSVファイルの各カラムを定義
''***************************************
'Public Const TSV_UPLOAD_AMAZON_EARNINGS_DATE_COL_INDEX As Integer = 0     '売上日時
'Public Const TSV_UPLOAD_AMAZON_EARNINGS_ORDER_NO_COL_INDEX As Integer = 1     '注文番号
'Public Const TSV_UPLOAD_AMAZON_EARNINGS_SKU_COL_INDEX As Integer = 2     'SKU
'Public Const TSV_UPLOAD_AMAZON_EARNINGS_TRANSACTION_KATEGORY_COL_INDEX As Integer = 3     'トランザクションの種類
'Public Const TSV_UPLOAD_AMAZON_EARNINGS_PAYMENT_TYPE_COL_INDEX As Integer = 4     '支払い種別
'Public Const TSV_UPLOAD_AMAZON_EARNINGS_PRODUCT_NM_COL_INDEX As Integer = 5     '商品名
'Public Const TSV_UPLOAD_AMAZON_EARNINGS_SALES_COL_INDEX As Integer = 6     '売上金
