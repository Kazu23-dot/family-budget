Attribute VB_Name = "mdlCmnConst"

'''*************************************************************************************
'''<summary>
''' �萔�ꗗ(�d���V�[�g[01�`12])�@���e�X�g�R�����g
'''</summary>
'''<option>
''' 2021.08.10 K.Hiraoka Created
'''</option>
'''*************************************************************************************
''�V�[�g�̓ǎ�ʒu���`
''***********************
'Public Const PURCHASING_DATA_SHEET_START_COL As String = "C"  '�d���V�[�g(0�`12)�̓ǎ�J�n��
'Public Const PURCHASING_DATA_SHEET_START_ROW As Integer = 14  '�d���V�[�g(0�`12)�̓ǎ�J�n�s
'
''�V�[�g����擾�����f�[�^�̊i�[�p�e�[�u���\�����`
''***************************************************
'Public Const ARRAY_POSTING_TARGET_DATA_MAX_COL_NUM As Integer = 16 '�擾�f�[�^�̍ő�J���������`
'
'Public Const ARRAY_POSTING_TARGET_CHECK_FLAG_COL As String = "B"     '�]�L�`�F�b�N�t���O
'Public Const ARRAY_POSTING_TARGET_DATE_COL As String = "C"           '���t
'Public Const ARRAY_POSTING_TARGET_PRODUCT_NM_COL As String = "S"     '���i��
'Public Const ARRAY_POSTING_TARGET_PURCHASING_STORE_COL As String = "E" '�d���X��
'Public Const ARRAY_POSTING_TARGET_SUPPLIER_COL As String = "F" '�d����
'Public Const ARRAY_POSTING_TARGET_POSTAGE_COL As String = "I"          '����
'Public Const ARRAY_POSTING_TARGET_POINT_PAYMENT_AMOUNT_COL As String = "J" 'P�����z
'Public Const ARRAY_POSTING_TARGET_PURCHASE_AMOUNT_COL As String = "M"      '�w���z
'Public Const ARRAY_POSTING_TARGET_REDUCTION_RATE_COL As String = "P"       '�Ҍ���
'Public Const ARRAY_POSTING_TARGET_SALES_CHANNEL_COL As String = "Q"        '�̘H
'Public Const ARRAY_POSTING_TARGET_ESTIMATED_SELLING_PRICE_COL As String = "T" '�z��̔����i
'Public Const ARRAY_POSTING_TARGET_ASIN_COL As String = "R"                    'ASIN
'Public Const ARRAY_POSTING_TARGET_ESTIMATED_POSTAGE_COL As String = "U"       '�z�著��
'Public Const ARRAY_POSTING_TARGET_OTHER_FEES_COL As String = "W"              '���̑��萔��
'Public Const ARRAY_POSTING_TARGET_QUANTITY_COL As String = "X"                '����
'Public Const ARRAY_POSTING_TARGET_GROUPING_FLAG_COL As String = "AH"          '�O���[�v���t���O
'
''�擾�f�[�^�i�[�p�z��̃J�����C���f�b�N�X���`
''**********************************************
'Public Const ARRAY_POSTING_TARGET_CHECK_FLAG_COL_INDEX As Integer = 0     '�]�L�`�F�b�N�t���O
'Public Const ARRAY_POSTING_TARGET_DATE_COL_INDEX As Integer = 1           '���t
'Public Const ARRAY_POSTING_TARGET_PRODUCT_NM_COL_INDEX As Integer = 2     '���i��
'Public Const ARRAY_POSTING_TARGET_PURCHASING_STORE_COL_INDEX As Integer = 3 '�d���X��
'Public Const ARRAY_POSTING_TARGET_SUPPLIER_COL_INDEX As Integer = 4         '�d����
'Public Const ARRAY_POSTING_TARGET_POSTAGE_COL_INDEX As Integer = 5         '����
'Public Const ARRAY_POSTING_TARGET_POINT_PAYMENT_AMOUNT_COL_INDEX As Integer = 6 'P�����z
'Public Const ARRAY_POSTING_TARGET_PURCHASE_AMOUNT_COL_INDEX As Integer = 7      '�w���z
'Public Const ARRAY_POSTING_TARGET_REDUCTION_RATE_COL_INDEX As Integer = 8       '�Ҍ���
'Public Const ARRAY_POSTING_TARGET_SALES_CHANNEL_COL_INDEX As Integer = 9        '�̘H
'Public Const ARRAY_POSTING_TARGET_ESTIMATED_SELLING_PRICE_COL_INDEX As Integer = 10 '�z��̔����i
'Public Const ARRAY_POSTING_TARGET_ASIN_COL_INDEX As Integer = 11                    'ASIN
'Public Const ARRAY_POSTING_TARGET_ESTIMATED_POSTAGE_COL_INDEX As Integer = 12       '�z�著��
'Public Const ARRAY_POSTING_TARGET_OTHER_FEES_COL_INDEX As Integer = 13              '���̑��萔��
'Public Const ARRAY_POSTING_TARGET_QUANTITY_COL_INDEX As Integer = 14                '����
'Public Const ARRAY_POSTING_TARGET_GROUPING_FLAG_COL_INDEX As Integer = 15           '�O���[�v���t���O
'Public Const ARRAY_POSTING_TARGET_SERIAL_NUMBER_COL_INDEX As Integer = 16           '�V���A��No


'''*************************************************************************************
'''<summary>
''' �萔�ꗗ(Amazon����V�[�g[TSV�A�b�v���[�h])
'''</summary>
'''<option>
''' 2021.08.14 K.Hiraoka Created
'''</option>
'''*************************************************************************************
''�擾�������㔽�f�p�z��̊e�J�������`
''***************************************************
'Public Const POSTING_UPLOAD_AMAZON_EARNINGS_DATE_COL_INDEX As Integer = 0     '�������
'Public Const POSTING_UPLOAD_AMAZON_EARNINGS_ORDER_NO_COL_INDEX As Integer = 1     '�����ԍ�
'Public Const POSTING_UPLOAD_AMAZON_EARNINGS_TRANSACTION_KATEGORY_COL_INDEX As Integer = 2     '�g�����U�N�V�����̎��
'Public Const POSTING_UPLOAD_AMAZON_EARNINGS_SKU_COL_INDEX As Integer = 3     'SKU
'Public Const POSTING_UPLOAD_AMAZON_EARNINGS_PRODUCT_NM_COL_INDEX As Integer = 4     '���i��
'Public Const POSTING_UPLOAD_AMAZON_EARNINGS_SALES_COL_INDEX As Integer = 5     '���i���
'Public Const POSTING_UPLOAD_AMAZON_EARNINGS_PROMOTIONAL_DISCOUNT_COL_INDEX As Integer = 6     '�v�����[�V��������
'Public Const POSTING_UPLOAD_AMAZON_EARNINGS_AMAZON_COMMISSION_COL_INDEX As Integer = 7     'Amazon�萔��
'Public Const POSTING_UPLOAD_AMAZON_EARNINGS_OTHER_COST_COL_INDEX As Integer = 8     '���̑���p
'Public Const POSTING_UPLOAD_AMAZON_EARNINGS_QUANTITY_COL_INDEX As Integer = 9     '����
'
''�擾��������TSV�t�@�C���̊e�J�������`
''***************************************
'Public Const TSV_UPLOAD_AMAZON_EARNINGS_DATE_COL_INDEX As Integer = 0     '�������
'Public Const TSV_UPLOAD_AMAZON_EARNINGS_ORDER_NO_COL_INDEX As Integer = 1     '�����ԍ�
'Public Const TSV_UPLOAD_AMAZON_EARNINGS_SKU_COL_INDEX As Integer = 2     'SKU
'Public Const TSV_UPLOAD_AMAZON_EARNINGS_TRANSACTION_KATEGORY_COL_INDEX As Integer = 3     '�g�����U�N�V�����̎��
'Public Const TSV_UPLOAD_AMAZON_EARNINGS_PAYMENT_TYPE_COL_INDEX As Integer = 4     '�x�������
'Public Const TSV_UPLOAD_AMAZON_EARNINGS_PRODUCT_NM_COL_INDEX As Integer = 5     '���i��
'Public Const TSV_UPLOAD_AMAZON_EARNINGS_SALES_COL_INDEX As Integer = 6     '�����
