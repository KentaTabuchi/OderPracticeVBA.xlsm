Attribute VB_Name = "ConstNum"
Option Explicit
'data�V�[�g�̗�萔

Enum DATA_COLUMN
    DATE_2 = 2
    delivery = 3
    sales = 4
    loss = 5
    priceCut = 6
    STOCK = 7
    Oder = 8
    CUSTOMER_NUM = 9
End Enum

Enum DATA_ROW
    TABLE_TOP = 2
    TABLE_END = 34
End Enum

'main�V�[�g�̍s�萔
Enum MAIN_ROW
    date_ = 2
    CARRY_OVER_STOCK = 3
    PRICE_CUT = 4
    delivery = 5
    sales = 6
    loss = 7
    CURRENT_STOCK = 8
    ITEMS = 15
End Enum
'main�V�[�g�̗�萔
Enum MAIN_COLUMN
    TABLE_LEFT_EDGE = 6
    number = 6
    Price = 7
End Enum

Enum MAIN_PHASE_PAIN
    ROW = 24
    COLUMN = 2
End Enum

Enum PhaseNumber
    START_0 = 0
    DELIVERY_1 = 1
    SELL_2 = 2
    LOSS_3 = 3
    LAST_4 = 4
End Enum




