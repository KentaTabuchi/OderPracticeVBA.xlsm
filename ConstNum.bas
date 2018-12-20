Attribute VB_Name = "ConstNum"
Option Explicit
'dataシートの列定数

Enum DATA_COLUMN
    DATE_2 = 2
    DELIVERY = 3
    sales = 4
    loss = 5
    priceCut = 6
    STOCK = 7
    Oder = 8
End Enum

'mainシートの行定数
Enum MAIN_ROW
    DATE_ = 2
    CARRY_OVER_STOCK = 3
    PRICE_CUT = 4
    DELIVERY = 5
    sales = 6
    loss = 7
    CURRENT_STOCK = 8
    ITEMS = 15
End Enum
'mainシートの列定数
Enum MAIN_COLUMN
    Number = 7
    Price = 8
End Enum

Enum PhaseNumber
    START_0 = 0
    DELIVERY_1 = 1
    SELL_2 = 2
    LOSS_3 = 3
    LAST_4 = 4
End Enum




