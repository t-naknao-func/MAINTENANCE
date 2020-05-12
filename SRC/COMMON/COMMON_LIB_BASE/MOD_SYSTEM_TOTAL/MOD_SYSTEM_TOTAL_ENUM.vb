Public Module MOD_SYSTEM_TOTAL_ENUM

    'テーブルMNG_M_KINDのCODE_FLAG
    Public Enum ENM_MNG_M_KIND_CODE_FLAG
        FLAG_GRANT = 20 'ユーザー権限種別
    End Enum

    Public Enum ENM_SYSTEM_TOTAL_FLAG_GRANT
        SYSTEM_ADMIN = 0 'システム管理者
        UNNYO_ADMIN = 1 '運用管理者
        TANTO = 2 '業務担当者
        IPPAN = 3 '一般職員
    End Enum

    Public Enum ENM_SYSTEM_TOTAL_FLAG_FDMAKE
        NO_MAKE = 0
        MAKE = 1
    End Enum
End Module
