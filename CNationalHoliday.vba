Option Explicit

'//////////////////////////////////////////////////
'                   概　　要
'//////////////////////////////////////////////////
'
'用　　　途：祝日取得、確認用
'処理対象日：1948/7/20以降（2050年までは、春分の日、秋分の日確認済み）
'備　　　考：2020/2/23天皇誕生日対応済み
'作　　　成：2018/5/13

'//////////////////////////////////////////////////
'                   参照設定
'//////////////////////////////////////////////////

'Dictionary用
'Microsoft Scripting Runtime


'//////////////////////////////////////////////////
'                ユーザー定義型
'//////////////////////////////////////////////////

'月日固定の祝日情報
Private Type FixMD
    sMD         As String
    lBeginYear  As Long
    lEndYear    As Long
    sName       As String
End Type

'月週曜日固定の祝日情報
Private Type FixWN
    lMonth      As Long
    lNthWeek    As Long
    lDayOfWeek  As Long
    lBeginYear  As Long
    lEndYear    As Long
    sName       As String
End Type

'//////////////////////////////////////////////////
'                   定数
'//////////////////////////////////////////////////

'「国民の祝日に関する法律」施行年月日
Private Const BEGIN_DATE    As Date = #7/20/1948#

'「振替休日」施行年月日
Private Const TRANSFER_HOLIDAY1_BEGIN_DATE    As Date = #4/12/1973#
Private Const TRANSFER_HOLIDAY2_BEGIN_DATE    As Date = #1/1/2007#

'「国民の休日」施行年月日
Private Const NATIONAL_HOLIDAY_BEGIN_DATE       As Date = #12/27/1985#

'年上限
Private Const YEAR_MAX      As Long = 2050

'エラーコード（パラメータ異常）
Private Const ERROR_INVALID_PARAMETER   As Long = &H57


'//////////////////////////////////////////////////
'               Private変数
'//////////////////////////////////////////////////

'国民の祝日格納用ディクショナリ
'キー：年月日（DateTime型）
'値　：祝日名
Private dicHoliday_ As Object

Private lInitializedLastYear_   As Long


'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'                       祝日情報の定義
'
'　基本的な祝日情報は、以下の２つのメソッド内で定義する。
'　　getNationalHolidayInfoMD
'　　getNationalHolidayInfoWN
'
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'//////////////////////////////////////////////////
'月日固定の祝日情報生成
'//////////////////////////////////////////////////
Private Sub getNationalHolidayInfoMD(ByRef uFixMD() As FixMD)

    Dim sFixMD(26)  As String   '祝日データを追加削除した場合、この配列要素数を変更すること
    Dim sResult()   As String
    Dim i           As Long

    '//////////////////////////////////////////////////
    '               月日固定の祝日
    '//////////////////////////////////////////////////
    '適用開始年について
    '　元旦（1/1）
    '　成人の日（1/15）
    '　天皇誕生日（4/29）
    '　憲法記念日（5/3）
    '　こどもの日（5/5）
    'の５つは、「国民の祝日に関する法律」施行年（1948年）に制定されているが
    '同法の施行が7/20であり、それ以前となるため、適用開始年を翌年（1949年）に補正してある。
    '
    '月日,適用開始年,適用終了年,名前
    '適用終了年；9999は、現在も適用中
    sFixMD(0) = "01/01,1949,9999,元日"          '適用開始年補正済み
    sFixMD(1) = "01/15,1949,1999,成人の日"      '適用開始年補正済み
    sFixMD(2) = "02/11,1967,9999,建国記念の日"
    sFixMD(3) = "02/23,2020,9999,天皇誕生日"    '適用開始年補正済み
    sFixMD(4) = "02/24,1989,1989,昭和天皇の大喪の礼"
    sFixMD(5) = "04/10,1959,1959,皇太子明仁親王の結婚の儀"
    sFixMD(6) = "04/29,1949,1988,天皇誕生日"    '適用開始年補正済み
    sFixMD(7) = "04/29,1989,2006,みどりの日"
    sFixMD(8) = "04/29,2007,9999,昭和の日"
    sFixMD(9) = "05/01,2019,2019,天皇の即位"
    sFixMD(10) = "05/03,1949,9999,憲法記念日"    '適用開始年補正済み
    sFixMD(11) = "05/04,2007,9999,みどりの日"
    sFixMD(12) = "05/05,1949,9999,こどもの日"    '適用開始年補正済み
    sFixMD(13) = "06/09,1993,1993,皇太子徳仁親王の結婚の儀"
    sFixMD(14) = "07/20,1996,2002,海の日"
    sFixMD(15) = "07/23,2020,2020,海の日"
    sFixMD(16) = "07/24,2020,2020,スポーツの日"
    sFixMD(17) = "08/10,2020,2020,山の日"
    sFixMD(18) = "08/11,2016,2019,山の日"
    sFixMD(19) = "08/11,2021,9999,山の日"
    sFixMD(20) = "09/15,1966,2002,敬老の日"
    sFixMD(21) = "10/10,1966,1999,体育の日"
    sFixMD(22) = "10/22,2019,2019,即位の礼正殿の儀"
    sFixMD(23) = "11/03,1948,9999,文化の日"
    sFixMD(24) = "11/12,1990,1990,即位礼正殿の儀"
    sFixMD(25) = "11/23,1948,9999,勤労感謝の日"
    sFixMD(26) = "12/23,1989,2018,天皇誕生日"

    ReDim uFixMD(UBound(sFixMD))

    For i = 0 To UBound(sFixMD)
        sResult = Split(sFixMD(i), ",")

        uFixMD(i).sMD = sResult(0)
        uFixMD(i).lBeginYear = CLng(sResult(1))
        uFixMD(i).lEndYear = CLng(sResult(2))
        uFixMD(i).sName = sResult(3)
    Next i

End Sub

'//////////////////////////////////////////////////
'月週曜日固定の祝日情報生成
'//////////////////////////////////////////////////
Private Sub getNationalHolidayInfoWN(ByRef uFixWN() As FixWN)

    Dim sFixWN(3)   As String   '祝日データを追加削除した場合、この配列要素数を変更すること
    Dim sResult()   As String
    Dim i           As Long

    '//////////////////////////////////////////////////
    '               月週曜日固定の祝日
    '//////////////////////////////////////////////////
    '月,週,曜日,適用開始年,適用終了年,名前
    '曜日：日 1
    '　　　月 2
    '　　　火 3
    '　　　水 4
    '　　　木 5
    '　　　金 6
    '　　　土 7
    '適用終了年；9999は、現在も適用中
    sFixWN(0) = "01,2,2,2000,9999,成人の日"
    sFixWN(1) = "07,3,2,2003,9999,海の日"
    sFixWN(2) = "09,3,2,2003,9999,敬老の日"
    sFixWN(3) = "10,2,2,2000,9999,体育の日"

    ReDim uFixWN(UBound(sFixWN))

    For i = 0 To UBound(sFixWN)
        sResult = Split(sFixWN(i), ",")

        uFixWN(i).lMonth = CLng(sResult(0))
        uFixWN(i).lNthWeek = CLng(sResult(1))
        uFixWN(i).lDayOfWeek = CLng(sResult(2))
        uFixWN(i).lBeginYear = CLng(sResult(3))
        uFixWN(i).lEndYear = CLng(sResult(4))
        uFixWN(i).sName = sResult(5)
    Next i
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'                       祝日情報の定義　ここまで
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub Class_Initialize()

    Set dicHoliday_ = CreateObject("Scripting.Dictionary")

    lInitializedLastYear_ = &H80000000

    'デフォルトで、現在の５年後までデータを生成する
    InitializedLastYear = Year(Now) + 5

End Sub

Private Sub Class_Terminate()

    Set dicHoliday_ = Nothing

End Sub


'//////////////////////////////////////////////////
'指定日が国民の祝日（休日）か？
'//////////////////////////////////////////////////
Public Function IsNationalHoliday(ByVal dtDate As Date) As Boolean

    Dim dtDateW As Date

    '時分秒データを切り捨てる
    dtDateW = DateSerial(Year(dtDate), Month(dtDate), Day(dtDate))

    If dtDateW < BEGIN_DATE Then
        Err.Raise ERROR_INVALID_PARAMETER, "isNationalHoliday", format$(dtDateW, "yyyy/mm/dd") & "は、適用範囲外です。"

        Exit Function
    ElseIf Year(dtDateW) > YEAR_MAX Then
        Err.Raise ERROR_INVALID_PARAMETER, "isNationalHoliday", format$(YEAR_MAX + 1, "yyyy年") & "以降は、適用範囲外です。"

        Exit Function
    ElseIf Year(dtDateW) > InitializedLastYear Then
        Err.Raise ERROR_INVALID_PARAMETER, "isNationalHoliday", format$(dtDateW, "yyyy年") & "は、データが生成されていないため、判定できません。" _
                            & vbCrLf & "reInitializeメソッドで対象年を設定後、再度確認してみて下さい。"

        Exit Function
    End If

    IsNationalHoliday = dicHoliday_.Exists(dtDateW)

End Function

'//////////////////////////////////////////////////
'指定日が国民の祝日（休日）か？そうであれば、その祝日名を合わせて返す
'//////////////////////////////////////////////////
Public Function isNationalHoliday2(ByVal dtDate As Date, ByRef sHolidayName As String) As Boolean

    Dim dtDateW As Date

    '時分秒データを切り捨てる
    dtDateW = DateSerial(Year(dtDate), Month(dtDate), Day(dtDate))

    isNationalHoliday2 = IsNationalHoliday(dtDateW)

    sHolidayName = GetNationalHolidayName(dtDateW)

End Function

'//////////////////////////////////////////////////
'指定年の祝日を配列に格納して返す
'//////////////////////////////////////////////////
Public Function getNationalHolidays(ByVal lYear As Long, ByRef dtHolidays() As Date) As Long

    Dim dtHolidaysW()   As Date
    Dim lHolidays       As Long
    Dim i As Long

    lHolidays = 0
    ReDim dtHolidaysW(lHolidays)

    For i = 0 To dicHoliday_.Count - 1
        If Year(dicHoliday_.Keys(i)) = lYear Then
            ReDim Preserve dtHolidaysW(lHolidays)

            dtHolidaysW(lHolidays) = dicHoliday_.Keys(i)

            lHolidays = lHolidays + 1
        End If
    Next i

    '昇順並べ替え
    Call qSort(dtHolidaysW, 0, UBound(dtHolidaysW))

    Erase dtHolidays
    dtHolidays = dtHolidaysW

    getNationalHolidays = lHolidays

End Function

'//////////////////////////////////////////////////
'指定日の祝日名を返す
'//////////////////////////////////////////////////
Public Function GetNationalHolidayName(ByVal dtHoliday As Date) As String

    Dim dtDateW As Date

    '時分秒データを切り捨てる
    dtDateW = DateSerial(Year(dtHoliday), Month(dtHoliday), Day(dtHoliday))

    If IsNationalHoliday(dtDateW) = True Then
        GetNationalHolidayName = dicHoliday_.Item(dtDateW)
    End If

End Function

'//////////////////////////////////////////////////
'何年までの祝日データが生成されているか
'//////////////////////////////////////////////////
Public Property Get InitializedLastYear() As Long

    InitializedLastYear = lInitializedLastYear_

End Property

'//////////////////////////////////////////////////
'指定年までの祝日データを生成させる（YEAR_MAX以下）
'　外部からの要求は、reInitializeで行うことが出来る
'//////////////////////////////////////////////////
Private Property Let InitializedLastYear(ByVal lInitializedLastYear As Long)

    If lInitializedLastYear < lInitializedLastYear_ Then
        '要求された最終年が初期化済みの年より前ならば、処理しない
        Exit Property
    ElseIf lInitializedLastYear > YEAR_MAX Then
        lInitializedLastYear = YEAR_MAX
    End If

    Call initDictionary(lInitializedLastYear)

    lInitializedLastYear_ = lInitializedLastYear

End Property

'//////////////////////////////////////////////////
'指定年までの祝日データを生成させる
'//////////////////////////////////////////////////
Public Sub reInitialize(ByVal lLastYear As Long)

    InitializedLastYear = lLastYear

End Sub

'//////////////////////////////////////////////////
'Dictionaryへ祝日情報を格納
'//////////////////////////////////////////////////
Private Sub initDictionary(ByVal lLastYear As Long)

    Dim uFixMD()    As FixMD
    Dim uFixWN()    As FixWN

    '月日固定の祝日情報
    Call getNationalHolidayInfoMD(uFixMD)

    '月週曜日固定の祝日情報
    Call getNationalHolidayInfoWN(uFixWN)

    'Dictionaryへ追加
    Call add2Dictionary(lLastYear, uFixMD, uFixWN)

End Sub

'//////////////////////////////////////////////////
'祝日情報をDictionaryへ格納
'//////////////////////////////////////////////////
Private Sub add2Dictionary(ByVal lLastYear As Long, ByRef uFixMD() As FixMD, ByRef uFixWN() As FixWN)

    Dim lInitializedLastYear    As Long
    Dim lBeginYear          As Long
    Dim lEndYear            As Long
    Dim dtHoliday           As Date
    Dim lAddedDays          As Long
    Dim dtAdded()           As Date
    Dim existsHoliday       As Boolean
    Dim lYear               As Long
    Dim i                   As Long

    '初期化済みの最終年を取得
    lInitializedLastYear = InitializedLastYear

    If lInitializedLastYear < Year(BEGIN_DATE) Then
        '施工年より前ならば、施工年を開始年とする
        lBeginYear = Year(BEGIN_DATE)
    Else
        '施工年以後なら、初期化済みの翌年を開始年とする
        lBeginYear = lInitializedLastYear + 1
    End If

    lEndYear = lLastYear

    For lYear = lBeginYear To lEndYear
        '年間の祝日格納用配列クリア
        lAddedDays = 0
        ReDim dtAdded(lAddedDays)

        '月日固定の祝日
        For i = 0 To UBound(uFixMD)
            '適用期間のみを対象とする
            If uFixMD(i).lBeginYear <= lYear And uFixMD(i).lEndYear >= lYear Then
                dtHoliday = CDate(CStr(lYear) & "/" & uFixMD(i).sMD)

                dicHoliday_.Add dtHoliday, uFixMD(i).sName

                ReDim Preserve dtAdded(lAddedDays)
                dtAdded(lAddedDays) = dtHoliday
                lAddedDays = lAddedDays + 1
            End If
        Next i

        '月週曜日固定の祝日
        For i = 0 To UBound(uFixWN)
            '適用期間のみを対象とする
            If uFixWN(i).lBeginYear <= lYear And uFixWN(i).lEndYear >= lYear Then
                dtHoliday = getNthWeeksDayOfWeek(lYear, uFixWN(i).lMonth, uFixWN(i).lNthWeek, uFixWN(i).lDayOfWeek)

                dicHoliday_.Add dtHoliday, uFixWN(i).sName

                ReDim Preserve dtAdded(lAddedDays)
                dtAdded(lAddedDays) = dtHoliday
                lAddedDays = lAddedDays + 1
            End If
        Next i

        '春分の日
        dtHoliday = getVernalEquinoxDay(lYear)
        dicHoliday_.Add dtHoliday, "春分の日"

        ReDim Preserve dtAdded(lAddedDays)
        dtAdded(lAddedDays) = dtHoliday
        lAddedDays = lAddedDays + 1

        '秋分の日
        dtHoliday = getAutumnalEquinoxDay(lYear)
        dicHoliday_.Add dtHoliday, "秋分の日"

        ReDim Preserve dtAdded(lAddedDays)
        dtAdded(lAddedDays) = dtHoliday
        lAddedDays = lAddedDays + 1

        '振替休日
        For i = 0 To lAddedDays - 1
            existsHoliday = existsSubstituteHoliday(dtAdded(i), dtHoliday)

            If existsHoliday = True Then
                dicHoliday_.Add dtHoliday, "振替休日"
            End If
        Next i

        '国民の休日
        For i = 0 To lAddedDays - 1
            existsHoliday = existsNationalHoliday(dtAdded(i), dtHoliday)

            If existsHoliday = True Then
                dicHoliday_.Add dtHoliday, "国民の休日"
            End If
        Next i

        Erase dtAdded
    Next lYear

End Sub

'//////////////////////////////////////////////////
'振替休日の有無
'　祝日（dtDate）に対する振替休日の有無（ある場合は、dtSubstituteHolidayに代入される）
'//////////////////////////////////////////////////
Private Function existsSubstituteHoliday(ByVal dtDate As Date, ByRef dtSubstituteHoliday As Date) As Boolean

    Dim dtNextDay   As Date

    existsSubstituteHoliday = False

    If dicHoliday_.Exists(dtDate) = False Then
        'dtDateが祝日でなければ終了
        Exit Function
    End If

    '適用期間のみを対象とする
    If dtDate >= TRANSFER_HOLIDAY1_BEGIN_DATE And dtDate < TRANSFER_HOLIDAY2_BEGIN_DATE Then
        If Weekday(dtDate) = vbSunday Then
            '祝日が日曜日であれば、翌日（月曜日）が振替休日
            dtSubstituteHoliday = DateAdd("d", 1, dtDate)

            existsSubstituteHoliday = True
        End If
    ElseIf dtDate >= TRANSFER_HOLIDAY2_BEGIN_DATE Then
        '「国民の祝日」が日曜日に当たるときは、その日後においてその日に最も近い「国民の祝日」でない日を休日とする
        If Weekday(dtDate) = vbSunday Then
            dtNextDay = DateAdd("d", 1, dtDate)

            '直近の祝日でない日を取得
            Do Until dicHoliday_.Exists(dtNextDay) = False
                dtNextDay = DateAdd("d", 1, dtNextDay)
            Loop

            dtSubstituteHoliday = dtNextDay

            existsSubstituteHoliday = True
        End If
    End If

End Function

'//////////////////////////////////////////////////
'国民の休日の有無
'　祝日（dtDate）に対す国民の休日の有無（ある場合は、dtNationalHolidayに代入される）
'//////////////////////////////////////////////////
Private Function existsNationalHoliday(ByVal dtDate As Date, ByRef dtNationalHoliday As Date) As Boolean

    Dim dtBaseDay   As Date
    Dim dtNextDay   As Date

    existsNationalHoliday = False

    If dicHoliday_.Exists(dtDate) = False Then
        'dtDateが祝日でなければ終了
        Exit Function
    End If

    '適用期間のみを対象とする
    If dtDate >= NATIONAL_HOLIDAY_BEGIN_DATE Then
        dtBaseDay = DateAdd("d", 1, dtDate)

        '直近の祝日でない日を取得
        Do Until dicHoliday_.Exists(dtBaseDay) = False
            dtBaseDay = DateAdd("d", 1, dtBaseDay)
        Loop

        '日曜日であれば対象外
        If Weekday(dtBaseDay) <> vbSunday Then
            dtNextDay = DateAdd("d", 1, dtBaseDay)

            '翌日が祝日であれば対象
            If dicHoliday_.Exists(dtNextDay) = True Then
                existsNationalHoliday = True

                dtNationalHoliday = dtBaseDay
            End If
        End If
    End If

End Function

'//////////////////////////////////////////////////
'月の第N W曜日の日時を取得
'//////////////////////////////////////////////////
Private Function getNthWeeksDayOfWeek(ByVal lYear As Long, _
                                      ByVal lMonth As Long, _
                                      ByVal lNth As Long, _
                                      ByVal lDayOfWeek As VbDayOfWeek) As Date

    Dim dt1stDate       As Date
    Dim lDayOfWeek1st   As Long
    Dim lOffset         As Long

    '指定年月の１日を取得
    dt1stDate = DateSerial(lYear, lMonth, 1)

    '１日の曜日を取得
    lDayOfWeek1st = Weekday(dt1stDate)

    '指定日へのオフセットを取得
    lOffset = lDayOfWeek - lDayOfWeek1st

    If lDayOfWeek1st > lDayOfWeek Then
        lOffset = lOffset + 7
    End If

    lOffset = lOffset + 7 * (lNth - 1)

    getNthWeeksDayOfWeek = DateAdd("d", lOffset, dt1stDate)

End Function

'//////////////////////////////////////////////////
'春分の日を取得
'//////////////////////////////////////////////////
Private Function getVernalEquinoxDay(ByVal lYear As Long) As Date

    Dim lDay    As Long

    lDay = Int(20.8431 + 0.242194 * (lYear - 1980) - Int((lYear - 1980) / 4))

    getVernalEquinoxDay = DateSerial(lYear, 3, lDay)

End Function

'//////////////////////////////////////////////////
'秋分の日を取得
'//////////////////////////////////////////////////
Private Function getAutumnalEquinoxDay(ByVal lYear As Long) As Date

    Dim lDay    As Long

    lDay = Int(23.2488 + 0.242194 * (lYear - 1980) - Int((lYear - 1980) / 4))

    getAutumnalEquinoxDay = DateSerial(lYear, 9, lDay)

End Function

Private Sub qSort(ByRef dtHolidays() As Date, ByVal lLeft As Long, ByVal lRight As Long)

    Dim dtCenter    As Date
    Dim dtTemp      As Date
    Dim i           As Long
    Dim j           As Long

    If lLeft < lRight Then
        dtCenter = dtHolidays((lLeft + lRight) \ 2)

        i = lLeft - 1
        j = lRight + 1

        Do While (True)
            i = i + 1
            Do While (dtHolidays(i) < dtCenter)
                i = i + 1
            Loop

            j = j - 1
            Do While (dtHolidays(j) > dtCenter)
                j = j - 1
            Loop

            If i >= j Then
                Exit Do
            End If

            dtTemp = dtHolidays(i)
            dtHolidays(i) = dtHolidays(j)
            dtHolidays(j) = dtTemp
        Loop

        Call qSort(dtHolidays, lLeft, i - 1)
        Call qSort(dtHolidays, j + 1, lRight)
    End If

End Sub

