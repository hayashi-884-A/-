CBtnHandler　　　クラスモジュール
'=== CBtnHandler.cls ===
Option Explicit

Public WithEvents Btn As MSForms.CommandButton
Public Parent As frmCalendarRange

Private Sub Btn_Click()
    If Not Parent Is Nothing Then
        Parent.HandleDayClick Btn
    End If
End Sub
ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

CCatHandler　　　クラスモジュール
'=== CCatHandler.cls ===
Option Explicit

Public WithEvents Btn As MSForms.CommandButton
Public Parent As frmCalendarRange

Private Sub Btn_Click()
    If Not Parent Is Nothing Then
        Parent.HandleCategoryClick Btn
    End If
End Sub
ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

frmCalendarRange　　　　　ユーザーフォーム

'=== frmCalendarRange ===
Option Explicit
Public ContinuousMode As Boolean  ' ← 連続入力モードかどうか
'===========================================================
' ★ レイアウト調整（必要ならここだけ。位置・サイズは固定のまま）
Private Const FORM_WIDTH As Single = 340
Private Const FORM_HEIGHT As Single = 520
Private Const FORM_SHIFT_X As Single = 0
Private Const FORM_SHIFT_Y As Single = 0
Private Const SHIFT_X As Single = 35        ' 中身の右シフト（左側に分類ボタンの余白を確保）
Private Const SHIFT_Y As Single = 0
Private Const RANGE_LINE_TOP_OFFSET As Single = 8
Private Const GAP_BELOW_RANGE As Single = 8
Private Const BOTTOM_GAP As Single = 14
'===========================================================

' 内部状態
Private mMonthDate As Date
Private mStartDate As Variant
Private mEndDate As Variant
Private mBtns As Collection              ' 日付ボタンのイベントハンドラ(CBtnHandler)を保持

' 動的に生成するコントロール
Private lblMonth As MSForms.label
Private WithEvents btnPrev As MSForms.CommandButton
Private WithEvents btnNext As MSForms.CommandButton
Private lblRange As MSForms.label
Private WithEvents txtMemo As MSForms.TextBox
Private WithEvents btnOK As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton
Private weekLbl(0 To 6) As MSForms.label

'=== サイドの分類ボタン（左側） ===
Private Const CAT_BTN_COUNT As Long = 10
Private Const CAT_BTN_W As Single = 26
Private Const CAT_BTN_H As Single = 22
Private Const CAT_BTN_GAP As Single = 4
Private mCatHandlers As Collection       ' 分類ボタンのイベントハンドラ(CCatHandler)
Private mSelectedCat As Long             ' 0=未選択, 1～10=選択中
Private mCatIdx(1 To CAT_BTN_COUNT) As Long  ' 分類1～10の色番号
Private mCatIdxSelected(1 To CAT_BTN_COUNT) As Long  ' ★追加：選択時の「色番号」
'=== Quick Palette：番号で配色を切替（ここだけ数字を変えればOK） ===
Private Const PALETTE_PRESET As String = "VIVID"  ' "VIVID" / "PASTEL" / "EARTH"
Private palette(1 To 22) As Long  ' 1～22（21/22は薄い青を追加）

' 役割 → 色番号（数字を変えるだけで色チェンジ）
Private Const IDX_BG              As Long = 16 ' フォーム背景
Private Const IDX_TEXT            As Long = 13 ' 基本文字
Private Const IDX_TEXT_SUB        As Long = 14 ' 補助文字
Private Const IDX_MONTH_TEXT      As Long = 13 ' 見出し(月名)
Private Const IDX_WEEK_BG         As Long = 16 ' 曜日帯の背景
Private Const IDX_WEEK_TEXT       As Long = 14 ' 曜日帯の文字
Private Const IDX_SUNDAY_TEXT     As Long = 6  ' 日曜文字
Private Const IDX_SATURDAY_TEXT   As Long = 2  ' 土曜文字
Private Const IDX_DAY_BG          As Long = 17 ' 日付セル通常背景
Private Const IDX_DAY_TEXT        As Long = 13 ' 日付セル通常文字
Private Const IDX_TODAY_BG        As Long = 11 ' 今日背景（薄い青：21/22が超薄）
Private Const IDX_RANGE_BG        As Long = 21 ' 間の期間背景（薄い青）
Private Const IDX_SELECTED_BG     As Long = 1  ' 開始/終了背景（濃いアクセント）
Private Const IDX_SELECTED_TEXT   As Long = 17 ' 開始/終了文字（白）
Private Const IDX_OTHERMONTH_BG   As Long = 16 ' 他月セル背景
Private Const IDX_OTHERMONTH_TEXT As Long = 14 ' 他月セル文字
Private Const IDX_NAV_BG          As Long = 15 ' ＜/＞ボタン背景
Private Const IDX_NAV_TEXT        As Long = 13 ' ＜/＞ボタン文字
Private Const IDX_OK_BG           As Long = 1  ' OK背景
Private Const IDX_OK_TEXT         As Long = 17 ' OK文字
Private Const IDX_CANCEL_BG       As Long = 15 ' キャンセル背景
Private Const IDX_CANCEL_TEXT     As Long = 13 ' キャンセル文字
Private Const IDX_MEMO_BG         As Long = 17 ' メモ背景
Private Const IDX_MEMO_TEXT       As Long = 13 ' メモ文字
Private Const IDX_RANGE_TEXT      As Long = 3  ' 「○月○日～○月○日」文字
' 分類ボタン（1～10）の背景色番号（CSVで指定）
Private Const CAT_BG_IDX_CSV As String = "6,1,5,3,7,8,2,4,11,19"
Private Const IDX_CAT_TEXT   As Long = 17      ' 分類ボタン文字（白推奨）
' ★ 追加：選択中の見せ方を設定
' 1) CSVで「選択中の背景色番号」を個別指定したい場合は下を埋める（例: "22,22,22,22,22,22,22,22,22,22"）
Private Const CAT_BG_IDX_SELECTED_CSV As String = ""   ' 空＝未使用 → 自動で“明るくする”

' 2) 自動で“明るくする”場合の割合（0~1）。0.2=20% 明るく。
Private Const CAT_SELECTED_LIGHTEN As Double = -0.6

' 実色（上の番号割り当てから自動で決定）
Private COLOR_BG               As Long, COLOR_TEXT As Long, COLOR_TEXT_SUB As Long
Private COLOR_MONTH_TEXT       As Long, COLOR_WEEK_BG As Long, COLOR_WEEK_TEXT As Long
Private COLOR_SUNDAY           As Long, COLOR_SATURDAY As Long
Private COLOR_DAY_BG           As Long, COLOR_DAY_TEXT As Long
Private COLOR_TODAY_BG         As Long, COLOR_RANGE_BG As Long
Private COLOR_SELECTED_BG      As Long, COLOR_SELECTED_TEXT As Long
Private COLOR_OTHERMONTH_BG    As Long, COLOR_OTHERMONTH_TEXT As Long
Private COLOR_NAV_BG           As Long, COLOR_NAV_TEXT As Long
Private COLOR_OK_BG            As Long, COLOR_OK_TEXT As Long
Private COLOR_CANCEL_BG        As Long, COLOR_CANCEL_TEXT As Long
Private COLOR_MEMO_BG          As Long, COLOR_MEMO_TEXT As Long
Private COLOR_RANGE_TEXT       As Long

' 呼び出し側へ返す公開プロパティ
Public SelectedStart As Date
Public SelectedEnd As Date
Public MemoText As String
Public SelectedCategory As Long     ' 左の分類（1～10、0=未選択）
Public ClickedOK As Boolean

' レイアウト定数（通常は触らない）
Private Const MARGIN As Single = 12
Private Const gap As Single = 4
Private Const CELL_W As Single = 30
Private Const CELL_H As Single = 24
Private Const HEADER_H As Single = 22
Private Const DOW_H As Single = 16

'================= 初期化 =================
Private Sub UserForm_Initialize()
    Me.caption = "日付範囲＋メモの入力"

    Me.Width = FORM_WIDTH
    Me.Height = FORM_HEIGHT

    ClickedOK = False
    mSelectedCat = 0
    SelectedCategory = 0

    mMonthDate = DateSerial(Year(Date), Month(Date), 1)
    Set mBtns = New Collection

    ' Excelウィンドウ中央に表示
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2 + FORM_SHIFT_X
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2 + FORM_SHIFT_Y

    InitThemeColors
    BuildUI
    BuildCalendar mMonthDate
    UpdateDateLabels
    ApplyTheme
End Sub

'================= UI構築 =================
Private Sub BuildUI()
    Dim marginLeft As Single, marginTop As Single
    marginLeft = MARGIN + SHIFT_X
    marginTop = MARGIN + SHIFT_Y

    ' --- 左の分類ボタン ---
    BuildCategoryPanel marginTop

    ' カレンダーグリッドの位置とサイズ
    Dim gridLeft As Single, gridTop As Single, gridW As Single, gridH As Single
    gridLeft = marginLeft
    gridTop = marginTop + HEADER_H + gap + DOW_H + gap
    gridW = 7 * CELL_W + 6 * gap
    gridH = 6 * CELL_H + 5 * gap

    ' 月ナビ（前月・表示月・翌月）
    Set btnPrev = Controls.Add("Forms.CommandButton.1", "btnPrev", True)
    With btnPrev
        .caption = "<"
        .Width = 28: .Height = HEADER_H
        .Left = marginLeft
        .Top = marginTop
        .TakeFocusOnClick = False
    End With

    Set lblMonth = Controls.Add("Forms.Label.1", "lblMonth", True)
    With lblMonth
        .caption = ""
        .TextAlign = fmTextAlignCenter
        .Width = gridW - (btnPrev.Width + 28 + gap * 2)
        .Height = HEADER_H
        .Left = btnPrev.Left + btnPrev.Width + gap
        .Top = marginTop + 2
        .Font.Bold = True
    End With

    Set btnNext = Controls.Add("Forms.CommandButton.1", "btnNext", True)
    With btnNext
        .caption = ">"
        .Width = 28: .Height = HEADER_H
        .Left = lblMonth.Left + lblMonth.Width + gap
        .Top = marginTop
        .TakeFocusOnClick = False
    End With

    ' 曜日ラベル（日～土）
    Dim i As Long, dowTop As Single
    dowTop = marginTop + HEADER_H + gap
    Dim dowNames As Variant
    dowNames = Array("日", "月", "火", "水", "木", "金", "土")
    For i = 0 To 6
        Set weekLbl(i) = Controls.Add("Forms.Label.1", "lblDow" & i, True)
        With weekLbl(i)
            .caption = CStr(dowNames(i))
            .TextAlign = fmTextAlignCenter
            .Width = CELL_W: .Height = DOW_H
            .Left = gridLeft + i * (CELL_W + gap)
            .Top = dowTop
        End With
    Next i

    ' 日付ボタン（6×7=42）
    Dim R As Long, c As Long, idx As Long
    For R = 0 To 5
        For c = 0 To 6
            idx = R * 7 + c + 1
            Dim B As MSForms.CommandButton
            Set B = Controls.Add("Forms.CommandButton.1", "btnDay" & idx, True)
            With B
                .caption = ""
                .Width = CELL_W: .Height = CELL_H
                .Left = gridLeft + c * (CELL_W + gap)
                .Top = gridTop + R * (CELL_H + gap)
                .Enabled = False
                .TakeFocusOnClick = False
            End With
            Dim hdl As CBtnHandler
            Set hdl = New CBtnHandler
            Set hdl.Btn = B
            Set hdl.Parent = Me
            mBtns.Add hdl
        Next c
    Next R

    ' 1行の選択表示ラベル
    Set lblRange = Controls.Add("Forms.Label.1", "lblRange", True)
    With lblRange
        .caption = "〇月〇日～〇月〇日"
        .AutoSize = True
        .WordWrap = False
        .Left = gridLeft
        .Top = gridTop + gridH + RANGE_LINE_TOP_OFFSET
        .Font.Size = 10
    End With

    ' メモ欄
    Set txtMemo = Controls.Add("Forms.TextBox.1", "txtMemo", True)
    With txtMemo
        .Multiline = True
        .WordWrap = True
        .EnterKeyBehavior = False   ' Enterは自前制御（Shift/Alt+Enterで改行）
        .Width = gridW
        .Left = gridLeft
        .Top = lblRange.Top + lblRange.Height + GAP_BELOW_RANGE
        .Height = 150
    End With

    ' 下部ボタン（下端固定）
    Set btnCancel = Controls.Add("Forms.CommandButton.1", "btnCancel", True)
    With btnCancel
        .caption = "キャンセル"
        .Width = 90: .Height = 24
        .Left = gridLeft
        .Top = Me.InsideHeight - BOTTOM_GAP - .Height
        .TakeFocusOnClick = False
            .Cancel = True         ' ← これを追加（Escでキャンセル）
    End With

    Set btnOK = Controls.Add("Forms.CommandButton.1", "btnOK", True)
    With btnOK
        .caption = "OK"
        .Width = 90: .Height = 24
        .Left = gridLeft + gridW - .Width
        .Top = btnCancel.Top
        .Default = True
        .TakeFocusOnClick = False
    End With

    ' メモ欄の高さを自動調整
    Dim memoHeight As Single
    memoHeight = btnCancel.Top - (gap * 2) - txtMemo.Top
    If memoHeight < 60 Then memoHeight = 60
    txtMemo.Height = memoHeight

    UpdateMonthCaption
End Sub

'================= カレンダー構築 =================
Private Sub BuildCalendar(ByVal baseMonth As Date)
    ' クリア
    Dim k As Long
    For k = 1 To mBtns.Count
        With mBtns(k).Btn
            .caption = ""
            .Enabled = False
            .Tag = vbNullString
        End With
    Next k

    ' 当月の開始位置と日数
    Dim y As Long, m As Long
    y = Year(baseMonth): m = Month(baseMonth)
    Dim firstDay As Date: firstDay = DateSerial(y, m, 1)
    Dim startCol As Long: startCol = Weekday(firstDay, vbSunday) - 1  ' 0=日曜
    Dim daysInMonth As Long: daysInMonth = Day(DateSerial(y, m + 1, 0))

    ' 埋め込み
    Dim d As Long, idx As Long
    idx = startCol + 1 ' コレクションは1基点
    For d = 1 To daysInMonth
        With mBtns(idx).Btn
            .caption = CStr(d)
            .Enabled = True
            .Tag = CStr(DateSerial(y, m, d))
        End With
        idx = idx + 1
    Next d

    UpdateMonthCaption
    RefreshDayButtonStyles
End Sub

Private Sub UpdateMonthCaption()
    lblMonth.caption = Format$(mMonthDate, "yyyy年m月")
End Sub

Private Sub UpdateDateLabels()
    Dim sFrom As String, sTo As String
    If IsEmpty(mStartDate) Then
        sFrom = "〇月〇日"
    Else
        sFrom = Format$(mStartDate, "m月d日")
    End If
    If IsEmpty(mEndDate) Then
        sTo = "〇月〇日"
    Else
        sTo = Format$(mEndDate, "m月d日")
    End If

    lblRange.caption = sFrom & "～" & sTo
    lblRange.AutoSize = True
    lblRange.WordWrap = False

    ' 1行ラベルの高さ変化に追従
    txtMemo.Top = lblRange.Top + lblRange.Height + GAP_BELOW_RANGE
    btnCancel.Top = Me.InsideHeight - BOTTOM_GAP - btnCancel.Height
    btnOK.Top = btnCancel.Top

    Dim mh As Single
    mh = btnCancel.Top - (gap * 2) - txtMemo.Top
    If mh < 60 Then mh = 60
    txtMemo.Height = mh
End Sub

'=== 日付クリック（CBtnHandlerから呼ばれる） ===
Public Sub HandleDayClick(ByVal B As MSForms.CommandButton)
    If Len(B.Tag) = 0 Then Exit Sub
    Dim dt As Date: dt = CDate(B.Tag)

    If IsEmpty(mStartDate) Then
        mStartDate = dt
    ElseIf IsEmpty(mEndDate) Then
        If dt < mStartDate Then
            mEndDate = mStartDate
            mStartDate = dt
        Else
            mEndDate = dt
        End If
    Else
        mStartDate = dt
        mEndDate = Empty
    End If

    UpdateDateLabels
    RefreshDayButtonStyles
End Sub

'=== 分類クリック（CCatHandlerから） ===
Public Sub HandleCategoryClick(ByVal B As MSForms.CommandButton)
    Dim idx As Long
    idx = CatIndexFromTag(B.Tag)
    If idx < 1 Or idx > CAT_BTN_COUNT Then Exit Sub

    mSelectedCat = idx
    SelectedCategory = idx
    RefreshCategoryButtonStyles
End Sub

'=== ナビゲーション ===
Private Sub btnPrev_Click()
    mMonthDate = DateAdd("m", -1, mMonthDate)
    BuildCalendar mMonthDate
End Sub

Private Sub btnNext_Click()
    mMonthDate = DateAdd("m", 1, mMonthDate)
    BuildCalendar mMonthDate
End Sub

Private Sub btnOK_Click()
    ' --- 入力値を拾う ---
    Dim memo As String, d1 As Date, d2 As Date, cat As Long
    memo = txtMemo.Text
    If Not IsEmpty(mStartDate) Then d1 = mStartDate
    If Not IsEmpty(mEndDate) Then d2 = mEndDate
    cat = SelectedCategory

    ' 「選択分類」への反映（ある場合）。無ければ B2 に退避
    On Error Resume Next
    ActiveSheet.Parent.Names("選択分類").RefersToRange.Value = cat
    If Err.Number <> 0 Then Err.Clear: ActiveSheet.Range("B2").Value = cat
    On Error GoTo 0

    If Me.ContinuousMode Then
        ' ▼連続入力モード：ここで保存までやって、フォームは閉じない
        On Error Resume Next
        ' ※ modCalendarRange の WriteMemoAndRange を Public にします（後述）
        modCalendarRange.WriteMemoAndRange memo, d1, d2, cat
        If Err.Number <> 0 Then
            ' もし参照に失敗したら文字列実行でフォールバック
            Application.Run "WriteMemoAndRange", memo, d1, d2, cat
            Err.Clear
        End If
        On Error GoTo 0

        ' 入力欄だけクリアして次の入力へ
        ClearInputsForNext_連続
        Exit Sub
    Else
        ' ▼単発モード：従来どおり値を返して閉じる
        ClickedOK = True
        SelectedStart = d1
        SelectedEnd = d2
        MemoText = memo
        Me.Hide
    End If
End Sub


Private Sub btnCancel_Click()
    ClickedOK = False
    Me.Hide
End Sub

'=== メモ欄：Enter/改行制御 ===
Private Sub txtMemo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Const SHIFT_MASK As Integer = 1
    Const ALT_MASK   As Integer = 4

    If KeyCode = vbKeyReturn Then
        If (Shift And SHIFT_MASK) <> 0 Or (Shift And ALT_MASK) <> 0 Then
            KeyCode = 0
            txtMemo.SelText = vbCrLf        ' Shift+Enter / Alt+Enter → 改行
        Else
            KeyCode = 0
            btnOK_Click                     ' Enter単独 → OK実行
        End If
    End If
End Sub

'================= テーマ（パレット→役割色） =================
Private Sub InitThemeColors()
    InitPalette
    AssignColorsByIndex
    InitCategoryColorIndexes
    InitCategorySelectedColorIndexes   ' ★ 追加：選択中の色番号（CSVが空なら全部0）
End Sub


Private Sub InitPalette()
    Select Case UCase$(PALETTE_PRESET)
    Case "PASTEL"
        palette(1) = RGB(173, 208, 255)
        palette(2) = RGB(184, 178, 255)
        palette(3) = RGB(162, 227, 192)
        palette(4) = RGB(201, 232, 158)
        palette(5) = RGB(255, 199, 146)
        palette(6) = RGB(255, 148, 148)
        palette(7) = RGB(255, 179, 219)
        palette(8) = RGB(219, 179, 255)
        palette(9) = RGB(173, 227, 230)
        palette(10) = RGB(181, 240, 255)
        palette(11) = RGB(255, 239, 124)
        palette(12) = RGB(206, 170, 125)
        palette(13) = RGB(50, 60, 75)
        palette(14) = RGB(130, 140, 155)
        palette(15) = RGB(230, 236, 242)
        palette(16) = RGB(250, 252, 255)
        palette(17) = RGB(255, 255, 255)
        palette(18) = RGB(0, 0, 0)
        palette(19) = RGB(231, 224, 255)
        palette(20) = RGB(255, 220, 190)
        palette(21) = RGB(205, 230, 255)  ' Lighter Pastel Blue
        palette(22) = RGB(235, 245, 255)  ' Ultra Pastel Blue
    Case "EARTH"
        palette(1) = RGB(20, 78, 120)
        palette(2) = RGB(85, 68, 137)
        palette(3) = RGB(40, 100, 70)
        palette(4) = RGB(134, 147, 35)
        palette(5) = RGB(210, 110, 30)
        palette(6) = RGB(176, 62, 62)
        palette(7) = RGB(184, 80, 120)
        palette(8) = RGB(125, 73, 142)
        palette(9) = RGB(18, 110, 115)
        palette(10) = RGB(25, 135, 160)
        palette(11) = RGB(190, 150, 40)
        palette(12) = RGB(133, 94, 66)
        palette(13) = RGB(40, 45, 52)
        palette(14) = RGB(115, 120, 128)
        palette(15) = RGB(226, 228, 232)
        palette(16) = RGB(246, 247, 248)
        palette(17) = RGB(255, 255, 255)
        palette(18) = RGB(0, 0, 0)
        palette(19) = RGB(54, 110, 75)
        palette(20) = RGB(196, 120, 75)
        palette(21) = RGB(180, 210, 224)  ' Soft Sky
        palette(22) = RGB(220, 238, 245)  ' Pale Sky
    Case Else ' VIVID（既定）
        palette(1) = RGB(58, 124, 232)
        palette(2) = RGB(73, 88, 214)
        palette(3) = RGB(30, 142, 104)
        palette(4) = RGB(128, 184, 34)
        palette(5) = RGB(242, 134, 38)
        palette(6) = RGB(220, 70, 70)
        palette(7) = RGB(224, 72, 140)
        palette(8) = RGB(156, 77, 202)
        palette(9) = RGB(32, 152, 162)
        palette(10) = RGB(44, 177, 217)
        palette(11) = RGB(250, 200, 50)
        palette(12) = RGB(156, 102, 46)
        palette(13) = RGB(40, 48, 60)
        palette(14) = RGB(115, 125, 140)
        palette(15) = RGB(230, 233, 238)
        palette(16) = RGB(248, 249, 252)
        palette(17) = RGB(255, 255, 255)
        palette(18) = RGB(0, 0, 0)
        palette(19) = RGB(231, 224, 255)
        palette(20) = RGB(207, 238, 216)
        palette(21) = RGB(200, 236, 252)  ' Light Cyan
        palette(22) = RGB(232, 246, 255)  ' Extra Light Cyan
    End Select
End Sub

Private Function p(ByVal i As Long) As Long
    If i < LBound(palette) Or i > UBound(palette) Then
        p = RGB(255, 0, 255) ' 異常時は目立つマゼンタ
    Else
        p = palette(i)
    End If
End Function

Private Sub AssignColorsByIndex()
    COLOR_BG = p(IDX_BG)
    COLOR_TEXT = p(IDX_TEXT)
    COLOR_TEXT_SUB = p(IDX_TEXT_SUB)
    COLOR_MONTH_TEXT = p(IDX_MONTH_TEXT)
    COLOR_WEEK_BG = p(IDX_WEEK_BG)
    COLOR_WEEK_TEXT = p(IDX_WEEK_TEXT)
    COLOR_SUNDAY = p(IDX_SUNDAY_TEXT)
    COLOR_SATURDAY = p(IDX_SATURDAY_TEXT)
    COLOR_DAY_BG = p(IDX_DAY_BG)
    COLOR_DAY_TEXT = p(IDX_DAY_TEXT)
    COLOR_TODAY_BG = p(IDX_TODAY_BG)
    COLOR_RANGE_BG = p(IDX_RANGE_BG)
    COLOR_SELECTED_BG = p(IDX_SELECTED_BG)
    COLOR_SELECTED_TEXT = p(IDX_SELECTED_TEXT)
    COLOR_OTHERMONTH_BG = p(IDX_OTHERMONTH_BG)
    COLOR_OTHERMONTH_TEXT = p(IDX_OTHERMONTH_TEXT)
    COLOR_NAV_BG = p(IDX_NAV_BG)
    COLOR_NAV_TEXT = p(IDX_NAV_TEXT)
    COLOR_OK_BG = p(IDX_OK_BG)
    COLOR_OK_TEXT = p(IDX_OK_TEXT)
    COLOR_CANCEL_BG = p(IDX_CANCEL_BG)
    COLOR_CANCEL_TEXT = p(IDX_CANCEL_TEXT)
    COLOR_MEMO_BG = p(IDX_MEMO_BG)
    COLOR_MEMO_TEXT = p(IDX_MEMO_TEXT)
    COLOR_RANGE_TEXT = p(IDX_RANGE_TEXT)
End Sub

'================= 見た目適用 =================
Private Sub ApplyTheme()
    ' ベース
    Me.BackColor = COLOR_BG
    On Error Resume Next
    Me.Font.name = "Meiryo UI"
    Me.Font.Size = 9
    On Error GoTo 0

    ' 見出し
    With lblMonth
        .ForeColor = COLOR_MONTH_TEXT
        .Font.Size = 12
        .Font.Bold = True
        .BackStyle = fmBackStyleTransparent
    End With

    ' 曜日ラベル
    Dim i As Long
    For i = 0 To 6
        With weekLbl(i)
            .BackStyle = fmBackStyleOpaque
            .BackColor = COLOR_WEEK_BG
            .ForeColor = COLOR_WEEK_TEXT
            .Font.Bold = True
        End With
    Next i
    weekLbl(0).ForeColor = COLOR_SUNDAY
    weekLbl(6).ForeColor = COLOR_SATURDAY

    ' 期間表示
    With lblRange
        .ForeColor = COLOR_RANGE_TEXT
        .BackStyle = fmBackStyleTransparent
        .Font.Bold = True
    End With

    ' メモ
    With txtMemo
        .BackColor = COLOR_MEMO_BG
        .ForeColor = COLOR_MEMO_TEXT
        .SpecialEffect = fmSpecialEffectSunken
    End With

    ' ナビボタン
    With btnPrev: .BackColor = COLOR_NAV_BG: .ForeColor = COLOR_NAV_TEXT: End With
    With btnNext: .BackColor = COLOR_NAV_BG: .ForeColor = COLOR_NAV_TEXT: End With

    ' 下部ボタン
    With btnOK
        .BackColor = COLOR_OK_BG
        .ForeColor = COLOR_OK_TEXT
        .Font.Bold = True
    End With
    With btnCancel
        .BackColor = COLOR_CANCEL_BG
        .ForeColor = COLOR_CANCEL_TEXT
    End With

    ' 日付ベース色
    Dim k As Long
    For k = 1 To mBtns.Count
        With mBtns(k).Btn
            .BackColor = COLOR_DAY_BG
            .ForeColor = COLOR_DAY_TEXT
        End With
    Next k

    RefreshDayButtonStyles
    RefreshCategoryButtonStyles
End Sub

'================= 日付ボタンの色更新 =================
Private Sub RefreshDayButtonStyles()
    Dim i As Long, dt As Date, dow As Long

    For i = 1 To mBtns.Count
        With mBtns(i).Btn
            .BackColor = COLOR_DAY_BG
            .ForeColor = COLOR_DAY_TEXT
            .Font.Bold = False

            If Len(.Tag) > 0 Then
                dt = CDate(.Tag)
                dow = Weekday(dt, vbSunday)

                If dow = vbSunday Then .ForeColor = COLOR_SUNDAY
                If dow = vbSaturday Then .ForeColor = COLOR_SATURDAY

                If dt = Date Then
                    .BackColor = COLOR_TODAY_BG
                    .Font.Bold = True
                End If

                If Not IsEmpty(mStartDate) Then
                    If IsEmpty(mEndDate) Then
                        If dt = mStartDate Then
                            .BackColor = COLOR_SELECTED_BG
                            .ForeColor = COLOR_SELECTED_TEXT
                            .Font.Bold = True
                        End If
                    Else
                        If dt >= mStartDate And dt <= mEndDate Then
                            .BackColor = COLOR_RANGE_BG
                        End If
                        If dt = mStartDate Or dt = mEndDate Then
                            .BackColor = COLOR_SELECTED_BG
                            .ForeColor = COLOR_SELECTED_TEXT
                            .Font.Bold = True
                        End If
                    End If
                End If
            Else
                .BackColor = COLOR_OTHERMONTH_BG
                .ForeColor = COLOR_OTHERMONTH_TEXT
            End If
        End With
    Next i
End Sub

'================= 分類ボタン =================
Private Sub BuildCategoryPanel(ByVal topBase As Single)
    Dim i As Long, leftPos As Single
    leftPos = MARGIN ' 左の余白内に縦配置（カレンダーは SHIFT_X で右へ）

    Set mCatHandlers = New Collection

    For i = 1 To CAT_BTN_COUNT
        Dim B As MSForms.CommandButton
        Set B = Controls.Add("Forms.CommandButton.1", "btnCat" & i, True)
        With B
            .caption = CStr(i)
            .ControlTipText = "分類" & CStr(i)
            .Width = CAT_BTN_W: .Height = CAT_BTN_H
            .Left = leftPos
            .Top = topBase + (i - 1) * (CAT_BTN_H + CAT_BTN_GAP)
            .TakeFocusOnClick = False
            .Tag = "CAT:" & CStr(i)
        End With

        Dim H As CCatHandler
        Set H = New CCatHandler
        Set H.Btn = B
        Set H.Parent = Me
        mCatHandlers.Add H
    Next i
End Sub

Private Sub InitCategoryColorIndexes()
    Dim arr() As String, i As Long
    arr = Split(CAT_BG_IDX_CSV, ",")
    For i = 1 To CAT_BTN_COUNT
        If i <= UBound(arr) + 1 Then
            mCatIdx(i) = Val(Trim$(arr(i - 1)))
        Else
            mCatIdx(i) = 15 ' 不足時は薄グレー
        End If
    Next i
End Sub
' ★ 追加：選択中の背景色番号を CSV から読み込み（空なら 0 のまま＝自動明度UP）
Private Sub InitCategorySelectedColorIndexes()
    Dim s As String: s = Trim$(CAT_BG_IDX_SELECTED_CSV)
    Dim i As Long
    If Len(s) = 0 Then
        For i = 1 To CAT_BTN_COUNT
            mCatIdxSelected(i) = 0
        Next
        Exit Sub
    End If

    Dim arr() As String: arr = Split(s, ",")
    For i = 1 To CAT_BTN_COUNT
        If i <= UBound(arr) + 1 Then
            mCatIdxSelected(i) = Val(Trim$(arr(i - 1)))
        Else
            mCatIdxSelected(i) = 0
        End If
    Next
End Sub


Private Function CatIndexFromTag(ByVal s As String) As Long
    If Left$(s, 4) = "CAT:" Then
        CatIndexFromTag = Val(Mid$(s, 5))
    Else
        CatIndexFromTag = 0
    End If
End Function

Private Sub RefreshCategoryButtonStyles()
    If mCatHandlers Is Nothing Then Exit Sub

    Dim i As Long
    For i = 1 To mCatHandlers.Count
        With mCatHandlers(i).Btn
            Dim idx As Long: idx = CatIndexFromTag(.Tag)
            If idx >= 1 And idx <= CAT_BTN_COUNT Then
                Dim baseColor As Long
                baseColor = p(mCatIdx(idx))                 ' 通常のカテゴリ色

                If idx = mSelectedCat Then
                    ' ★ 選択中：CSVに選択色があればそれ、無ければ明度UP
                    Dim bgSel As Long
                    If mCatIdxSelected(idx) > 0 Then
                        bgSel = p(mCatIdxSelected(idx))
                    Else
                        bgSel = LightenColor(baseColor, CAT_SELECTED_LIGHTEN)
                    End If

                    .BackColor = bgSel
                    .ForeColor = IdealTextColor(bgSel)      ' 背景に応じて白/濃グレーを自動
                    .Font.Bold = True
                Else
                    ' 非選択
                    .BackColor = baseColor
                    .ForeColor = p(IDX_CAT_TEXT)            ' 既定の文字色
                    .Font.Bold = False
                End If
            End If
        End With
    Next i
End Sub


' ★ 追加：色を「明るく」する（ratio=0.2 なら 20% 白に近づける）
' ratio の範囲：[-1, 1]
'   ratio > 0  … 白に近づけて「明るく」
'   ratio = 0  … 変更なし
'   ratio < 0  … 黒に近づけて「暗く」
Private Function LightenColor(ByVal c As Long, ByVal ratio As Double) As Long
    If ratio < -1 Then ratio = -1
    If ratio > 1 Then ratio = 1

    Dim R As Long, G As Long, B As Long
    R = (c And &HFF&)
    G = (c And &HFF00&) \ &H100&
    B = (c And &HFF0000) \ &H10000

    If ratio >= 0 Then
        ' 明るく（白方向へ線形補間）
        R = CLng(R + (255 - R) * ratio)
        G = CLng(G + (255 - G) * ratio)
        B = CLng(B + (255 - B) * ratio)
    Else
        ' 暗く（黒方向へ縮小）
        Dim f As Double: f = 1 + ratio   ' 例: -0.35 → 0.65 倍
        R = CLng(R * f)
        G = CLng(G * f)
        B = CLng(B * f)
    End If

    If R < 0 Then R = 0
    If G < 0 Then G = 0
    If B < 0 Then B = 0
    If R > 255 Then R = 255
    If G > 255 Then G = 255
    If B > 255 Then B = 255

    LightenColor = RGB(R, G, B)
End Function


' ★ 追加：背景色に対して見やすい文字色（白 or 濃グレー）を返す
Private Function IdealTextColor(ByVal bg As Long) As Long
    Dim R As Long, G As Long, B As Long
    R = (bg And &HFF&)
    G = (bg And &HFF00&) \ &H100&
    B = (bg And &HFF0000) \ &H10000
    ' 知覚輝度（おおよそ）：0～255
    Dim luma As Double
    luma = 0.299 * R + 0.587 * G + 0.114 * B
    If luma >= 160# Then
        IdealTextColor = p(13)   ' 明るい背景には濃グレー
    Else
        IdealTextColor = p(17)   ' 暗い背景には白
    End If
End Function
' フォーム全体で Enter と Esc を捕まえる
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Const SHIFT_MASK As Integer = 1
    Const ALT_MASK   As Integer = 4

    ' Enter：Shift/Alt を伴わない場合は OK を実行
    If KeyCode = vbKeyReturn Then
        If (Shift And SHIFT_MASK) = 0 And (Shift And ALT_MASK) = 0 Then
            KeyCode = 0
            btnOK_Click          ' ← 既存の OK クリック処理を直接呼ぶ
            Exit Sub
        End If
        ' Shift+Enter / Alt+Enter は txtMemo 側の処理（改行）に流す
    End If

    ' Esc：キャンセル
    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        btnCancel_Click
    End If
End Sub
'==== Enter/Esc を OnKey で割り当て（フォームが見えている間だけ）====
Private Sub UserForm_Activate()
    Application.OnKey "~", "FormEnterHook"      ' Enter
    Application.OnKey "{ESC}", "FormEscHook"    ' Esc
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' フォームを閉じる時は必ず解除
    Application.OnKey "~"
    Application.OnKey "{ESC}"
End Sub

' 標準モジュールから呼ぶための公開ラッパー
Public Sub DoOK():     btnOK_Click:     End Sub
Public Sub DoCancel(): btnCancel_Click: End Sub

'=== 事前入力を受け取るエントリポイント ===
Public Sub LoadDefaults(ByVal memo As String, ByVal d1 As Variant, ByVal d2 As Variant, ByVal cat As Variant)
    ' メモ
    SetIfExists "txtMemo", memo
    SetIfExists "tbMemo", memo
    SetIfExists "Memo", memo

    ' 日付（DatePicker や TextBox どちらにも対応／存在しないコントロールは無視）
    If IsDate(d1) Then
        SetIfExists "dtpStart", CDate(d1)
        SetIfExists "txtStart", Format$(CDate(d1), DATE_FMT_015)
        SetIfExists "tbStart", Format$(CDate(d1), DATE_FMT_015)
    End If
    If IsDate(d2) Then
        SetIfExists "dtpEnd", CDate(d2)
        SetIfExists "txtEnd", Format$(CDate(d2), DATE_FMT_015)
        SetIfExists "tbEnd", Format$(CDate(d2), DATE_FMT_015)
    End If

    ' 分類（1～10）：コンボ／スピン／オプションボタンのどれがあっても対応
    If IsNumeric(cat) Then
        Dim k As Long: k = CLng(cat)
        SetIfExists "cmbCategory", k
        SetIfExists "txtCategory", CStr(k)
        SetIfExists "spnCategory", k
        ' オプションボタン optCat1..optCat10 を想定
        If k >= 1 And k <= 10 Then SetIfExists "optCat" & CStr(k), True
    End If
End Sub

'--- 存在すれば .Value/.Text を両方試して代入する汎用セッタ ---
Private Sub SetIfExists(ByVal ctrlName As String, ByVal v As Variant)
    On Error Resume Next
    Me.Controls(ctrlName).Value = v
    Me.Controls(ctrlName).Text = CStr(v)
    On Error GoTo 0
End Sub

Private Sub ClearInputsForNext_連続()
    ' 期間・分類・メモを空に戻す（＝次の入力待ち）
    mStartDate = Empty
    mEndDate = Empty
    txtMemo.Text = ""

    mSelectedCat = 0
    SelectedCategory = 0

    ' 見た目更新
    UpdateDateLabels
    RefreshDayButtonStyles
    RefreshCategoryButtonStyles

    ' 次の入力開始位置（お好みで変更：例 メモ欄にフォーカス）
    On Error Resume Next
    txtMemo.SetFocus
    On Error GoTo 0
End Sub





ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー



modAutomation　　　　標準モジュール

Option Explicit

'=== ① 選択行の値を事前入力してフォームを開く（OKで「選択行」を上書き） ===
Public Sub Automation_ShowForm_EditSelectedRow()
    On Error GoTo Clean
    Application.ScreenUpdating = False

    Dim ws As Worksheet: Set ws = ActiveSheet

    If TypeName(Selection) = "Range" Then
        Dim R As Long: R = Selection.Row
        If R >= DATA_START_ROW_015 Then
            Dim memo As String, cat As Variant
            Dim d1 As Variant, d2 As Variant

            memo = CStr(ws.Cells(R, "B").Value)      ' イベント名／内容
            cat = ws.Cells(R, "A").Value              ' 分類番号（1～10）
            d1 = GetDateOrEmpty(ws.Cells(R, "D").Value) ' 開始日
            d2 = GetDateOrEmpty(ws.Cells(R, "E").Value) ' 終了日

            ' D/E が空で、C に旧「m/d~m/d」書式が残っている場合のみ救済
            If IsEmpty(d1) And IsEmpty(d2) Then
                Dim t0 As Date, t1 As Date
                If ParseDateRange_015(CStr(ws.Cells(R, "C").Value), t0, t1) Then
                    d1 = t0: d2 = t1
                End If
            End If

' 表示前に必ず破棄（常に新品から）
On Error Resume Next
Unload frmCalendarRange
On Error GoTo 0

With frmCalendarRange
    On Error Resume Next
    .LoadDefaults memo, d1, d2, cat   ' 事前入力
    On Error GoTo 0
    .Show

    If .ClickedOK Then
        WriteMemoAndRange_ToRow ws, R, .MemoText, .SelectedStart, .SelectedEnd, .SelectedCategory
        SetupPeriodConditionalFormatting_020
        ApplyCategoryDividers_020 ws
        GreyOutPastDateColumns_041 ws, CAL_DATE_ROW_015
    End If
End With

' 表示後も必ず破棄（次回に持ち越さない）
On Error Resume Next
Unload frmCalendarRange
On Error GoTo 0

            GoTo Clean
        End If
    End If

    ' 選択が無効（ヘッダ行など）のときは ② の挙動にフォールバック
    Call Automation_ShowForm_Blank

Clean:
    Application.ScreenUpdating = True
End Sub

'=== ② まっさらのフォームで開く（従来どおり新規追加） ===
Public Sub Automation_ShowForm_Blank()
    On Error GoTo Clean
    Application.ScreenUpdating = False

    Dim ws As Worksheet: Set ws = ActiveSheet

    ' 従来フロー（空のフォーム → OKで新規行に B/D/E/A を書く）
    ShowCalendarForm

    ' 見た目と帯を再適用
    SetupPeriodConditionalFormatting_020
    ApplyCategoryDividers_020 ws
    GreyOutPastDateColumns_041 ws, CAL_DATE_ROW_015

Clean:
    Application.ScreenUpdating = True
End Sub

'--- ヘルパー：日付なら Date、それ以外は Empty を返す ---
Private Function GetDateOrEmpty(ByVal v As Variant) As Variant
    If IsDate(v) Then
        GetDateOrEmpty = CDate(v)
    Else
        GetDateOrEmpty = Empty
    End If
End Function



ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー


modCalendarRange　　　標準モジュール



'=== modCalendarRange ===
Option Explicit

' ▼必要に応じて変更
Private Const TARGET_SHEET_NAME As String = "" ' ""=アクティブシート。固定するならシート名を指定。
Private Const TARGET_COL As Long = 2           ' B列=2（内容）
Private Const START_ROW As Long = 6            ' 6行目開始（B6から）

Public Sub ShowCalendarForm()
    ' 表示のたびに既定インスタンスを破棄（前回値の残留対策）
    On Error Resume Next
    Unload frmCalendarRange
    On Error GoTo 0

    ' ★ここが抜けていた：実際にフォームを表示する
    frmCalendarRange.Show

    ' OKが押された場合のみ、入力内容を反映
   If frmCalendarRange.ClickedOK Then
    Dim memo As String
    memo = Trim$(frmCalendarRange.MemoText)

    Dim d1 As Date, d2 As Date
    d1 = frmCalendarRange.SelectedStart
    d2 = frmCalendarRange.SelectedEnd

    ' ★メモが空でも書き込む
    WriteMemoAndRange memo, d1, d2, frmCalendarRange.SelectedCategory
End If


    ' 表示後は必ず破棄（次回は新品で始める）
    On Error Resume Next
    Unload frmCalendarRange
    On Error GoTo 0
End Sub


Public Sub WriteMemoAndRange(ByVal memo As String, ByVal d1 As Date, ByVal d2 As Date, ByVal cat As Long)

    Dim ws As Worksheet
    Dim R As Long
    Dim T As Date
    Dim ev As Boolean

    Set ws = ResolveTargetSheet(TARGET_SHEET_NAME)

    ev = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo CleanUp

    ' 次の空行（B列基準）
    R = FirstEmptyRow(ws, TARGET_COL, START_ROW)

    ' 1) B列：内容
    ws.Cells(R, TARGET_COL).Value = memo

    ' 2) D/E列：開始/終了（未入力補完＋大小入替）→ 同時書き込み
    If d1 <> 0 Or d2 <> 0 Then
        If d1 = 0 Then d1 = d2
        If d2 = 0 Then d2 = d1
        If d2 < d1 Then T = d1: d1 = d2: d2 = T

With ws
    .Cells(R, "D").Value = d1
    .Cells(R, "D").NumberFormatLocal = DATE_FMT_015
    .Cells(R, "E").Value = d2
    .Cells(R, "E").NumberFormatLocal = DATE_FMT_015
End With
    End If

    ' 3) A列：分類（ここで並べ替えまで実施）
    If cat > 0 Then
        SetCategoryToRow_017 ws, R, cat, True
        On Error Resume Next
        Application.Run "ApplyPriorityColors_AllRows_016", ws
        On Error GoTo 0
    End If

CleanUp:
    Application.EnableEvents = ev   ' ★必ず復帰

    ' 4) 帯のCFと区切り線（見た目を確定）
    SetupPeriodConditionalFormatting_020
    ApplyCategoryDividers_020 ws
End Sub




Private Function FirstEmptyRow(ByVal ws As Worksheet, ByVal col As Long, ByVal startRow As Long) As Long
    Dim R As Long: R = startRow
    Do While _
        (LenB(ws.Cells(R, "B").Value2) <> 0) Or _
        (LenB(ws.Cells(R, "D").Value2) <> 0) Or _
        (LenB(ws.Cells(R, "E").Value2) <> 0)
        R = R + 1
    Loop
    FirstEmptyRow = R
End Function


'--- シート解決：名前が空なら ActiveSheet、指定があればそのシート（なければ ActiveSheet） ---
Private Function ResolveTargetSheet(ByVal sheetName As String) As Worksheet
    If Len(sheetName) = 0 Then
        Set ResolveTargetSheet = ActiveSheet
    Else
        On Error Resume Next
        Set ResolveTargetSheet = ThisWorkbook.Worksheets(sheetName)
        On Error GoTo 0
        If ResolveTargetSheet Is Nothing Then
            Set ResolveTargetSheet = ActiveSheet
        End If
    End If
End Function


'--- 指定行を上書き（B=メモ / D=開始 / E=終了 / A=分類）。Cは空欄維持 ---
Public Sub WriteMemoAndRange_ToRow(ByVal ws As Worksheet, ByVal R As Long, _
                                   ByVal memo As String, ByVal d1 As Date, ByVal d2 As Date, ByVal cat As Long)
    Dim T As Date
    If ws Is Nothing Then Set ws = ActiveSheet
    If R < DATA_START_ROW_015 Then Exit Sub

    Application.EnableEvents = False

    ' 1) メモ（B）
    ws.Cells(R, "B").Value = memo

    ' 2) 日付（D/E）※未入力補完・大小入替
    If d1 = 0 And d2 = 0 Then
        ' 何もしない（D/E 既存値を尊重）
    Else
        If d1 = 0 Then d1 = d2
        If d2 = 0 Then d2 = d1
        If d2 < d1 Then T = d1: d1 = d2: d2 = T
        ws.Cells(R, "D").Value = d1: ws.Cells(R, "D").NumberFormatLocal = DATE_FMT_015
        ws.Cells(R, "E").Value = d2: ws.Cells(R, "E").NumberFormatLocal = DATE_FMT_015
    End If

    ' 3) C は常に空欄（メモ専用のため自動は書かない）
    ws.Cells(R, "C").ClearContents

   ' 4) 分類（A）
If cat > 0 Then
    SetCategoryToRow_017 ws, R, cat, True   ' ← ここで並べ替えまで実施
    ' ★追加：カテゴリ色をB列に再適用（全行）
    On Error Resume Next
    Application.Run "ApplyPriorityColors_AllRows_016", ws
    On Error GoTo 0
End If

Application.EnableEvents = True

' 5) 帯のCFと区切り線を貼り直し（B列の色が決まった“後”に呼ぶことが重要）
SetupPeriodConditionalFormatting_020
ApplyCategoryDividers_020 ws

End Sub



ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー




modRowColorsFix　　　標準モジュール


Option Explicit

Public Sub ApplyPriorityColors_AllRows_016(ByVal ws As Worksheet)
    If ws Is Nothing Then Set ws = ActiveSheet

    Const MAX_CAT As Long = 10
    Const CHILD_LIGHTEN As Double = 0.35     ' ← 子の薄さ（0.0～1.0）

Dim firstRow As Long: firstRow = DATA_START_ROW_015
Dim lastRow  As Long
lastRow = Application.Max( _
             ws.Cells(ws.Rows.Count, "A").End(xlUp).Row, _
             ws.Cells(ws.Rows.Count, "B").End(xlUp).Row, _
             ws.Cells(ws.Rows.Count, "C").End(xlUp).Row, _
             ws.Cells(ws.Rows.Count, "D").End(xlUp).Row, _
             ws.Cells(ws.Rows.Count, "E").End(xlUp).Row, _
             firstRow)
If lastRow < firstRow Then Exit Sub


    ' ===== 指定パレット =====
    ' 1=赤, 2=ライトブルー, 3=オレンジ, 4=緑, 5=ピンク,
    ' 6=紫, 7=紺, 8=黄緑, 9=黄, 10=薄い灰
    Dim catCol(1 To MAX_CAT) As Long
    catCol(1) = RGB(220, 53, 69)      ' 1 赤
    catCol(2) = RGB(0, 153, 204)      ' 2 ライトブルー
    catCol(3) = RGB(255, 165, 0)      ' 3 オレンジ
    catCol(4) = RGB(76, 175, 80)      ' 4 緑
    catCol(5) = RGB(255, 99, 132)     ' 5 ピンク
    catCol(6) = RGB(102, 0, 153)      ' 6 紫
    catCol(7) = RGB(0, 51, 102)       ' 7 紺
    catCol(8) = RGB(146, 208, 80)     ' 8 黄緑
    catCol(9) = RGB(255, 215, 0)      ' 9 黄
catCol(10) = RGB(150, 75, 0)      ' 10 ★茶色 (#964B00)

    Dim R As Long, curA As Variant, prevA As Variant
    Dim isParent As Boolean, baseColor As Long, paintColor As Long

    prevA = Empty
    Application.ScreenUpdating = False

    For R = firstRow To lastRow
        curA = ws.Cells(R, "A").Value
        If IsNumeric(curA) And curA >= 1 And curA <= MAX_CAT Then
            ' 親=ブロック先頭, 子=2行目以降
            isParent = Not (IsNumeric(prevA) And CLng(prevA) = CLng(curA))
            baseColor = catCol(CLng(curA))
            paintColor = IIf(isParent, baseColor, LightenColor_016(baseColor, CHILD_LIGHTEN))
            With ws.Range(ws.Cells(R, "B"), ws.Cells(R, "D")).Interior
                .Pattern = xlSolid
                .Color = paintColor
            End With
        Else
            ws.Range(ws.Cells(R, "B"), ws.Cells(R, "D")).Interior.Pattern = xlNone
        End If
        prevA = curA
    Next R

    Application.ScreenUpdating = True
End Sub

' 色を明るく（0.0～1.0で白に近づける）
Private Function LightenColor_016(ByVal c As Long, ByVal p As Double) As Long
    If p < 0 Then p = 0
    If p > 1 Then p = 1
    Dim R As Long, G As Long, B As Long
    R = (c And &HFF)
    G = (c \ &H100) And &HFF
    B = (c \ &H10000) And &HFF
    R = R + CLng((255 - R) * p)
    G = G + CLng((255 - G) * p)
    B = B + CLng((255 - B) * p)
    LightenColor_016 = RGB(R, G, B)
End Function



ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー


modSeparators　　　標準モジュール


'=== 分類の境目に二重線（横方向にB～カレンダー末尾まで） ===
Public Sub DrawCategorySeparators_030(ByVal ws As Worksheet)
    Dim calStartCol As Long: calStartCol = CAL_START_COL_015
    Dim calDateRow  As Long: calDateRow = CAL_DATE_ROW_015
    Dim dataStartRow As Long: dataStartRow = DATA_START_ROW_015

    Dim lastCalCol As Long
    lastCalCol = ws.Cells(calDateRow, ws.Columns.Count).End(xlToLeft).Column

    Dim lastRow As Long
lastRow = Application.Max( _
    ws.Cells(ws.Rows.Count, "A").End(xlUp).Row, _
    ws.Cells(ws.Rows.Count, "B").End(xlUp).Row, _
    ws.Cells(ws.Rows.Count, "D").End(xlUp).Row, _
    ws.Cells(ws.Rows.Count, "E").End(xlUp).Row, _
    dataStartRow)

    If lastRow < dataStartRow Then Exit Sub

    ' いったん既存の上線をクリア（B～最右）
    Dim rngAll As Range
    Set rngAll = ws.Range(ws.Cells(dataStartRow, 2), ws.Cells(lastRow, lastCalCol))
    rngAll.Borders(xlEdgeTop).LineStyle = xlNone

    ' 先頭行にも線を入れる（見やすさ）
    ws.Range(ws.Cells(dataStartRow, 2), ws.Cells(dataStartRow, lastCalCol)) _
        .Borders(xlEdgeTop).LineStyle = xlDouble

    ' 分類が変わる行に二重線
    Dim R As Long, aPrev As Variant, aNow As Variant
    aPrev = ws.Cells(dataStartRow, "A").Value
    For R = dataStartRow + 1 To lastRow
        aNow = ws.Cells(R, "A").Value
        If CStr(aNow) <> CStr(aPrev) Then
            With ws.Range(ws.Cells(R, 2), ws.Cells(R, lastCalCol)).Borders(xlEdgeTop)
                .LineStyle = xlDouble
                .weight = xlThick
                .Color = vbBlack
            End With
        End If
        aPrev = aNow
    Next R
End Sub




ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー



Module1　　標準モジュール

Option Explicit

Public Sub FormEnterHook()
    On Error Resume Next
    If frmCalendarRange.Visible Then frmCalendarRange.DoOK
End Sub

Public Sub FormEscHook()
    On Error Resume Next
    If frmCalendarRange.Visible Then frmCalendarRange.DoCancel
End Sub



ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー



Sheet1　　シートモジュール

Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ExitHandler
    If Target Is Nothing Then Exit Sub

    Dim hit As Range

'===== A列：分類番号を変更したとき（手入力/貼り付けの両方） =====
Set hit = Intersect(Target, Me.Columns("A"))
If Not hit Is Nothing Then
    Application.EnableEvents = False
    Dim a As Range, v As Variant
    Dim needResortOnEmptyA As Boolean    ' ★追加：Aが空になったら並び替えを走らせるフラグ

    For Each a In hit.Cells
        If a.Row >= DATA_START_ROW_015 Then
            v = Trim$(a.Value & "")
            Select Case True
                Case Len(v) = 0
                    ' 空になった行は見た目を初期化（行色など）
                    WipeRowColors_016 Me, a.Row
                    RemoveChildrenIfParentCleared_016 Me, a.Row
                    needResortOnEmptyA = True          ' ★削除があったので後で全体ソート

                Case IsNumeric(v) And CLng(v) >= 1
                    ' Aを確定 → A→Dの二段ソート（従来どおり）
                    Me.Cells(a.Row, "A").Value = CLng(v)
                    ResortByCategory_031 Me

                Case Else
                    a.ClearContents: Beep
            End Select
        End If
    Next a

    ' ★追加：Aが空になった行が1つでもあれば、全体の既定順で再ソート
    If needResortOnEmptyA Then
        ResortAllByCategory_032 Me        ' 例：A→D→E のキー順
    End If

    ' ★追加：行色（B～D）を先に再塗装
'modRowColorsFix.ApplyPriorityColors_AllRows_016 Me


    ' ★順番重要：帯（期間）の条件付き書式を次に再構築
    SetupPeriodConditionalFormatting_020
' ★追加：今日より前の列をグレー化
GreyOutPastDateColumns_041 Me, CAL_DATE_ROW_015

    ' ★最後に分類の区切り線（二重線）
    ApplyCategoryDividers_020 Me

    Application.EnableEvents = True
End If


'===== B列：イベント名が空なら行ワイプ（A/C/色/帯）＋ 並び替え＋ 区切り線/CF再生成 =====
Set hit = Intersect(Target, Me.Columns("B"))
If Not hit Is Nothing Then
    Application.EnableEvents = False
    Dim rB As Range
    For Each rB In hit.Cells
        If rB.Row >= DATA_START_ROW_015 Then
            If Len(Trim$(rB.Value & "")) = 0 Then
                WipeRowOnEmptyB_016 Me, rB.Row
            End If
        End If
    Next rB

    ' 並び替え
    ResortAllByCategory_032 Me

    ' ★追加：行色（B～D）を再塗装
'modRowColorsFix.ApplyPriorityColors_AllRows_016 Me

    ' ★順番重要：帯CF → 区切り線
    SetupPeriodConditionalFormatting_020
    ' ★追加：今日より前の列をグレー化
GreyOutPastDateColumns_041 Me, CAL_DATE_ROW_015

    ApplyCategoryDividers_020 Me

    Application.EnableEvents = True
End If


'===== D列：開始日編集（年補完＆正規化）→ 並び替え＋行色→帯CF→区切り線 =====
Set hit = Intersect(Target, Me.Columns("D"))
If Not hit Is Nothing Then
    Application.EnableEvents = False
    Dim c As Range, s As String
    For Each c In hit.Cells
        If c.Row >= DATA_START_ROW_015 Then
            s = NormalizeDateRangeWithYearFromCalendar_015(CStr(c.Value), Me, CAL_START_COL_015, CAL_DATE_ROW_015)
            If Len(s) > 0 Then c.Value = s
        End If
    Next c

    ' 並び替え
    ResortAllByCategory_032 Me

    ' ★追加：行色（B～D）を再塗装
'modRowColorsFix.ApplyPriorityColors_AllRows_016 Me


    ' ★順番重要：帯CF → 区切り線
    SetupPeriodConditionalFormatting_020
    ' ★追加：今日より前の列をグレー化
GreyOutPastDateColumns_041 Me, CAL_DATE_ROW_015

    ApplyCategoryDividers_020 Me

    Application.EnableEvents = True
End If

'===== E列：終了日編集 → 並び替え＋行色→帯CF→区切り線 =====
Set hit = Intersect(Target, Me.Columns("E"))
If Not hit Is Nothing Then
    Application.EnableEvents = False

    ' 並び替え
    ResortAllByCategory_032 Me

    ' ★追加：行色（B～D）を再塗装
'modRowColorsFix.ApplyPriorityColors_AllRows_016 Me


    ' ★順番重要：帯CF → 区切り線
    SetupPeriodConditionalFormatting_020
    ' ★追加：今日より前の列をグレー化
GreyOutPastDateColumns_041 Me, CAL_DATE_ROW_015

    ApplyCategoryDividers_020 Me

    Application.EnableEvents = True
End If


ExitHandler:
    Application.EnableEvents = True
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo ExitHandler

    ' カレンダーの列（E列～最終カレンダー列）上で選択が動いたときだけ更新
    Dim firstCol As Long: firstCol = CAL_START_COL_015        ' 5 = E列
    Dim lastCol  As Long: lastCol = Me.Cells(CAL_DATE_ROW_015, Me.Columns.Count).End(xlToLeft).Column

    If Intersect(Target, Me.Range(Me.Cells(1, firstCol), Me.Cells(Me.Rows.Count, lastCol))) Is Nothing Then
        Exit Sub
    End If

    Application.EnableEvents = False
    UpdateMonthIndicator_B4_015 Me, CAL_START_COL_015         ' B4 を yyyy年m月 に更新
ExitHandler:
    Application.EnableEvents = True
End Sub

ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー



ThisWorkbook　　　




Option Explicit

' 次の「任意：真夜中自動更新」で使う（使わなければ削ってOK）
Private nextMidnight As Date

Private Sub Workbook_Open()
    ' ① 開いた直後に一回だけ灰色化
    GreyOutEveryCalendarSheet
    ' ②（任意）日付が変わったら自動更新したい場合
    ' ScheduleMidnightRefresh
End Sub

' すべてのワークシートのうち、カレンダーっぽいものだけ灰色化
Public Sub GreyOutEveryCalendarSheet()
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        On Error Resume Next
        ' E4 が日付なら “そのシートはカレンダー” とみなす
        If IsDate(sh.Cells(CAL_DATE_ROW_015, CAL_START_COL_015).Value) Then
            GreyOutPastDateColumns_041 sh, CAL_DATE_ROW_015
        End If
        On Error GoTo 0
    Next sh
End Sub

' ========= 任意：真夜中に自動再実行 =========
Private Sub ScheduleMidnightRefresh()
    On Error Resume Next
    nextMidnight = Date + 1  ' 明日の 0:00
    Application.OnTime nextMidnight, "GreyOutEveryCalendarSheet"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    If nextMidnight <> 0 Then
        Application.OnTime nextMidnight, "GreyOutEveryCalendarSheet", , False
    End If
End Sub



ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー



ジャンプ　　標準モジュール

Option Explicit

Public Sub ジャンプ()
    Dim ws As Worksheet
    Dim rowRange As Range
    Dim c As Range
    Dim found As Range

    ' 対象シート：必要なら Worksheets("シート名") に変更
    Set ws = ActiveSheet

    ' ４行目のうち「使用中の列」だけを見る
    Set rowRange = Intersect(ws.Rows(4), ws.UsedRange)
    If rowRange Is Nothing Then
        MsgBox "４行目にデータがありません。", vbExclamation
        Exit Sub
    End If

    ' 今日と一致する日付セルを検索（時刻は無視）
    For Each c In rowRange.Cells
        If Not IsError(c.Value) And Len(c.Value) > 0 Then
            If IsDate(c.Value) Then
                If DateValue(c.Value) = Date Then
                    Set found = c
                    Exit For
                End If
            End If
        End If
    Next c

    If found Is Nothing Then
        MsgBox "今日（" & Format$(Date, "yyyy/mm/dd") & "）に一致するセルが見つかりませんでした。", vbInformation
    Else
        Application.GoTo found, True  ' セルが見える位置までスクロール
        found.Select                  ' セルにフォーカス
    End If
End Sub

Public Sub 連続入力モード_起動()
    On Error Resume Next
    Unload frmCalendarRange           ' 前回の残りを念のため破棄
    On Error GoTo 0

    With frmCalendarRange
        .ContinuousMode = True        ' ★連続モードON
        .Show vbModal                 ' キャンセルが押されるまで開きっぱなし
    End With

    On Error Resume Next
    Unload frmCalendarRange           ' 終了時は破棄
    On Error GoTo 0
End Sub




ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー


表　　標準モジュール

Option Explicit

' ===== 設定（必要なら変えてOK） =====
Public Const CAL_START_COL_015 As Long = 6    ' F列
Public Const CAL_DATE_ROW_015  As Long = 4    ' 日付行
Public Const CAL_WEEK_ROW_015  As Long = 5    ' 曜日行
Public Const DATA_START_ROW_015 As Long = 6   ' データ表の開始行（C=開始日, D=終了日）
Public Const MONTH_SPAN_015    As Long = 24   ' 並べる月数（標準：24か月 = 2年）
Public Const DATE_FMT_015      As String = "yyyy年m月d日" ' C/Dの日付表示形式
' === 共通設定 ===
Public Const CAT_SELECTED_ADDR_017 As String = "B2"   ' 分類番号を読むセル（必要なら "E2" などに変更）
' 帯用：色を明るく（0.0～1.0で白に近づける）
Private Function Lighten_020(ByVal c As Long, ByVal p As Double) As Long
    Dim R As Long, G As Long, B As Long
    R = (c And &HFF): G = (c \ &H100) And &HFF: B = (c \ &H10000) And &HFF
    R = R + CLng((255 - R) * p)
    G = G + CLng((255 - G) * p)
    B = B + CLng((255 - B) * p)
    Lighten_020 = RGB(R, G, B)
End Function


' ===== オーケストレーター（これ1本で外枠まで） =====
Public Sub 表作成()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Application.ScreenUpdating = False

    ' 1) トップボタン作成（A1/B1/A2/B2/A3/B3）
    MakeTopbarButtons6_015 ws

    ' 2) ペイン固定（A～D列 + 1～5行）
    Freeze_AB_Rows1to5_015

    ' 3) カレンダー生成（E4 から MONTH_SPAN_015 か月分）
    MakeLongCalendar_015 ws, CAL_START_COL_015, CAL_DATE_ROW_015, CAL_WEEK_ROW_015, MONTH_SPAN_015

    ' 4) B4 の初期表示更新
    UpdateMonthIndicator_B4_015 ws, CAL_START_COL_015

    ' 5) 既存データの正規化（Cに「開始～終了」形式があれば C/D に分割し日付化）
    NormalizeAllCD_015 ws

CleanupManualBandFormatting_020

' 帯CF だけ
SetupPeriodConditionalFormatting_020
DrawCategorySeparators_030 ws

    Application.ScreenUpdating = True
    MsgBox "カレンダーの作成と期間描画を完了しました。", vbInformation
End Sub

' ===== ボタン作成（6個、動作割当なし） =====
Public Sub MakeTopbarButtons6_015(ByVal ws As Worksheet)
    Dim addrs As Variant, captions As Variant
    Dim fills As Variant, lines As Variant, i As Long

    addrs = Array("A1", "B1", "A2", "B2", "A3", "B3")
    captions = Array("A1", "B1", "A2", "B2", "A3", "B3")

    fills = Array( _
        RGB(0, 0, 0), _
        RGB(102, 0, 153), _
        RGB(0, 153, 204), _
        RGB(76, 175, 80), _
        RGB(255, 159, 64), _
        RGB(255, 99, 132))

    lines = Array( _
        RGB(0, 0, 0), _
        RGB(70, 0, 105), _
        RGB(0, 102, 153), _
        RGB(56, 142, 60), _
        RGB(204, 108, 28), _
        RGB(179, 60, 92))

    For i = LBound(addrs) To UBound(addrs)
        CreateTopbarLikeButton_015 ws, CStr(addrs(i)), _
            "btn_" & CStr(addrs(i)), CStr(captions(i)), _
            CLng(fills(i)), CLng(lines(i))
    Next i
End Sub

Private Sub CreateTopbarLikeButton_015( _
    ByVal ws As Worksheet, _
    ByVal cellAddress As String, _
    ByVal shapeName As String, _
    ByVal caption As String, _
    ByVal fillColor As Long, _
    ByVal lineColor As Long)

    Dim rng As Range: Set rng = ws.Range(cellAddress)

    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0

    Dim m As Single: m = 2
    Dim W As Single, H As Single
    W = rng.Width - (m * 2): If W < 4 Then W = 4
    H = rng.Height - (m * 2): If H < 4 Then H = 4

    Dim sh As Shape
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, rng.Left + m, rng.Top + m, W, H)

    With sh
        .name = shapeName
        .Placement = xlMoveAndSize
        .LockAspectRatio = msoFalse
        .FILL.ForeColor.RGB = fillColor
        .line.ForeColor.RGB = lineColor
        With .TextFrame
            .Characters.Text = caption
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
            .marginTop = 1.5: .MarginBottom = 1.5
            .marginLeft = 2: .MarginRight = 2
            .AutoSize = False
            With .Characters.Font
                .Size = 9
                .Color = RGB(255, 255, 255)
                .Bold = True
            End With
        End With
        ' OnAction は割り当てません（必要に応じて設定してください）
    End With
End Sub

' ===== ペイン固定（A～D列 + 1～5行） =====
Public Sub Freeze_AB_Rows1to5_015()
    If ActiveWindow Is Nothing Then Exit Sub
    With ActiveWindow
        .FreezePanes = False
.SplitColumn = 5   ' E列で分割（左にA:E）…開始/終了（D/E）を左側に含める
        .SplitRow = 5      ' ← 上5行も固定
        .FreezePanes = True
    End With
End Sub

Public Sub Unfreeze_Panes_015()
    If ActiveWindow Is Nothing Then Exit Sub
    ActiveWindow.FreezePanes = False
End Sub

' ===== カレンダー生成（E4から横に MONTH_SPAN か月） =====
Public Sub MakeLongCalendar_015( _
    ByVal ws As Worksheet, _
    ByVal startCol As Long, _
    ByVal dateRow As Long, _
    ByVal weekRow As Long, _
    ByVal monthSpan As Long)

    Dim d0 As Date, d1 As Date, n As Long
    d0 = DateSerial(Year(Date), Month(Date), 1)              ' 当月1日
    d1 = DateSerial(Year(Date), Month(Date) + monthSpan, 0)  ' monthSpanか月後の月末
    n = d1 - d0 + 1

    ' 既存の2行ぶんをクリア（必要範囲）
    Dim lastCol As Long: lastCol = startCol + n - 1
    ws.Range(ws.Cells(dateRow, startCol), ws.Cells(weekRow, lastCol)).Clear

    ' 月ごとの薄色パレット（ループします）
    Dim bands As Variant
    bands = Array(RGB(235, 248, 255), RGB(255, 244, 230), RGB(238, 240, 255), _
                  RGB(234, 255, 241), RGB(255, 240, 246), RGB(245, 255, 250))

    Dim i As Long, dt As Date, c As Long
    For i = 0 To n - 1
        dt = d0 + i
        c = startCol + i

        ' 上段：実日付（表示は「日」だけ）
        With ws.Cells(dateRow, c)
            .Value = dt
            .NumberFormatLocal = "d"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        ' 下段：曜日（1文字）
        With ws.Cells(weekRow, c)
            .Value = WeekdayJP_015(dt)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlVAlignCenter
        End With

        ' 月帯の薄色
        Dim moIdx As Long
        moIdx = ((Year(dt) - Year(d0)) * 12 + Month(dt) - Month(d0)) Mod (UBound(bands) + 1)
        ws.Range(ws.Cells(dateRow, c), ws.Cells(weekRow, c)).Interior.Color = bands(moIdx)

        ' 週末の文字色
        ' 祝日・週末の文字色（祝日を最優先）
Dim wd As VbDayOfWeek: wd = Weekday(dt, vbSunday)
Dim hName As String, isHol As Boolean
isHol = IsHolidayJP_042(ws, dt, hName)

If isHol Then
    ws.Cells(dateRow, c).Font.Color = RGB(255, 0, 0) ' 祝日は赤
    ws.Cells(weekRow, c).Font.Color = RGB(255, 0, 0)
ElseIf wd = vbSunday Then
    ws.Cells(dateRow, c).Font.Color = RGB(255, 0, 0) ' 日曜も赤
    ws.Cells(weekRow, c).Font.Color = RGB(255, 0, 0)
ElseIf wd = vbSaturday Then
    ws.Cells(dateRow, c).Font.Color = RGB(0, 0, 255) ' 土曜は青
    ws.Cells(weekRow, c).Font.Color = RGB(0, 0, 255)
Else
    ws.Cells(dateRow, c).Font.Color = RGB(0, 0, 0)   ' 平日は黒
    ws.Cells(weekRow, c).Font.Color = RGB(0, 0, 0)
End If

    Next i

    ' 行の高さを少し整える（任意）
    ws.Rows(dateRow & ":" & weekRow).RowHeight = 18
End Sub

Private Function WeekdayJP_015(ByVal d As Date) As String
    Select Case Weekday(d, vbSunday)
        Case vbSunday:    WeekdayJP_015 = "日"
        Case vbMonday:    WeekdayJP_015 = "月"
        Case vbTuesday:   WeekdayJP_015 = "火"
        Case vbWednesday: WeekdayJP_015 = "水"
        Case vbThursday:  WeekdayJP_015 = "木"
        Case vbFriday:    WeekdayJP_015 = "金"
        Case vbSaturday:  WeekdayJP_015 = "土"
    End Select
End Function

' ===== B4 を “選択列の年月” に更新 =====
Public Sub UpdateMonthIndicator_B4_015(ByVal ws As Worksheet, ByVal startCol As Long)
    Dim col As Long, v As Variant, lastCol As Long
    lastCol = ws.Cells(CAL_DATE_ROW_015, ws.Columns.Count).End(xlToLeft).Column
    col = startCol
    If TypeName(Selection) = "Range" Then
        col = Selection.Column
        If col < startCol Or col > lastCol Then col = startCol
    End If
    v = ws.Cells(CAL_DATE_ROW_015, col).Value
    If IsDate(v) Then
        ws.Range("B4").Value = Format$(CDate(v), "yyyy年m月")
    Else
        ws.Range("B4").Value = ""
    End If
End Sub

'=== DEPRECATED: 手塗りは廃止。条件付き書式に委譲する ===
Public Sub ColorizeRowByPeriod_015( _
    ByVal ws As Worksheet, ByVal rowIndex As Long, _
    ByVal calStartCol As Long, ByVal calDateRow As Long, _
    ByVal periodColor As Long, Optional ByVal addGrid As Boolean = True)

    ' 帯は再塗りしない。必要なら帯を一旦クリアして CF に任せるのみ。
    ClearCalendarRow_016 ws, rowIndex, calStartCol, calDateRow
End Sub


' 対象セルに黒い格子（枠線）を引く。weightで太さを指定（既定: 細線）
Public Sub SetCellGrid_016(ByVal tgt As Range, _
                           Optional ByVal lineColor As Long = vbBlack, _
                           Optional ByVal weight As XlBorderWeight = xlThin)
    With tgt.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous: .weight = weight: .Color = lineColor
    End With
    With tgt.Borders(xlEdgeTop)
        .LineStyle = xlContinuous: .weight = weight: .Color = lineColor
    End With
    With tgt.Borders(xlEdgeRight)
        .LineStyle = xlContinuous: .weight = weight: .Color = lineColor
    End With
    With tgt.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .weight = weight: .Color = lineColor
    End With
End Sub

' ===== C/D全体を正規化（Cに「開始～終了」があればC/D分割＆日付化） =====
Public Sub NormalizeAllCD_015(ByVal ws As Worksheet)
    ' ※名前は互換のまま。中身を D/E 版に差し替え（Cはメモ専用）
    Dim lastRowD As Long, lastRowE As Long, lastRow As Long
    lastRowD = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    lastRowE = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    lastRow = Application.Max(lastRowD, lastRowE, DATA_START_ROW_015)

    Dim R As Long
    For R = DATA_START_ROW_015 To lastRow
        ' 旧データ救済：C が日付/期間 かつ D/E が空なら D/E に移し、C は空にする
        NormalizeOrSplit_C_to_CD_015 ws, R   ' ←中身も D/E 版に直します（次項）

        ' D/E を日付セルとして正規化
        Dim dStart As Date, dEnd As Date
        Dim hasD As Boolean, hasE As Boolean
        hasD = TryReadDateFromCell_015(ws, R, "D", dStart)
        hasE = TryReadDateFromCell_015(ws, R, "E", dEnd)

        If hasD And Not hasE Then
            ws.Cells(R, "E").Value = dStart
            ws.Cells(R, "E").NumberFormatLocal = DATE_FMT_015
        ElseIf hasE And Not hasD Then
            ws.Cells(R, "D").Value = dEnd
            ws.Cells(R, "D").NumberFormatLocal = DATE_FMT_015
        End If

        If IsDate(ws.Cells(R, "D").Value) And IsDate(ws.Cells(R, "E").Value) Then
            If CDate(ws.Cells(R, "E").Value) < CDate(ws.Cells(R, "D").Value) Then
                Dim T As Variant
                T = ws.Cells(R, "D").Value
                ws.Cells(R, "D").Value = ws.Cells(R, "E").Value
                ws.Cells(R, "E").Value = T
                ws.Cells(R, "D").NumberFormatLocal = DATE_FMT_015
                ws.Cells(R, "E").NumberFormatLocal = DATE_FMT_015
            End If
        End If
    Next R
End Sub


' ===== C/D を見て全行を再描画 =====
Public Sub RepaintAllPeriods_015(ByVal ws As Worksheet)

End Sub

' ====== 下位ユーティリティ ======

' Cセルが「開始～終了」表記でDが空ならC/Dへ分割して日付化
Private Sub NormalizeOrSplit_C_to_CD_015(ByVal ws As Worksheet, ByVal rowIndex As Long)
    Dim s As String: s = Trim$(CStr(ws.Cells(rowIndex, "C").Value))
    If Len(s) = 0 Then Exit Sub
    If Len(Trim$(CStr(ws.Cells(rowIndex, "D").Value))) > 0 Then Exit Sub
    Dim T As String: T = NormalizeSeparators_015(s)
    If InStr(T, "~") = 0 Then Exit Sub

    Dim a As String, B As String
    a = Split(T, "~")(0)
    B = Split(T, "~")(1)

    Dim d0 As Date, d1 As Date
    If ParseDateLooseJP_015(a, d0) And ParseDateLooseJP_015(B, d1) Then
        ws.Cells(rowIndex, "C").Value = d0
        ws.Cells(rowIndex, "D").Value = d1
        ws.Cells(rowIndex, "C").NumberFormatLocal = DATE_FMT_015
        ws.Cells(rowIndex, "D").NumberFormatLocal = DATE_FMT_015
      Dim tdt As Date
If d1 < d0 Then
    tdt = d0: d0 = d1: d1 = tdt
    ws.Cells(rowIndex, "C").Value = d0
    ws.Cells(rowIndex, "D").Value = d1
End If

    End If
End Sub

' C/Dから開始・終了日を取得（C/Dどちらか欠けていればFalse）
Private Function TryGetCDatesForRow_015(ByVal ws As Worksheet, ByVal rowIndex As Long, _
                                        ByRef d0 As Date, ByRef d1 As Date) As Boolean
    Dim okC As Boolean, okD As Boolean
    okC = TryReadDateFromCell_015(ws, rowIndex, "C", d0)
    okD = TryReadDateFromCell_015(ws, rowIndex, "D", d1)
    If Not okC And Not okD Then
        TryGetCDatesForRow_015 = False
    ElseIf okC And Not okD Then
        d1 = d0
        TryGetCDatesForRow_015 = True
    ElseIf Not okC And okD Then
        d0 = d1
        TryGetCDatesForRow_015 = True
    Else
        TryGetCDatesForRow_015 = True
    End If
End Function

' セルの値を日付に読み替えて書式も付与（成功時 True）
Private Function TryReadDateFromCell_015(ByVal ws As Worksheet, ByVal rowIndex As Long, _
                                         ByVal colLetter As String, ByRef dt As Date) As Boolean
    Dim v As Variant
    v = ws.Cells(rowIndex, colLetter).Value
    If IsDate(v) Then
        dt = CDate(v)
        ws.Cells(rowIndex, colLetter).NumberFormatLocal = DATE_FMT_015
        TryReadDateFromCell_015 = True
        Exit Function
    End If
    Dim s As String: s = Trim$(CStr(v))
    If Len(s) = 0 Then TryReadDateFromCell_015 = False: Exit Function
    If ParseDateLooseJP_015(s, dt) Then
        ws.Cells(rowIndex, colLetter).Value = dt
        ws.Cells(rowIndex, colLetter).NumberFormatLocal = DATE_FMT_015
        TryReadDateFromCell_015 = True
    Else
        TryReadDateFromCell_015 = False
    End If
End Function

' ゆるめの和式→日付パース（yyyy年m月d日 / yyyy/m/d / yyyy-m-d 等）
Private Function ParseDateLooseJP_015(ByVal s As String, ByRef d As Date) As Boolean
    On Error GoTo Fail
    Dim T As String: T = CStr(s)
    ' 全角→半角
    T = StrConv(T, vbNarrow)
    ' 区切りを "/" に寄せる
    T = Replace(T, "年", "/")
    T = Replace(T, "月", "/")
    T = Replace(T, "日", "")
    T = Replace(T, ".", "/")
    T = Replace(T, "-", "/")
    T = Replace(T, "－", "/")
    T = Replace(T, "ー", "/")
    T = Replace(T, "―", "/")
    T = Replace(T, "\", "/")
    T = Replace(T, "／", "/")
    T = Replace(T, " ", "")
    ' トークン分解して DateSerial を優先
    Dim p As Variant: p = Split(T, "/")
    If UBound(p) = 2 Then
        d = DateSerial(CLng(p(0)), CLng(p(1)), CLng(p(2)))
        ParseDateLooseJP_015 = True
        Exit Function
    End If
    ' それ以外は CDate に委譲
    d = CDate(T)
    ParseDateLooseJP_015 = True
    Exit Function
Fail:
    ParseDateLooseJP_015 = False
End Function

' 全角/半角の波線・ダッシュ・ハイフン等を全部 "~" に統一＋年/月/日→"/"
Private Function NormalizeSeparators_015(ByVal s As String) As String
    Dim T As String: T = CStr(s)
    Dim k As Variant
    For Each k In Array(ChrW(&HFF5E), ChrW(&H301C), ChrW(&H2015), ChrW(&H2014), ChrW(&H2013), "-", "ｰ", "－", " to ", " TO ", "~", "～")
        T = Replace(T, CStr(k), "~")
    Next k
    T = StrConv(T, vbNarrow)
    T = Replace(T, "年", "/")
    T = Replace(T, "月", "/")
    T = Replace(T, "日", "")
    T = Replace(T, "\", "/")
    T = Replace(T, ".", "/")
    T = Replace(T, " ", "")
    NormalizeSeparators_015 = T
End Function

'=== Bが空になった行を全消去（Aの優先、C～Dの値/色、横の期間帯の塗り/格子）===
'=== Bが空になった行を全消去（Aの優先、C～Dの値/色、横の期間帯の塗り/格子）===
Public Sub WipeRowOnEmptyB_016(ByVal ws As Worksheet, ByVal rowIndex As Long)
    ' A(優先)とC/D(開始/終了)の値を消す。Bはユーザーが空にしている前提。
    ws.Cells(rowIndex, "A").ClearContents
    ws.Cells(rowIndex, "C").ClearContents
    ws.Cells(rowIndex, "D").ClearContents
    ' B～Dの塗りを消す（A列はもともと塗らない方針）
    ws.Range(ws.Cells(rowIndex, "B"), ws.Cells(rowIndex, "D")).Interior.Pattern = xlNone
    ' カレンダー側（横の帯）の塗りと格子を全消去
    ClearCalendarRow_016 ws, rowIndex, CAL_START_COL_015, CAL_DATE_ROW_015

    ' ★ この行に残っている水平線（上・下）をA～最右列まで明示的に除去
    Dim rightCol As Long
    rightCol = LastCalendarCol_016(ws)
    If rightCol < 4 Then rightCol = 4 ' 最低でもA～Dまでは対象
    With ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, rightCol))
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
End Sub


'=== カレンダー側（行）をクリア（塗りも枠も消す）===
Public Sub ClearCalendarRow_016(ByVal ws As Worksheet, ByVal rowIndex As Long, _
                                ByVal calStartCol As Long, ByVal calDateRow As Long)
    Dim lastCalCol As Long
    lastCalCol = ws.Cells(calDateRow, ws.Columns.Count).End(xlToLeft).Column
    If lastCalCol < calStartCol Then Exit Sub
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(rowIndex, calStartCol), ws.Cells(rowIndex, lastCalCol))
    rng.Interior.Pattern = xlNone
    rng.Borders.LineStyle = xlNone
End Sub

'=== 行の見た目を「セット前」に戻す
'   ・B～D の塗り：なし
'   ・カレンダー帯：C/D に日付があれば 既定の淡黄色＋黒格子／なければ無し
Public Sub WipeRowColors_016(ByVal ws As Worksheet, ByVal rowIndex As Long)
    ws.Range(ws.Cells(rowIndex, "B"), ws.Cells(rowIndex, "D")).Interior.Pattern = xlNone

End Sub

'=== 親のAが消えたら、その下に連続する同じ優先番号（=子ブロック）を全部「セット前」に戻す
Public Sub RemoveChildrenIfParentCleared_016(ByVal ws As Worksheet, ByVal parentRow As Long)
    Dim pNext As Variant, pPrev As Variant
    pNext = ws.Cells(parentRow + 1, "A").Value
    If Not IsNumeric(pNext) Then Exit Sub

    pPrev = ws.Cells(parentRow - 1, "A").Value
    If IsNumeric(pPrev) Then
        If CLng(pPrev) = CLng(pNext) Then Exit Sub
    End If

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim R As Long
    Application.EnableEvents = False
    For R = parentRow + 1 To lastRow
        If IsNumeric(ws.Cells(R, "A").Value) Then
            If CLng(ws.Cells(R, "A").Value) = CLng(pNext) Then
                ws.Cells(R, "A").ClearContents
                WipeRowColors_016 ws, R
            Else
                Exit For
            End If
        Else
            Exit For
        End If
    Next R
    Application.EnableEvents = True
End Sub

'=== 旧API名は残しつつ、安全な二段ソートに委譲 ===
Public Sub AddRowIntoPrioritySet_016(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal p As Long)
    If ws Is Nothing Then Set ws = ActiveSheet
    If rowIndex < DATA_START_ROW_015 Then Exit Sub
    If p < 1 Then Exit Sub

    Application.EnableEvents = False
    ws.Cells(rowIndex, "A").Value = CLng(p) ' その行の分類を確定
    ws.Cells(rowIndex, "A").NumberFormat = "0"
    Application.EnableEvents = True

    ' 全体を A→D で並べ替え
    ResortByCategory_031 ws
End Sub


'=== C/Dの「終了日」を取得。Dが日付→D、D不正→C、どちらも不正→9999/12/31 ===
Public Function GetRangeEndDate_016(ByVal ws As Worksheet, ByVal rowIndex As Long) As Date
    Dim d As Date
    If TryReadDateFromCell_015(ws, rowIndex, "D", d) Then
        GetRangeEndDate_016 = d: Exit Function
    End If
    If TryReadDateFromCell_015(ws, rowIndex, "C", d) Then
        GetRangeEndDate_016 = d: Exit Function
    End If
    GetRangeEndDate_016 = DateSerial(9999, 12, 31)
End Function

'=== 子の行配列を終了日 昇順で並べる ===
Public Sub SortRowsByDate_Ascending_016(ByRef rowsArr() As Long, ByRef endsArr() As Date)
    Dim i As Long, j As Long, n As Long
    n = UBound(rowsArr)
    For i = 1 To n - 1
        For j = i + 1 To n
            If endsArr(j) < endsArr(i) Then
                SwapL_016 rowsArr(i), rowsArr(j)
                SwapD_016 endsArr(i), endsArr(j)
            End If
        Next j
    Next i
End Sub

Private Sub SwapL_016(ByRef a As Long, ByRef B As Long)
    Dim T As Long: T = a: a = B: B = T
End Sub
Private Sub SwapD_016(ByRef a As Date, ByRef B As Date)
    Dim T As Date: T = a: a = B: B = T
End Sub

'=== データの最終行（B列基準） ===
Public Function LastDataRow_016(ByVal ws As Worksheet) As Long
    Dim R As Long
    R = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If R < DATA_START_ROW_015 Then R = DATA_START_ROW_015
    LastDataRow_016 = R
End Function

'=== カレンダーの最終列（4行目の右端） ===
Public Function LastCalendarCol_016(ByVal ws As Worksheet) As Long
    LastCalendarCol_016 = ws.Cells(CAL_DATE_ROW_015, ws.Columns.Count).End(xlToLeft).Column
End Function

' ===== 互換維持：旧API（文字列）→C/Dへ書き込みに置き換え =====
'   ・"開始～終了" 文字列：C/Dへ分割して保存
'   ・単一日文字列：C/Dとも同じ日に保存
'   ・その後に行の色塗りを実行
Public Sub WritePeriodToC_015(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal rawText As String)
    Dim T As String: T = NormalizeSeparators_015(rawText)
    Dim d0 As Date, d1 As Date
    If InStr(T, "~") > 0 Then
        If ParseDateLooseJP_015(Split(T, "~")(0), d0) And ParseDateLooseJP_015(Split(T, "~")(1), d1) Then
            Dim x As Date
If d1 < d0 Then
    x = d0: d0 = d1: d1 = x
End If

            Application.EnableEvents = False
            ws.Cells(rowIndex, "C").Value = d0: ws.Cells(rowIndex, "C").NumberFormatLocal = DATE_FMT_015
            ws.Cells(rowIndex, "D").Value = d1: ws.Cells(rowIndex, "D").NumberFormatLocal = DATE_FMT_015
            Application.EnableEvents = True
            EnsureCategoryToRow_017 ws, rowIndex

            Exit Sub
        End If
    Else
        If ParseDateLooseJP_015(T, d0) Then
            Application.EnableEvents = False
            ws.Cells(rowIndex, "C").Value = d0: ws.Cells(rowIndex, "C").NumberFormatLocal = DATE_FMT_015
            ws.Cells(rowIndex, "D").Value = d0: ws.Cells(rowIndex, "D").NumberFormatLocal = DATE_FMT_015
            Application.EnableEvents = True
            EnsureCategoryToRow_017 ws, rowIndex
            Exit Sub
        End If
    End If
    ' パース不可：生のテキストをCに置く（Dは空のまま）
    Application.EnableEvents = False
    ws.Cells(rowIndex, "C").Value = rawText
    Application.EnableEvents = True
    EnsureCategoryToRow_017 ws, rowIndex

  
End Sub

' ==== 互換用：旧APIを復活（年補完つきの文字列正規化） =====================

Public Function NormalizeDateRangeWithYearFromCalendar_015( _
    ByVal raw As String, ByVal ws As Worksheet, _
    ByVal startCol As Long, ByVal dateRow As Long) As String

    Dim T As String: T = NormalizeSeparators_015(raw)
    If InStr(T, "~") = 0 Then NormalizeDateRangeWithYearFromCalendar_015 = "": Exit Function

    Dim a As String, B As String
    a = Split(T, "~")(0)
    B = Split(T, "~")(1)

    Dim y1 As Long, m1 As Long, d1 As Long
    Dim y2 As Long, m2 As Long, d2 As Long
    If Not ParseMDY_015(a, y1, m1, d1) Then Exit Function
    If Not ParseMDY_015(B, y2, m2, d2) Then Exit Function

    ' 年が無い側はカレンダー（4行目）から補完
    If y1 = 0 Then y1 = FindYearInCalendar_015(ws, startCol, dateRow, m1, d1)
    If y2 = 0 Then y2 = FindYearInCalendar_015(ws, startCol, dateRow, m2, d2)

    On Error GoTo Fail
    Dim dt1 As Date, dt2 As Date
    dt1 = DateSerial(y1, m1, d1)
    dt2 = DateSerial(y2, m2, d2)
    NormalizeDateRangeWithYearFromCalendar_015 = _
        Format$(dt1, DATE_FMT_015) & "～" & Format$(dt2, DATE_FMT_015)
    Exit Function
Fail:
    NormalizeDateRangeWithYearFromCalendar_015 = ""
End Function

' 年あり/なしの m/d(/y) を分解（年なしは y=0）
Private Function ParseMDY_015(ByVal s As String, ByRef y As Long, ByRef m As Long, ByRef d As Long) As Boolean
    Dim u As String: u = NormalizeSeparators_015(CStr(s))
    Dim p As Variant: p = Split(u, "/")
    On Error GoTo Fail
    Select Case UBound(p)
        Case 2 ' y/m/d
            y = CLng(p(0)): m = CLng(p(1)): d = CLng(p(2))
        Case 1 ' m/d（年なし）
            y = 0: m = CLng(p(0)): d = CLng(p(1))
        Case Else
            Dim dt As Date
            dt = CDate(u)
            y = Year(dt): m = Month(dt): d = Day(dt)
    End Select
    ParseMDY_015 = True
    Exit Function
Fail:
    ParseMDY_015 = False
End Function

' カレンダー4行目から該当 m/d の年を見つける（見つからなければ今年）
Private Function FindYearInCalendar_015(ByVal ws As Worksheet, ByVal startCol As Long, _
                                        ByVal dateRow As Long, ByVal m As Long, ByVal d As Long) As Long
    Dim lastCol As Long: lastCol = ws.Cells(dateRow, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long, v As Variant
    For c = startCol To lastCol
        v = ws.Cells(dateRow, c).Value
        If IsDate(v) Then
            If Month(CDate(v)) = m And Day(CDate(v)) = d Then
                FindYearInCalendar_015 = Year(CDate(v))
                Exit Function
            End If
        End If
    Next c
    FindYearInCalendar_015 = Year(Date)
End Function

' 互換：旧の区切り解析（必要にしておく）。"yyyy/m/d~yyyy/m/d" を日付に。
Public Function ParseDateRange_015(ByVal s As String, ByRef d0 As Date, ByRef d1 As Date) As Boolean
    Dim T As String: T = NormalizeSeparators_015(s)
    Dim parts As Variant: parts = Split(T, "~")
    If UBound(parts) <> 1 Then Exit Function
    On Error GoTo Fail
    d0 = CDate(parts(0))
    d1 = CDate(parts(1))
    ParseDateRange_015 = True
    Exit Function
Fail:
    ParseDateRange_015 = False
End Function

' 互換：旧名で呼ばれても新処理に委譲
Public Sub NormalizeAllC_015(ByVal ws As Worksheet)
    NormalizeAllCD_015 ws
End Sub




' 選択中の分類番号を取得（優先順位：名前付き範囲「選択分類」→固定セル）
Private Function TryReadSelectedCategory_017(ByVal ws As Worksheet, ByRef p As Long) As Boolean
    Dim v As Variant

    ' 1) 名前付き範囲「選択分類」があれば最優先
    On Error Resume Next
    v = ws.Parent.Names("選択分類").RefersToRange.Value
    If Err.Number = 0 Then
        If IsNumeric(v) Then p = CLng(v): TryReadSelectedCategory_017 = True: On Error GoTo 0: Exit Function
    End If
    Err.Clear
    On Error GoTo 0

    ' 2) 予備：固定セル（CAT_SELECTED_ADDR_017）
    On Error Resume Next
    v = ws.Range(CAT_SELECTED_ADDR_017).Value
    On Error GoTo 0
    If IsNumeric(v) Then
        p = CLng(v)
        TryReadSelectedCategory_017 = True
    Else
        TryReadSelectedCategory_017 = False
    End If
End Function

' A列が空なら「選択中の分類番号」を書き込み、並べ替え規則に組み入れる
Public Sub EnsureCategoryToRow_017(ByVal ws As Worksheet, ByVal rowIndex As Long)
    Dim p As Long
    If Len(Trim$(CStr(ws.Cells(rowIndex, "A").Value))) > 0 Then Exit Sub
    If TryReadSelectedCategory_017(ws, p) Then
        ws.Cells(rowIndex, "A").Value = p
        ws.Cells(rowIndex, "A").NumberFormat = "0"
        ' 既存の並べ替えロジックへ組み込み（A昇順＋親子処理）
        AddRowIntoPrioritySet_016 ws, rowIndex, p
    End If
End Sub
'=== 次の書き込み行（C列ベース） ===
Public Function NextFreeRowCD_017(ByVal ws As Worksheet) As Long
    Dim R As Long
    R = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    If R < DATA_START_ROW_015 Then
        NextFreeRowCD_017 = DATA_START_ROW_015
    ElseIf Len(Trim$(CStr(ws.Cells(R, "C").Value))) = 0 And _
           Len(Trim$(CStr(ws.Cells(R, "D").Value))) = 0 And _
           Len(Trim$(CStr(ws.Cells(R, "B").Value))) = 0 And _
           Len(Trim$(CStr(ws.Cells(R, "A").Value))) = 0 Then
        NextFreeRowCD_017 = R
    Else
        NextFreeRowCD_017 = R + 1
    End If
End Function

'=== A列へ分類番号を書き込む（直接指定）。doSort=True で既存ソートへ組み込み ===
Public Sub SetCategoryToRow_017(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal cat As Long, _
                                Optional ByVal doSort As Boolean = True)
    If cat <= 0 Then Exit Sub
    ws.Cells(rowIndex, "A").Value = CLng(cat)
    ws.Cells(rowIndex, "A").NumberFormat = "0"
    If doSort Then ResortByCategory_031 ws

End Sub

'=== 期間バーの条件付き書式（分類色＋四辺の罫線補完） ===
Public Sub SetupPeriodConditionalFormatting_020()

    Dim ws As Worksheet: Set ws = ActiveSheet

    Const MAX_CAT As Long = 10
    Const FALLBACK As Long = &HA3F8FF       ' BGR: 255,248,163（予備色）

    Dim calStartCol As Long: calStartCol = CAL_START_COL_015   ' 例: 5
    Dim calDateRow  As Long: calDateRow = CAL_DATE_ROW_015     ' 例: 4
    Dim dataStartRow As Long: dataStartRow = DATA_START_ROW_015 ' 例: 6

    Dim lastCalCol As Long
    lastCalCol = ws.Cells(calDateRow, ws.Columns.Count).End(xlToLeft).Column
    If lastCalCol < calStartCol Then Exit Sub

    Dim lastRow As Long
    lastRow = Application.Max( _
                ws.Cells(ws.Rows.Count, "B").End(xlUp).Row, _
                ws.Cells(ws.Rows.Count, "C").End(xlUp).Row, _
                ws.Cells(ws.Rows.Count, "D").End(xlUp).Row, _
                dataStartRow)

    Dim rngApply As Range
    Set rngApply = ws.Range(ws.Cells(dataStartRow, calStartCol), ws.Cells(lastRow, lastCalCol))

    ' 参照形式は A1 固定
    Dim prevRef As XlReferenceStyle
    prevRef = Application.ReferenceStyle
    Application.ReferenceStyle = xlA1

    Dim sep As String: sep = Application.International(xlListSeparator)

    ' A1相対アドレス（左上セル = 行相対）
Dim addrHdr As String, addrA As String, addrStart As String, addrEnd As String
addrHdr = ws.Cells(calDateRow, calStartCol).Address(True, False, xlA1)   ' F$4（定数変更に追随）
addrA = ws.Cells(dataStartRow, 1).Address(False, True, xlA1)             ' $A6
addrStart = ws.Cells(dataStartRow, 4).Address(False, True, xlA1)         ' $D6
addrEnd = ws.Cells(dataStartRow, 5).Address(False, True, xlA1)           ' $E6

 ' 期間判定式
Dim inRangeCond As String
inRangeCond = addrStart & "<>""" & """" & sep & _
              addrHdr & ">=" & addrStart & sep & _
              addrHdr & "<=IF(" & addrEnd & "=""" & """" & sep & addrStart & sep & addrEnd & ")"

    Application.ScreenUpdating = False
    On Error GoTo Finally

    ' 既存CF全削除
    rngApply.FormatConditions.Delete

    '---------------------------
    ' 1) 横線をCFで補完するルール
    '    上線: (ROW()=先頭) or (同カテゴリが上にも続く)
    '    下線: (ROW()=最終) or (同カテゴリが下にも続く)
    '    ※ 切替行の「上線」はCFで出さない（→通常罫線の二重線が見える）
    '---------------------------
    Dim fTop As String, fBottom As String
    fTop = "=AND(" & inRangeCond & sep & "OR(ROW()=" & CStr(dataStartRow) & sep & _
           "$A" & CStr(dataStartRow) & "=$A" & CStr(dataStartRow - 1) & "))"

    fBottom = "=AND(" & inRangeCond & sep & "OR(ROW()=" & CStr(lastRow) & sep & _
              "$A" & CStr(dataStartRow) & "=$A" & CStr(dataStartRow + 1) & "))"

    Dim fC As FormatCondition

    ' 上線（細線）
    Set fC = rngApply.FormatConditions.Add(Type:=xlExpression, Formula1:=fTop)
    With fC
        .StopIfTrue = False
        With .Borders(xlTop)
            .LineStyle = xlContinuous
            .weight = xlThin
            .Color = vbBlack
        End With
    End With

    ' 下線（細線）
    Set fC = rngApply.FormatConditions.Add(Type:=xlExpression, Formula1:=fBottom)
    With fC
        .StopIfTrue = False
        With .Borders(xlBottom)
            .LineStyle = xlContinuous
            .weight = xlThin
            .Color = vbBlack
        End With
    End With

    '---------------------------
    ' 2) 分類ごとの「塗り＋左右線」（左/右のみ）
    '---------------------------
    ' 分類の色（B列の塗りが見つかれば優先）
    Dim catColor(1 To MAX_CAT) As Long
    Dim catHas(1 To MAX_CAT) As Boolean

    Dim R As Long, k As Long
    For R = dataStartRow To lastRow
        If IsNumeric(ws.Cells(R, "A").Value) Then
            k = CLng(ws.Cells(R, "A").Value)
            If k >= 1 And k <= MAX_CAT Then
                If Not catHas(k) Then
                    With ws.Cells(R, "B").Interior
                        If .Pattern <> xlNone Then
                            catColor(k) = .Color
                            catHas(k) = True
                        End If
                    End With
                End If
            End If
        End If
    Next R

    ' 既定色（VIVID 相当。BGR）
' 既定色（B列の塗りが無いカテゴリ用）
Dim defCol(1 To MAX_CAT) As Long
defCol(1) = RGB(220, 53, 69)      ' 1 赤
defCol(2) = RGB(0, 153, 204)      ' 2 ライトブルー
defCol(3) = RGB(255, 165, 0)      ' 3 オレンジ
defCol(4) = RGB(76, 175, 80)      ' 4 緑
defCol(5) = RGB(255, 99, 132)     ' 5 ピンク
defCol(6) = RGB(102, 0, 153)      ' 6 紫
defCol(7) = RGB(0, 51, 102)       ' 7 紺
defCol(8) = RGB(146, 208, 80)     ' 8 黄緑
defCol(9) = RGB(255, 215, 0)      ' 9 黄
defCol(10) = RGB(150, 75, 0)    ' ★茶色 (#964B00)


    For k = 1 To MAX_CAT
        If Not catHas(k) Then catColor(k) = defCol(k)
        If catColor(k) = 0 Then catColor(k) = FALLBACK
    Next k

' 親子で色を分ける（親=catColor、子=Lighten_020(catColor, 0.4)）
Dim f As String, fParent As String, fChild As String  ' ← f は後ろの「未分類」でも使うため一緒に宣言
Dim childCol As Long

For k = 1 To MAX_CAT
    ' === 親：A=k かつ 先頭行（= 上のAと違う行、または最初の行） ===
    fParent = "=AND(" & addrA & "=" & CStr(k) & sep & inRangeCond & sep & "OR(" & _
              "ROW()=" & CStr(dataStartRow) & sep & _
              "OFFSET(" & addrA & ",-1,0)<>" & addrA & "))"
    Set fC = rngApply.FormatConditions.Add(Type:=xlExpression, Formula1:=fParent)
    With fC
        .StopIfTrue = False
        .Interior.Pattern = xlSolid
        .Interior.Color = catColor(k)   ' 親＝基準色
        With .Borders(xlLeft):  .LineStyle = xlContinuous: .weight = xlThin: .Color = vbBlack: End With
        With .Borders(xlRight): .LineStyle = xlContinuous: .weight = xlThin: .Color = vbBlack: End With
    End With

    ' === 子：A=k かつ 上のAと同じ（ブロック2行目以降） ===
    childCol = Lighten_020(catColor(k), 0.7)  ' ← 薄さは好みで 0.25～0.6 などに調整
    fChild = "=AND(" & addrA & "=" & CStr(k) & sep & inRangeCond & sep & _
             "ROW()>" & CStr(dataStartRow) & sep & "OFFSET(" & addrA & ",-1,0)=" & addrA & ")"
    Set fC = rngApply.FormatConditions.Add(Type:=xlExpression, Formula1:=fChild)
    With fC
        .StopIfTrue = False
        .Interior.Pattern = xlSolid
        .Interior.Color = childCol      ' 子＝薄色
        With .Borders(xlLeft):  .LineStyle = xlContinuous: .weight = xlThin: .Color = vbBlack: End With
        With .Borders(xlRight): .LineStyle = xlContinuous: .weight = xlThin: .Color = vbBlack: End With
    End With
Next k


    ' 未分類など（予備色）
    f = "=AND((" & addrA & "=""" & """" & ")+(" & addrA & "<1)+(" & addrA & ">10)>0" & sep & inRangeCond & ")"
    Set fC = rngApply.FormatConditions.Add(Type:=xlExpression, Formula1:=f)
    With fC
        .StopIfTrue = False
        .Interior.Pattern = xlSolid
        .Interior.Color = FALLBACK
        With .Borders(xlLeft):  .LineStyle = xlContinuous: .weight = xlThin: .Color = vbBlack: End With
        With .Borders(xlRight): .LineStyle = xlContinuous: .weight = xlThin: .Color = vbBlack: End With
    End With

Finally:
    Application.ReferenceStyle = prevRef
    Application.ScreenUpdating = True
End Sub








' 既存の手塗り痕跡（塗り/罫線）を帯全体から除去（最初の一回だけ実行推奨）
Public Sub CleanupManualBandFormatting_020()
Dim ws As Worksheet
Set ws = ActiveSheet  ' ← Sheet4 固定をやめる
    Dim calStartCol As Long: calStartCol = CAL_START_COL_015
    Dim calDateRow As Long: calDateRow = CAL_DATE_ROW_015
    Dim dataStartRow As Long: dataStartRow = DATA_START_ROW_015

    Dim lastCalCol As Long
    lastCalCol = ws.Cells(calDateRow, ws.Columns.Count).End(xlToLeft).Column
    Dim lastRow As Long
lastRow = Application.Max( _
            ws.Cells(ws.Rows.Count, "B").End(xlUp).Row, _
            ws.Cells(ws.Rows.Count, "D").End(xlUp).Row, _
            ws.Cells(ws.Rows.Count, "E").End(xlUp).Row, _
            dataStartRow)

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(dataStartRow, calStartCol), ws.Cells(lastRow, lastCalCol))
    rng.Interior.Pattern = xlNone
    rng.Borders.LineStyle = xlNone
End Sub


'=== 分類の切替位置に二重線を引く（A列～カレンダー最終列まで） ==========
Public Sub ApplyCategoryDividers_020(ByVal ws As Worksheet)

    If ws Is Nothing Then Set ws = ActiveSheet

    Dim firstRow As Long: firstRow = DATA_START_ROW_015
    Dim lastRow  As Long: lastRow = LastDataRow_016(ws)
    If lastRow < firstRow Then Exit Sub

    ' カレンダーの一番右の列まで線を伸ばす
    Dim rightCol As Long: rightCol = LastCalendarCol_016(ws)   ' ← E以降の最終列

    Application.ScreenUpdating = False

    ' 既存の水平線を一旦クリア（A～最終列）
    With ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, rightCol))
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With

    ' A列の番号が変わる行の「上」に二重線を入れる
    Dim R As Long, prevCat As Variant, curCat As Variant
    prevCat = ws.Cells(firstRow, "A").Value

    For R = firstRow + 1 To lastRow
        curCat = ws.Cells(R, "A").Value
        If IsNumeric(prevCat) And IsNumeric(curCat) Then
            If CLng(prevCat) <> CLng(curCat) Then
                With ws.Range(ws.Cells(R, 1), ws.Cells(R, rightCol)).Borders(xlEdgeTop)
                    .LineStyle = xlDouble
                    .Color = vbBlack
                End With
            End If
        End If
        prevCat = curCat
    Next R

    Application.ScreenUpdating = True
End Sub

Public Sub ResortByCategory_031(ByVal ws As Worksheet)
    ResortAllByCategory_032 ws
End Sub



Public Sub ResortAllByCategory_032(ByVal ws As Worksheet)
    If ws Is Nothing Then Set ws = ActiveSheet

    Dim firstRow As Long: firstRow = DATA_START_ROW_015
    Dim lastRow  As Long: lastRow = LastDataRow_016(ws)
    If lastRow < firstRow Then Exit Sub

    Dim lastCol As Long: lastCol = LastCalendarCol_016(ws)
    If lastCol < 4 Then lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim ev As Boolean: ev = Application.EnableEvents
    Dim su As Boolean: su = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' 1) 子の並びを確定（A → D → E）
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range(ws.Cells(firstRow, "A"), ws.Cells(lastRow, "A")), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add key:=ws.Range(ws.Cells(firstRow, "D"), ws.Cells(lastRow, "D")), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add key:=ws.Range(ws.Cells(firstRow, "E"), ws.Cells(lastRow, "E")), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, lastCol))
        .Header = xlNo
        .Apply
    End With

    ' 2) 各Aブロックで「最長(E-D)」の行を親(先頭)へ昇格（インライン実装）
    Dim R As Long: R = firstRow
    Do While R <= lastRow
        If IsNumeric(ws.Cells(R, "A").Value) Then
            Dim rStart As Long: rStart = R
            Dim rEnd   As Long: rEnd = R
            ' 同じAが連続する最終行を探す
            Do While rEnd + 1 <= lastRow _
                And IsNumeric(ws.Cells(rEnd + 1, "A").Value) _
                And CLng(ws.Cells(rEnd + 1, "A").Value) = CLng(ws.Cells(rStart, "A").Value)
                rEnd = rEnd + 1
            Loop

            ' 最長を選ぶ（同長は 開始が早い→終了が早い を優先）
            Dim rMax As Long: rMax = rStart
            Dim maxDur As Double: maxDur = -1
            Dim rr As Long
            For rr = rStart To rEnd
                Dim s As Variant, e As Variant
                s = ws.Cells(rr, "D").Value
                e = ws.Cells(rr, "E").Value

                Dim dur As Double
                If IsDate(s) And IsDate(e) Then
                    dur = CDbl(CDate(e)) - CDbl(CDate(s))
                    If dur < 0 Then dur = 0
                Else
                    dur = 0
                End If

                Dim curS As Double: curS = IIf(IsDate(s), CDbl(CDate(s)), 0#)
                Dim curE As Double: curE = IIf(IsDate(e), CDbl(CDate(e)), 0#)
                Dim maxS As Double: maxS = IIf(IsDate(ws.Cells(rMax, "D").Value), CDbl(CDate(ws.Cells(rMax, "D").Value)), 0#)
                Dim maxE As Double: maxE = IIf(IsDate(ws.Cells(rMax, "E").Value), CDbl(CDate(ws.Cells(rMax, "E").Value)), 0#)

                If dur > maxDur _
                   Or (dur = maxDur And curS < maxS) _
                   Or (dur = maxDur And curS = maxS And curE < maxE) Then
                    maxDur = dur
                    rMax = rr
                End If
            Next rr

            ' 親としてブロック先頭へ移動（Cut→Insert）
            If rMax > rStart Then
                ws.Rows(rMax).Cut
                ws.Rows(rStart).Insert Shift:=xlDown
                Application.CutCopyMode = False
                rEnd = rEnd + 1   ' 挿入でインデックスがずれる分を補正
            End If

            R = rEnd + 1
        Else
            R = R + 1
        End If
    Loop

    ' 3) 見た目の再適用（色→帯CF→区切り線）
    On Error Resume Next

    On Error GoTo 0
    SetupPeriodConditionalFormatting_020
    ApplyCategoryDividers_020 ws

    Application.ScreenUpdating = su
    Application.EnableEvents = ev
End Sub



'=== 「今日より前」の日付列を“列まるごと”灰色にする ===
'   ・カレンダー開始列(CAL_START_COL_015)～最終日付列(LastCalendarCol_016)
'   ・行は dateRow(=4行目) から、B列の最終データ行まで
'   ・既存の月帯色や期間バー(CF)は消さない（上から塗るだけ）
Public Sub GreyOutPastDateColumns_041(ByVal ws As Worksheet, _
                                      Optional ByVal fromRow As Long = 0, _
                                      Optional ByVal grayBGR As Long = &HEEEEEE)
    If ws Is Nothing Then Set ws = ActiveSheet

    Dim calStartCol As Long: calStartCol = CAL_START_COL_015
    Dim dateRow     As Long: dateRow = CAL_DATE_ROW_015
    Dim weekRow     As Long: weekRow = CAL_WEEK_ROW_015

    Dim topRow As Long
    ' 既定は 4行目から塗る（丸ごと=4行目～データ最終行）
    topRow = IIf(fromRow > 0, fromRow, dateRow)

    Dim lastCalCol As Long: lastCalCol = LastCalendarCol_016(ws)

    Dim lastRow As Long
    ' データ最終行と曜日行のどちらか長い方まで塗る
    lastRow = Application.Max(LastDataRow_016(ws), weekRow)
    If lastRow < topRow Then lastRow = topRow

    Dim c As Long, v As Variant
    Dim tgt As Range, seg As Range

    Application.ScreenUpdating = False
    ' ★クリアはしない（既存の配色は残す）

    For c = calStartCol To lastCalCol
        v = ws.Cells(dateRow, c).Value
        If IsDate(v) Then
            If CDate(v) < Date Then
                Set seg = ws.Range(ws.Cells(topRow, c), ws.Cells(lastRow, c))
                If tgt Is Nothing Then
                    Set tgt = seg
                Else
                    Set tgt = Union(tgt, seg)
                End If
            End If
        End If
    Next c

    If Not tgt Is Nothing Then
        With tgt.Interior
            .Pattern = xlSolid
            .Color = grayBGR   ' 例: &HEEEEEE = #EEEEEE（BGR）
            .TintAndShade = 0
        End With
    End If

    Application.ScreenUpdating = True
End Sub





ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー
順番　標準モジュール　(使えないしいらない)


Option Explicit

'============================== エントリーポイント ==============================
Public Sub 順()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' --- 親イベントの前提チェック ---
    If Selection Is Nothing Then
        MsgBox "B列の親イベント（1セル）を選んでから実行してください。", vbExclamation
        Exit Sub
    End If
    If Selection.Cells.CountLarge <> 1 Or Selection.Column <> 2 Then
        MsgBox "B列の親イベント（1セル）を選んでから実行してください。", vbExclamation
        Exit Sub
    End If

    Dim parentRow As Long
    parentRow = Selection.Row
    If parentRow < DATA_START_ROW_015 Then
        MsgBox "データは " & DATA_START_ROW_015 & " 行目以降を想定しています。", vbExclamation
        Exit Sub
    End If
    If Len(Trim$(ws.Cells(parentRow, "B").Value)) = 0 Then
        MsgBox "親イベントのBセルが空です。", vbExclamation
        Exit Sub
    End If

    ' --- 優先順位（分類番号）を取得 ---
    Dim prio As Long
    prio = AskPriorityNumber_016()
    If prio = 0 Then Exit Sub

    ' --- 子イベントを複数選択（B列）。キャンセル可 ---
    Dim rngPick As Range, rngB As Range
    On Error Resume Next
    Set rngPick = Application.InputBox( _
        Prompt:="同じセットに入れる子イベント（B列の複数セル）を選択してOK。不要ならキャンセル。", _
        Title:="子イベントの選択", Type:=8)
    On Error GoTo 0
    If rngPick Is Nothing Then
        Set rngB = Nothing
    Else
        Set rngB = Intersect(rngPick, ws.Columns("B"))
    End If

    ' --- 子イベントの行番号を集める（親は除外／重複排除） ---
    Dim childRows() As Long, cnt As Long
    If Not rngB Is Nothing Then
        ReDim childRows(1 To rngB.Cells.Count)
        Dim c As Range
        For Each c In rngB.Cells
            If c.Row >= DATA_START_ROW_015 Then
                If c.Row <> parentRow Then
                    If Len(Trim$(ws.Cells(c.Row, "B").Value)) > 0 Then
                        cnt = cnt + 1
                        childRows(cnt) = c.Row
                    End If
                End If
            End If
        Next c
        If cnt > 0 Then
            ReDim Preserve childRows(1 To cnt)
            RemoveDupRowsInPlace_016 childRows
        End If
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' --- 0) 念のため A列を整数化（過去の小数痕跡をリセット） ---
    NormalizePriorityColumn_031 ws

    ' --- 1) 優先順位の挿入（既存の prio 以上を +1 に繰り下げ） ---
    ShiftPrioritiesDownFrom_016 ws, prio

    ' --- 2) 親・子に新しい優先順位を設定（A=prio） ---
    ws.Cells(parentRow, "A").Value = prio
    Dim i As Long
    If cnt > 0 Then
        For i = 1 To UBound(childRows)
            ws.Cells(childRows(i), "A").Value = prio
        Next i
    End If

    ' --- 3) 親を最上位にするための一時ソートキー列を用意 ---
    Dim firstRow As Long: firstRow = DATA_START_ROW_015
    Dim lastRow As Long:  lastRow = LastDataRow_016(ws)

    Dim tmpCol As Long
    tmpCol = LastCalendarCol_016(ws) + 1     ' カレンダーの右隣を一時列として使用

    ' いったん 0 で初期化
    ws.Range(ws.Cells(firstRow, tmpCol), ws.Cells(lastRow, tmpCol)).Value = 0

    ' 親=1、同じ優先の子=2 をマーク（他は 0）
    ws.Cells(parentRow, tmpCol).Value = 1
    If cnt > 0 Then
        For i = 1 To UBound(childRows)
            ws.Cells(childRows(i), tmpCol).Value = 2
        Next i
    End If

    ' --- 4) A→一時キー→D（終了日）の三段ソート ---
    Dim sortRightCol As Long
    sortRightCol = Application.WorksheetFunction.Max(LastCalendarCol_016(ws), tmpCol)

With ws.Sort
    .SortFields.Clear
    .SortFields.Add Key:=ws.Range(ws.Cells(firstRow, "A"), ws.Cells(lastRow, "A")), _
                    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SortFields.Add Key:=ws.Range(ws.Cells(firstRow, tmpCol), ws.Cells(lastRow, tmpCol)), _
                    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ' ここを D → E に
    .SortFields.Add Key:=ws.Range(ws.Cells(firstRow, "E"), ws.Cells(lastRow, "E")), _
                    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SetRange ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, sortRightCol))
    .Header = xlNo
    .Apply
End With

    End With

    ' 一時キー列をクリア
    ws.Range(ws.Cells(firstRow, tmpCol), ws.Cells(lastRow, tmpCol)).ClearContents

    ' --- 5) 見た目の更新（色／帯／区切り線） ---
    ApplyPriorityColors_AllRows_016 ws                 ' B～C の塗り
    SetupPeriodConditionalFormatting_020               ' 期間帯（分類色で塗る）
    ApplyCategoryDividers_020 ws                       ' 分類が切り替わる行に二重線

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "分類セットの作成・並べ替え・色分けを完了しました。", vbInformation
End Sub


'============================== ヘルパー群 ==============================

' 優先順位の入力（1以上の整数）。キャンセル時は 0 を返す。
Private Function AskPriorityNumber_016() As Long
    Dim x As Variant
    x = Application.InputBox(Prompt:="親イベントの優先順位を入力してください（1が最優先の整数）。", Title:="優先順位", Type:=1)
    If x = False Then
        AskPriorityNumber_016 = 0
        Exit Function
    End If
    On Error GoTo Bad
    If CLng(x) < 1 Then GoTo Bad
    AskPriorityNumber_016 = CLng(x)
    Exit Function
Bad:
    MsgBox "1以上の整数を入力してください。", vbExclamation
    AskPriorityNumber_016 = 0
End Function

' 既存の優先順位 >= p を +1（繰り下げ）
Private Sub ShiftPrioritiesDownFrom_016(ByVal ws As Worksheet, ByVal p As Long)
    Dim lastRow As Long: lastRow = LastDataRow_016(ws)
    Dim R As Long, v
    For R = DATA_START_ROW_015 To lastRow
        v = ws.Cells(R, "A").Value
        If IsNumeric(v) Then
            If CLng(v) >= p Then ws.Cells(R, "A").Value = CLng(v) + 1
        End If
    Next R
End Sub

' C列の「終了日」を取得。年欠落はカレンダーから補完して C も年付きへ置換。
Private Function GetRangeEndDate_016(ByVal ws As Worksheet, ByVal rowIndex As Long) As Date
    Dim s As String, d0 As Date, d1 As Date
    s = CStr(ws.Cells(rowIndex, "C").Value)
    If Len(Trim$(s)) = 0 Then
        GetRangeEndDate_016 = DateSerial(9999, 12, 31)
        Exit Function
    End If

    Dim norm As String
    norm = NormalizeDateRangeWithYearFromCalendar_015(s, ws, CAL_START_COL_015, CAL_DATE_ROW_015)
    If Len(norm) > 0 Then
        Application.EnableEvents = False
        ws.Cells(rowIndex, "C").Value = norm
        Application.EnableEvents = True
        If ParseDateRange_015(norm, d0, d1) Then
            GetRangeEndDate_016 = d1
            Exit Function
        End If
    End If

    On Error Resume Next
    d1 = CDate(s)
    On Error GoTo 0
    If d1 > 0 Then
        GetRangeEndDate_016 = d1
    Else
        GetRangeEndDate_016 = DateSerial(9999, 12, 31)
    End If
End Function

' 子の行配列を終了日昇順に並べ替え
Private Sub SortRowsByDate_Ascending_016(ByRef rowsArr() As Long, ByRef endsArr() As Date)
    Dim i As Long, j As Long, n As Long
    n = UBound(rowsArr)
    For i = 1 To n - 1
        For j = i + 1 To n
            If endsArr(j) < endsArr(i) Then
                SwapL_016 rowsArr(i), rowsArr(j)
                SwapD_016 endsArr(i), endsArr(j)
            End If
        Next j
    Next i
End Sub

Private Sub SwapL_016(ByRef a As Long, ByRef B As Long)
    Dim T As Long: T = a: a = B: B = T
End Sub
Private Sub SwapD_016(ByRef a As Date, ByRef B As Date)
    Dim T As Date: T = a: a = B: B = T
End Sub



' データの最終行（B列基準）
Private Function LastDataRow_016(ByVal ws As Worksheet) As Long
    Dim R As Long
    R = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If R < DATA_START_ROW_015 Then R = DATA_START_ROW_015
    LastDataRow_016 = R
End Function

' カレンダーの最終列（4行目の右端）
Private Function LastCalendarCol_016(ByVal ws As Worksheet) As Long
    LastCalendarCol_016 = ws.Cells(CAL_DATE_ROW_015, ws.Columns.Count).End(xlToLeft).Column
End Function

' 優先順位カラーを適用（親＝鮮やか、子＝パステル）／A列は塗らない
Public Sub ApplyPriorityColors_AllRows_016(ByVal ws As Worksheet)
    Dim lastRow As Long: lastRow = LastDataRow_016(ws)
    Dim lastCol As Long: lastCol = LastCalendarCol_016(ws)
    Dim R As Long, p As Long
    Dim prevP As Variant: prevP = Empty

    ' B～Cの塗りだけクリア（Aは触らない）
    ws.Range(ws.Cells(DATA_START_ROW_015, "B"), ws.Cells(lastRow, "C")).Interior.Pattern = xlNone

    For R = DATA_START_ROW_015 To lastRow
        If IsNumeric(ws.Cells(R, "A").Value) Then
            p = CLng(ws.Cells(R, "A").Value)
            If p >= 1 Then
                Dim baseCol As Long, useCol As Long
                baseCol = PriorityColor_016(p)                  ' 子用（従来のパステル）
                If prevP <> p Then                              ' ブロック先頭＝親
                    useCol = SaturateColor_016(baseCol, 1.8, 0.88) ' 親：彩度↑、わずかに濃く
                    prevP = p
                Else
                    useCol = baseCol                            ' 子：そのまま
                End If

                ' B～Cだけ塗る（A列は塗らない）
                ws.Range(ws.Cells(R, "B"), ws.Cells(R, "C")).Interior.Color = useCol

            End If
        Else
            prevP = Empty
        End If
    Next R
End Sub



' 優先順位→色（パステル10色。11以上はループ）
Private Function PriorityColor_016(ByVal p As Long) As Long
    Dim palette As Variant
    palette = Array(RGB(255, 235, 235), RGB(235, 245, 255), RGB(235, 255, 235), RGB(255, 245, 230), RGB(245, 235, 255), _
                    RGB(255, 240, 250), RGB(235, 255, 255), RGB(255, 255, 230), RGB(240, 248, 255), RGB(240, 255, 240))
    Dim idx As Long
    idx = (p - 1) Mod (UBound(palette) + 1)
    PriorityColor_016 = palette(idx)
End Function

' 指定色を暗く（濃く）する。factor=0.75 で 75% の明るさに。
Private Function DarkerColor_016(ByVal col As Long, Optional ByVal factor As Double = 0.75) As Long
    Dim R As Long, G As Long, B As Long
    R = col Mod 256
    G = (col \ 256) Mod 256
    B = (col \ 65536) Mod 256
    R = CLng(R * factor): If R < 0 Then R = 0
    G = CLng(G * factor): If G < 0 Then G = 0
    B = CLng(B * factor): If B < 0 Then B = 0
    DarkerColor_016 = RGB(R, G, B)
End Function

' 指定色を「同じ色相で鮮やかに」する（親用）
' satFactor: 彩度倍率、lightFactor: 明度倍率（<1で少し濃く）
Public Function SaturateColor_016(ByVal col As Long, _
                                  Optional ByVal satFactor As Double = 1.8, _
                                  Optional ByVal lightFactor As Double = 0.9) As Long
    Dim R As Double, G As Double, B As Double
    R = (col Mod 256) / 255#
    G = ((col \ 256) Mod 256) / 255#
    B = ((col \ 65536) Mod 256) / 255#

    Dim H As Double, s As Double, L As Double
    RGB_To_HSL_016 R, G, B, H, s, L

    s = s * satFactor: If s > 1# Then s = 1#
    L = L * lightFactor: If L < 0# Then L = 0#: If L > 1# Then L = 1#

    Dim rr As Double, gg As Double, bb As Double
    HSL_To_RGB_016 H, s, L, rr, gg, bb

    SaturateColor_016 = RGB(CLng(rr * 255#), CLng(gg * 255#), CLng(bb * 255#))
End Function

' --- RGB(0..1) → HSL(0..1)
Private Sub RGB_To_HSL_016(ByVal R As Double, ByVal G As Double, ByVal B As Double, _
                           ByRef H As Double, ByRef s As Double, ByRef L As Double)
    Dim maxV As Double, minv As Double, d As Double
    maxV = R: If G > maxV Then maxV = G: If B > maxV Then maxV = B
    minv = R: If G < minv Then minv = G: If B < minv Then minv = B
    L = (maxV + minv) / 2#
    If maxV = minv Then
        H = 0#: s = 0#
    Else
        d = maxV - minv
        If L < 0.5 Then s = d / (maxV + minv) Else s = d / (2# - maxV - minv)
        If maxV = R Then
            H = (G - B) / d + IIf(G < B, 6#, 0#)
        ElseIf maxV = G Then
            H = (B - R) / d + 2#
        Else
            H = (R - G) / d + 4#
        End If
        H = H / 6#
    End If
End Sub

' --- HSL(0..1) → RGB(0..1)
Private Sub HSL_To_RGB_016(ByVal H As Double, ByVal s As Double, ByVal L As Double, _
                            ByRef R As Double, ByRef G As Double, ByRef B As Double)
    Dim q As Double, p As Double
    If s = 0# Then
        R = L: G = L: B = L
    Else
        If L < 0.5 Then q = L * (1# + s) Else q = L + s - L * s
        p = 2# * L - q
        R = Hue2RGB_016(p, q, H + 1# / 3#)
        G = Hue2RGB_016(p, q, H)
        B = Hue2RGB_016(p, q, H - 1# / 3#)
    End If
End Sub

Private Function Hue2RGB_016(ByVal p As Double, ByVal q As Double, ByVal T As Double) As Double
    If T < 0# Then T = T + 1#
    If T > 1# Then T = T - 1#
    If T < 1# / 6# Then
        Hue2RGB_016 = p + (q - p) * 6# * T
    ElseIf T < 1# / 2# Then
        Hue2RGB_016 = q
    ElseIf T < 2# / 3# Then
        Hue2RGB_016 = p + (q - p) * (2# / 3# - T) * 6#
    Else
        Hue2RGB_016 = p
    End If
End Function


'=== A列を整数に正規化（過去の 2.000001 等を除去）===
Public Sub NormalizePriorityColumn_031(ByVal ws As Worksheet)
    Dim lastRow As Long: lastRow = LastDataRow_016(ws)
    Dim R As Long, v
    For R = DATA_START_ROW_015 To lastRow
        v = ws.Cells(R, "A").Value
        If IsNumeric(v) Then ws.Cells(R, "A").Value = CLng(v)
    Next R
End Sub

'=== 重複行番号を除去（0でマーキング→圧縮）===
Private Sub RemoveDupRowsInPlace_016(ByRef arr() As Long)
    Dim i As Long, j As Long, n As Long
    If (Not Not arr) = 0 Then Exit Sub
    n = UBound(arr)
    For i = 1 To n - 1
        If arr(i) <> 0 Then
            For j = i + 1 To n
                If arr(j) = arr(i) Then arr(j) = 0
            Next j
        End If
    Next i
    Dim tmp() As Long, k As Long
    ReDim tmp(1 To n)
    For i = 1 To n
        If arr(i) <> 0 Then
            k = k + 1
            tmp(k) = arr(i)
        End If
    Next i
    If k = 0 Then
        Erase arr
    Else
        ReDim arr(1 To k)
        For i = 1 To k
            arr(i) = tmp(i)
        Next i
    End If
End Sub








------------------------------------------------------------------------------------------------------------------------------------------------------------------


'=== modJapanHolidays_042 ===標準モジュール
Option Explicit

' 祝日シート（A列=日付、B列=名称）を見て、その日が祝日なら True を返す
' ※ ws は基準シート（ActiveSheet でOK）。同じブック内の「祝日」シートを参照します。
Public Function IsHolidayJP_042(ByVal ws As Worksheet, ByVal d As Date, ByRef nameOut As String) As Boolean
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ws.Parent.Worksheets("祝日")
    On Error GoTo 0
    If sh Is Nothing Then Exit Function ' 祝日シートが無ければ False（＝従来どおり土日だけ色）

    ' 時刻を切り捨てて一致判定（Excel内部は日付=連続小数）
    Dim key As Double: key = CDbl(DateSerial(Year(d), Month(d), Day(d)))
    Dim m As Variant
    m = Application.Match(key, sh.Columns(1), 0) ' A列に一致日付があれば行番号
    If IsError(m) Then
        IsHolidayJP_042 = False
    Else
        nameOut = CStr(sh.Cells(CLng(m), 2).Value) ' B列=祝日名
        IsHolidayJP_042 = True
    End If
End Function

'（任意）内閣府の「国民の祝日」CSVを読み込んで「祝日」シートを作る
' 使い方：ImportCabinetHolidaysCsv_042 "C:\path\syukujitsu.csv"
Public Sub ImportCabinetHolidaysCsv_042(ByVal csvPath As String)
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets("祝日")
    On Error GoTo 0
    If sh Is Nothing Then
        Set sh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        sh.name = "祝日"
    End If

    sh.Cells.Clear
    sh.Range("A1").Value = "日付"
    sh.Range("B1").Value = "名称"

    Dim f As Integer, line As String, arr As Variant, R As Long
    f = FreeFile
    Open csvPath For Input As #f
    R = 1
    Do While Not EOF(f)
        Line Input #f, line
        ' 1行目のヘッダ（"日付","名称" 等）はスキップ
        If InStr(line, "日付") > 0 And InStr(line, "名称") > 0 Then
            ' skip
        Else
            arr = Split(line, ",")
            If UBound(arr) >= 1 Then
                R = R + 1
                On Error Resume Next
                sh.Cells(R, 1).Value = CDate(Replace(arr(0), """", ""))
                sh.Cells(R, 1).NumberFormatLocal = "yyyy/m/d"
                sh.Cells(R, 2).Value = Replace(arr(1), """", "")
                On Error GoTo 0
            End If
        End If
    Loop
    Close #f

    ' 日付昇順に整列（重複は手元のCSVに従う）
    If R > 1 Then
        With sh.Sort
            .SortFields.Clear
            .SortFields.Add key:=sh.Range(sh.Cells(2, 1), sh.Cells(R, 1)), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange sh.Range(sh.Cells(1, 1), sh.Cells(R, 2))
            .Header = xlYes
            .Apply
        End With
    End If

    ' 見えないシートにしておく（必要なら xlSheetVisible に）
    sh.Visible = xlSheetHidden
End Sub




------------------------------------------------------------------------------------------------------------------------------------------------------------------
祝日　　---シート名が祝日で標準モジュール



国民の祝日・休日月日	国民の祝日・休日名称
1955/1/1	元日
1955/1/15	成人の日
1955/3/21	春分の日
1955/4/29	天皇誕生日
1955/5/3	憲法記念日
1955/5/5	こどもの日
1955/9/24	秋分の日
1955/11/3	文化の日
1955/11/23	勤労感謝の日
1956/1/1	元日
1956/1/15	成人の日
1956/3/21	春分の日
1956/4/29	天皇誕生日
1956/5/3	憲法記念日
1956/5/5	こどもの日
1956/9/23	秋分の日
1956/11/3	文化の日
1956/11/23	勤労感謝の日
1957/1/1	元日
1957/1/15	成人の日
1957/3/21	春分の日
1957/4/29	天皇誕生日
1957/5/3	憲法記念日
1957/5/5	こどもの日
1957/9/23	秋分の日
1957/11/3	文化の日
1957/11/23	勤労感謝の日
1958/1/1	元日
1958/1/15	成人の日
1958/3/21	春分の日
1958/4/29	天皇誕生日
1958/5/3	憲法記念日
1958/5/5	こどもの日
1958/9/23	秋分の日
1958/11/3	文化の日
1958/11/23	勤労感謝の日
1959/1/1	元日
1959/1/15	成人の日
1959/3/21	春分の日
1959/4/10	結婚の儀
1959/4/29	天皇誕生日
1959/5/3	憲法記念日
1959/5/5	こどもの日
1959/9/24	秋分の日
1959/11/3	文化の日
1959/11/23	勤労感謝の日
1960/1/1	元日
1960/1/15	成人の日
1960/3/20	春分の日
1960/4/29	天皇誕生日
1960/5/3	憲法記念日
1960/5/5	こどもの日
1960/9/23	秋分の日
1960/11/3	文化の日
1960/11/23	勤労感謝の日
1961/1/1	元日
1961/1/15	成人の日
1961/3/21	春分の日
1961/4/29	天皇誕生日
1961/5/3	憲法記念日
1961/5/5	こどもの日
1961/9/23	秋分の日
1961/11/3	文化の日
1961/11/23	勤労感謝の日
1962/1/1	元日
1962/1/15	成人の日
1962/3/21	春分の日
1962/4/29	天皇誕生日
1962/5/3	憲法記念日
1962/5/5	こどもの日
1962/9/23	秋分の日
1962/11/3	文化の日
1962/11/23	勤労感謝の日
1963/1/1	元日
1963/1/15	成人の日
1963/3/21	春分の日
1963/4/29	天皇誕生日
1963/5/3	憲法記念日
1963/5/5	こどもの日
1963/9/24	秋分の日
1963/11/3	文化の日
1963/11/23	勤労感謝の日
1964/1/1	元日
1964/1/15	成人の日
1964/3/20	春分の日
1964/4/29	天皇誕生日
1964/5/3	憲法記念日
1964/5/5	こどもの日
1964/9/23	秋分の日
1964/11/3	文化の日
1964/11/23	勤労感謝の日
1965/1/1	元日
1965/1/15	成人の日
1965/3/21	春分の日
1965/4/29	天皇誕生日
1965/5/3	憲法記念日
1965/5/5	こどもの日
1965/9/23	秋分の日
1965/11/3	文化の日
1965/11/23	勤労感謝の日
1966/1/1	元日
1966/1/15	成人の日
1966/3/21	春分の日
1966/4/29	天皇誕生日
1966/5/3	憲法記念日
1966/5/5	こどもの日
1966/9/15	敬老の日
1966/9/23	秋分の日
1966/10/10	体育の日
1966/11/3	文化の日
1966/11/23	勤労感謝の日
1967/1/1	元日
1967/1/15	成人の日
1967/2/11	建国記念の日
1967/3/21	春分の日
1967/4/29	天皇誕生日
1967/5/3	憲法記念日
1967/5/5	こどもの日
1967/9/15	敬老の日
1967/9/24	秋分の日
1967/10/10	体育の日
1967/11/3	文化の日
1967/11/23	勤労感謝の日
1968/1/1	元日
1968/1/15	成人の日
1968/2/11	建国記念の日
1968/3/20	春分の日
1968/4/29	天皇誕生日
1968/5/3	憲法記念日
1968/5/5	こどもの日
1968/9/15	敬老の日
1968/9/23	秋分の日
1968/10/10	体育の日
1968/11/3	文化の日
1968/11/23	勤労感謝の日
1969/1/1	元日
1969/1/15	成人の日
1969/2/11	建国記念の日
1969/3/21	春分の日
1969/4/29	天皇誕生日
1969/5/3	憲法記念日
1969/5/5	こどもの日
1969/9/15	敬老の日
1969/9/23	秋分の日
1969/10/10	体育の日
1969/11/3	文化の日
1969/11/23	勤労感謝の日
1970/1/1	元日
1970/1/15	成人の日
1970/2/11	建国記念の日
1970/3/21	春分の日
1970/4/29	天皇誕生日
1970/5/3	憲法記念日
1970/5/5	こどもの日
1970/9/15	敬老の日
1970/9/23	秋分の日
1970/10/10	体育の日
1970/11/3	文化の日
1970/11/23	勤労感謝の日
1971/1/1	元日
1971/1/15	成人の日
1971/2/11	建国記念の日
1971/3/21	春分の日
1971/4/29	天皇誕生日
1971/5/3	憲法記念日
1971/5/5	こどもの日
1971/9/15	敬老の日
1971/9/24	秋分の日
1971/10/10	体育の日
1971/11/3	文化の日
1971/11/23	勤労感謝の日
1972/1/1	元日
1972/1/15	成人の日
1972/2/11	建国記念の日
1972/3/20	春分の日
1972/4/29	天皇誕生日
1972/5/3	憲法記念日
1972/5/5	こどもの日
1972/9/15	敬老の日
1972/9/23	秋分の日
1972/10/10	体育の日
1972/11/3	文化の日
1972/11/23	勤労感謝の日
1973/1/1	元日
1973/1/15	成人の日
1973/2/11	建国記念の日
1973/3/21	春分の日
1973/4/29	天皇誕生日
1973/4/30	休日
1973/5/3	憲法記念日
1973/5/5	こどもの日
1973/9/15	敬老の日
1973/9/23	秋分の日
1973/9/24	休日
1973/10/10	体育の日
1973/11/3	文化の日
1973/11/23	勤労感謝の日
1974/1/1	元日
1974/1/15	成人の日
1974/2/11	建国記念の日
1974/3/21	春分の日
1974/4/29	天皇誕生日
1974/5/3	憲法記念日
1974/5/5	こどもの日
1974/5/6	休日
1974/9/15	敬老の日
1974/9/16	休日
1974/9/23	秋分の日
1974/10/10	体育の日
1974/11/3	文化の日
1974/11/4	休日
1974/11/23	勤労感謝の日
1975/1/1	元日
1975/1/15	成人の日
1975/2/11	建国記念の日
1975/3/21	春分の日
1975/4/29	天皇誕生日
1975/5/3	憲法記念日
1975/5/5	こどもの日
1975/9/15	敬老の日
1975/9/24	秋分の日
1975/10/10	体育の日
1975/11/3	文化の日
1975/11/23	勤労感謝の日
1975/11/24	休日
1976/1/1	元日
1976/1/15	成人の日
1976/2/11	建国記念の日
1976/3/20	春分の日
1976/4/29	天皇誕生日
1976/5/3	憲法記念日
1976/5/5	こどもの日
1976/9/15	敬老の日
1976/9/23	秋分の日
1976/10/10	体育の日
1976/10/11	休日
1976/11/3	文化の日
1976/11/23	勤労感謝の日
1977/1/1	元日
1977/1/15	成人の日
1977/2/11	建国記念の日
1977/3/21	春分の日
1977/4/29	天皇誕生日
1977/5/3	憲法記念日
1977/5/5	こどもの日
1977/9/15	敬老の日
1977/9/23	秋分の日
1977/10/10	体育の日
1977/11/3	文化の日
1977/11/23	勤労感謝の日
1978/1/1	元日
1978/1/2	休日
1978/1/15	成人の日
1978/1/16	休日
1978/2/11	建国記念の日
1978/3/21	春分の日
1978/4/29	天皇誕生日
1978/5/3	憲法記念日
1978/5/5	こどもの日
1978/9/15	敬老の日
1978/9/23	秋分の日
1978/10/10	体育の日
1978/11/3	文化の日
1978/11/23	勤労感謝の日
1979/1/1	元日
1979/1/15	成人の日
1979/2/11	建国記念の日
1979/2/12	休日
1979/3/21	春分の日
1979/4/29	天皇誕生日
1979/4/30	休日
1979/5/3	憲法記念日
1979/5/5	こどもの日
1979/9/15	敬老の日
1979/9/24	秋分の日
1979/10/10	体育の日
1979/11/3	文化の日
1979/11/23	勤労感謝の日
1980/1/1	元日
1980/1/15	成人の日
1980/2/11	建国記念の日
1980/3/20	春分の日
1980/4/29	天皇誕生日
1980/5/3	憲法記念日
1980/5/5	こどもの日
1980/9/15	敬老の日
1980/9/23	秋分の日
1980/10/10	体育の日
1980/11/3	文化の日
1980/11/23	勤労感謝の日
1980/11/24	休日
1981/1/1	元日
1981/1/15	成人の日
1981/2/11	建国記念の日
1981/3/21	春分の日
1981/4/29	天皇誕生日
1981/5/3	憲法記念日
1981/5/4	休日
1981/5/5	こどもの日
1981/9/15	敬老の日
1981/9/23	秋分の日
1981/10/10	体育の日
1981/11/3	文化の日
1981/11/23	勤労感謝の日
1982/1/1	元日
1982/1/15	成人の日
1982/2/11	建国記念の日
1982/3/21	春分の日
1982/3/22	休日
1982/4/29	天皇誕生日
1982/5/3	憲法記念日
1982/5/5	こどもの日
1982/9/15	敬老の日
1982/9/23	秋分の日
1982/10/10	体育の日
1982/10/11	休日
1982/11/3	文化の日
1982/11/23	勤労感謝の日
1983/1/1	元日
1983/1/15	成人の日
1983/2/11	建国記念の日
1983/3/21	春分の日
1983/4/29	天皇誕生日
1983/5/3	憲法記念日
1983/5/5	こどもの日
1983/9/15	敬老の日
1983/9/23	秋分の日
1983/10/10	体育の日
1983/11/3	文化の日
1983/11/23	勤労感謝の日
1984/1/1	元日
1984/1/2	休日
1984/1/15	成人の日
1984/1/16	休日
1984/2/11	建国記念の日
1984/3/20	春分の日
1984/4/29	天皇誕生日
1984/4/30	休日
1984/5/3	憲法記念日
1984/5/5	こどもの日
1984/9/15	敬老の日
1984/9/23	秋分の日
1984/9/24	休日
1984/10/10	体育の日
1984/11/3	文化の日
1984/11/23	勤労感謝の日
1985/1/1	元日
1985/1/15	成人の日
1985/2/11	建国記念の日
1985/3/21	春分の日
1985/4/29	天皇誕生日
1985/5/3	憲法記念日
1985/5/5	こどもの日
1985/5/6	休日
1985/9/15	敬老の日
1985/9/16	休日
1985/9/23	秋分の日
1985/10/10	体育の日
1985/11/3	文化の日
1985/11/4	休日
1985/11/23	勤労感謝の日
1986/1/1	元日
1986/1/15	成人の日
1986/2/11	建国記念の日
1986/3/21	春分の日
1986/4/29	天皇誕生日
1986/5/3	憲法記念日
1986/5/5	こどもの日
1986/9/15	敬老の日
1986/9/23	秋分の日
1986/10/10	体育の日
1986/11/3	文化の日
1986/11/23	勤労感謝の日
1986/11/24	休日
1987/1/1	元日
1987/1/15	成人の日
1987/2/11	建国記念の日
1987/3/21	春分の日
1987/4/29	天皇誕生日
1987/5/3	憲法記念日
1987/5/4	休日
1987/5/5	こどもの日
1987/9/15	敬老の日
1987/9/23	秋分の日
1987/10/10	体育の日
1987/11/3	文化の日
1987/11/23	勤労感謝の日
1988/1/1	元日
1988/1/15	成人の日
1988/2/11	建国記念の日
1988/3/20	春分の日
1988/3/21	休日
1988/4/29	天皇誕生日
1988/5/3	憲法記念日
1988/5/4	休日
1988/5/5	こどもの日
1988/9/15	敬老の日
1988/9/23	秋分の日
1988/10/10	体育の日
1988/11/3	文化の日
1988/11/23	勤労感謝の日
1989/1/1	元日
1989/1/2	休日
1989/1/15	成人の日
1989/1/16	休日
1989/2/11	建国記念の日
1989/2/24	大喪の礼
1989/3/21	春分の日
1989/4/29	みどりの日
1989/5/3	憲法記念日
1989/5/4	休日
1989/5/5	こどもの日
1989/9/15	敬老の日
1989/9/23	秋分の日
1989/10/10	体育の日
1989/11/3	文化の日
1989/11/23	勤労感謝の日
1989/12/23	天皇誕生日
1990/1/1	元日
1990/1/15	成人の日
1990/2/11	建国記念の日
1990/2/12	休日
1990/3/21	春分の日
1990/4/29	みどりの日
1990/4/30	休日
1990/5/3	憲法記念日
1990/5/4	休日
1990/5/5	こどもの日
1990/9/15	敬老の日
1990/9/23	秋分の日
1990/9/24	休日
1990/10/10	体育の日
1990/11/3	文化の日
1990/11/12	即位礼正殿の儀
1990/11/23	勤労感謝の日
1990/12/23	天皇誕生日
1990/12/24	休日
1991/1/1	元日
1991/1/15	成人の日
1991/2/11	建国記念の日
1991/3/21	春分の日
1991/4/29	みどりの日
1991/5/3	憲法記念日
1991/5/4	休日
1991/5/5	こどもの日
1991/5/6	休日
1991/9/15	敬老の日
1991/9/16	休日
1991/9/23	秋分の日
1991/10/10	体育の日
1991/11/3	文化の日
1991/11/4	休日
1991/11/23	勤労感謝の日
1991/12/23	天皇誕生日
1992/1/1	元日
1992/1/15	成人の日
1992/2/11	建国記念の日
1992/3/20	春分の日
1992/4/29	みどりの日
1992/5/3	憲法記念日
1992/5/4	休日
1992/5/5	こどもの日
1992/9/15	敬老の日
1992/9/23	秋分の日
1992/10/10	体育の日
1992/11/3	文化の日
1992/11/23	勤労感謝の日
1992/12/23	天皇誕生日
1993/1/1	元日
1993/1/15	成人の日
1993/2/11	建国記念の日
1993/3/20	春分の日
1993/4/29	みどりの日
1993/5/3	憲法記念日
1993/5/4	休日
1993/5/5	こどもの日
1993/6/9	結婚の儀
1993/9/15	敬老の日
1993/9/23	秋分の日
1993/10/10	体育の日
1993/10/11	休日
1993/11/3	文化の日
1993/11/23	勤労感謝の日
1993/12/23	天皇誕生日
1994/1/1	元日
1994/1/15	成人の日
1994/2/11	建国記念の日
1994/3/21	春分の日
1994/4/29	みどりの日
1994/5/3	憲法記念日
1994/5/4	休日
1994/5/5	こどもの日
1994/9/15	敬老の日
1994/9/23	秋分の日
1994/10/10	体育の日
1994/11/3	文化の日
1994/11/23	勤労感謝の日
1994/12/23	天皇誕生日
1995/1/1	元日
1995/1/2	休日
1995/1/15	成人の日
1995/1/16	休日
1995/2/11	建国記念の日
1995/3/21	春分の日
1995/4/29	みどりの日
1995/5/3	憲法記念日
1995/5/4	休日
1995/5/5	こどもの日
1995/9/15	敬老の日
1995/9/23	秋分の日
1995/10/10	体育の日
1995/11/3	文化の日
1995/11/23	勤労感謝の日
1995/12/23	天皇誕生日
1996/1/1	元日
1996/1/15	成人の日
1996/2/11	建国記念の日
1996/2/12	休日
1996/3/20	春分の日
1996/4/29	みどりの日
1996/5/3	憲法記念日
1996/5/4	休日
1996/5/5	こどもの日
1996/5/6	休日
1996/7/20	海の日
1996/9/15	敬老の日
1996/9/16	休日
1996/9/23	秋分の日
1996/10/10	体育の日
1996/11/3	文化の日
1996/11/4	休日
1996/11/23	勤労感謝の日
1996/12/23	天皇誕生日
1997/1/1	元日
1997/1/15	成人の日
1997/2/11	建国記念の日
1997/3/20	春分の日
1997/4/29	みどりの日
1997/5/3	憲法記念日
1997/5/5	こどもの日
1997/7/20	海の日
1997/7/21	休日
1997/9/15	敬老の日
1997/9/23	秋分の日
1997/10/10	体育の日
1997/11/3	文化の日
1997/11/23	勤労感謝の日
1997/11/24	休日
1997/12/23	天皇誕生日
1998/1/1	元日
1998/1/15	成人の日
1998/2/11	建国記念の日
1998/3/21	春分の日
1998/4/29	みどりの日
1998/5/3	憲法記念日
1998/5/4	休日
1998/5/5	こどもの日
1998/7/20	海の日
1998/9/15	敬老の日
1998/9/23	秋分の日
1998/10/10	体育の日
1998/11/3	文化の日
1998/11/23	勤労感謝の日
1998/12/23	天皇誕生日
1999/1/1	元日
1999/1/15	成人の日
1999/2/11	建国記念の日
1999/3/21	春分の日
1999/3/22	休日
1999/4/29	みどりの日
1999/5/3	憲法記念日
1999/5/4	休日
1999/5/5	こどもの日
1999/7/20	海の日
1999/9/15	敬老の日
1999/9/23	秋分の日
1999/10/10	体育の日
1999/10/11	休日
1999/11/3	文化の日
1999/11/23	勤労感謝の日
1999/12/23	天皇誕生日
2000/1/1	元日
2000/1/10	成人の日
2000/2/11	建国記念の日
2000/3/20	春分の日
2000/4/29	みどりの日
2000/5/3	憲法記念日
2000/5/4	休日
2000/5/5	こどもの日
2000/7/20	海の日
2000/9/15	敬老の日
2000/9/23	秋分の日
2000/10/9	体育の日
2000/11/3	文化の日
2000/11/23	勤労感謝の日
2000/12/23	天皇誕生日
2001/1/1	元日
2001/1/8	成人の日
2001/2/11	建国記念の日
2001/2/12	休日
2001/3/20	春分の日
2001/4/29	みどりの日
2001/4/30	休日
2001/5/3	憲法記念日
2001/5/4	休日
2001/5/5	こどもの日
2001/7/20	海の日
2001/9/15	敬老の日
2001/9/23	秋分の日
2001/9/24	休日
2001/10/8	体育の日
2001/11/3	文化の日
2001/11/23	勤労感謝の日
2001/12/23	天皇誕生日
2001/12/24	休日
2002/1/1	元日
2002/1/14	成人の日
2002/2/11	建国記念の日
2002/3/21	春分の日
2002/4/29	みどりの日
2002/5/3	憲法記念日
2002/5/4	休日
2002/5/5	こどもの日
2002/5/6	休日
2002/7/20	海の日
2002/9/15	敬老の日
2002/9/16	休日
2002/9/23	秋分の日
2002/10/14	体育の日
2002/11/3	文化の日
2002/11/4	休日
2002/11/23	勤労感謝の日
2002/12/23	天皇誕生日
2003/1/1	元日
2003/1/13	成人の日
2003/2/11	建国記念の日
2003/3/21	春分の日
2003/4/29	みどりの日
2003/5/3	憲法記念日
2003/5/5	こどもの日
2003/7/21	海の日
2003/9/15	敬老の日
2003/9/23	秋分の日
2003/10/13	体育の日
2003/11/3	文化の日
2003/11/23	勤労感謝の日
2003/11/24	休日
2003/12/23	天皇誕生日
2004/1/1	元日
2004/1/12	成人の日
2004/2/11	建国記念の日
2004/3/20	春分の日
2004/4/29	みどりの日
2004/5/3	憲法記念日
2004/5/4	休日
2004/5/5	こどもの日
2004/7/19	海の日
2004/9/20	敬老の日
2004/9/23	秋分の日
2004/10/11	体育の日
2004/11/3	文化の日
2004/11/23	勤労感謝の日
2004/12/23	天皇誕生日
2005/1/1	元日
2005/1/10	成人の日
2005/2/11	建国記念の日
2005/3/20	春分の日
2005/3/21	休日
2005/4/29	みどりの日
2005/5/3	憲法記念日
2005/5/4	休日
2005/5/5	こどもの日
2005/7/18	海の日
2005/9/19	敬老の日
2005/9/23	秋分の日
2005/10/10	体育の日
2005/11/3	文化の日
2005/11/23	勤労感謝の日
2005/12/23	天皇誕生日
2006/1/1	元日
2006/1/2	休日
2006/1/9	成人の日
2006/2/11	建国記念の日
2006/3/21	春分の日
2006/4/29	みどりの日
2006/5/3	憲法記念日
2006/5/4	休日
2006/5/5	こどもの日
2006/7/17	海の日
2006/9/18	敬老の日
2006/9/23	秋分の日
2006/10/9	体育の日
2006/11/3	文化の日
2006/11/23	勤労感謝の日
2006/12/23	天皇誕生日
2007/1/1	元日
2007/1/8	成人の日
2007/2/11	建国記念の日
2007/2/12	休日
2007/3/21	春分の日
2007/4/29	昭和の日
2007/4/30	休日
2007/5/3	憲法記念日
2007/5/4	みどりの日
2007/5/5	こどもの日
2007/7/16	海の日
2007/9/17	敬老の日
2007/9/23	秋分の日
2007/9/24	休日
2007/10/8	体育の日
2007/11/3	文化の日
2007/11/23	勤労感謝の日
2007/12/23	天皇誕生日
2007/12/24	休日
2008/1/1	元日
2008/1/14	成人の日
2008/2/11	建国記念の日
2008/3/20	春分の日
2008/4/29	昭和の日
2008/5/3	憲法記念日
2008/5/4	みどりの日
2008/5/5	こどもの日
2008/5/6	休日
2008/7/21	海の日
2008/9/15	敬老の日
2008/9/23	秋分の日
2008/10/13	体育の日
2008/11/3	文化の日
2008/11/23	勤労感謝の日
2008/11/24	休日
2008/12/23	天皇誕生日
2009/1/1	元日
2009/1/12	成人の日
2009/2/11	建国記念の日
2009/3/20	春分の日
2009/4/29	昭和の日
2009/5/3	憲法記念日
2009/5/4	みどりの日
2009/5/5	こどもの日
2009/5/6	休日
2009/7/20	海の日
2009/9/21	敬老の日
2009/9/22	休日
2009/9/23	秋分の日
2009/10/12	体育の日
2009/11/3	文化の日
2009/11/23	勤労感謝の日
2009/12/23	天皇誕生日
2010/1/1	元日
2010/1/11	成人の日
2010/2/11	建国記念の日
2010/3/21	春分の日
2010/3/22	休日
2010/4/29	昭和の日
2010/5/3	憲法記念日
2010/5/4	みどりの日
2010/5/5	こどもの日
2010/7/19	海の日
2010/9/20	敬老の日
2010/9/23	秋分の日
2010/10/11	体育の日
2010/11/3	文化の日
2010/11/23	勤労感謝の日
2010/12/23	天皇誕生日
2011/1/1	元日
2011/1/10	成人の日
2011/2/11	建国記念の日
2011/3/21	春分の日
2011/4/29	昭和の日
2011/5/3	憲法記念日
2011/5/4	みどりの日
2011/5/5	こどもの日
2011/7/18	海の日
2011/9/19	敬老の日
2011/9/23	秋分の日
2011/10/10	体育の日
2011/11/3	文化の日
2011/11/23	勤労感謝の日
2011/12/23	天皇誕生日
2012/1/1	元日
2012/1/2	休日
2012/1/9	成人の日
2012/2/11	建国記念の日
2012/3/20	春分の日
2012/4/29	昭和の日
2012/4/30	休日
2012/5/3	憲法記念日
2012/5/4	みどりの日
2012/5/5	こどもの日
2012/7/16	海の日
2012/9/17	敬老の日
2012/9/22	秋分の日
2012/10/8	体育の日
2012/11/3	文化の日
2012/11/23	勤労感謝の日
2012/12/23	天皇誕生日
2012/12/24	休日
2013/1/1	元日
2013/1/14	成人の日
2013/2/11	建国記念の日
2013/3/20	春分の日
2013/4/29	昭和の日
2013/5/3	憲法記念日
2013/5/4	みどりの日
2013/5/5	こどもの日
2013/5/6	休日
2013/7/15	海の日
2013/9/16	敬老の日
2013/9/23	秋分の日
2013/10/14	体育の日
2013/11/3	文化の日
2013/11/4	休日
2013/11/23	勤労感謝の日
2013/12/23	天皇誕生日
2014/1/1	元日
2014/1/13	成人の日
2014/2/11	建国記念の日
2014/3/21	春分の日
2014/4/29	昭和の日
2014/5/3	憲法記念日
2014/5/4	みどりの日
2014/5/5	こどもの日
2014/5/6	休日
2014/7/21	海の日
2014/9/15	敬老の日
2014/9/23	秋分の日
2014/10/13	体育の日
2014/11/3	文化の日
2014/11/23	勤労感謝の日
2014/11/24	休日
2014/12/23	天皇誕生日
2015/1/1	元日
2015/1/12	成人の日
2015/2/11	建国記念の日
2015/3/21	春分の日
2015/4/29	昭和の日
2015/5/3	憲法記念日
2015/5/4	みどりの日
2015/5/5	こどもの日
2015/5/6	休日
2015/7/20	海の日
2015/9/21	敬老の日
2015/9/22	休日
2015/9/23	秋分の日
2015/10/12	体育の日
2015/11/3	文化の日
2015/11/23	勤労感謝の日
2015/12/23	天皇誕生日
2016/1/1	元日
2016/1/11	成人の日
2016/2/11	建国記念の日
2016/3/20	春分の日
2016/3/21	休日
2016/4/29	昭和の日
2016/5/3	憲法記念日
2016/5/4	みどりの日
2016/5/5	こどもの日
2016/7/18	海の日
2016/8/11	山の日
2016/9/19	敬老の日
2016/9/22	秋分の日
2016/10/10	体育の日
2016/11/3	文化の日
2016/11/23	勤労感謝の日
2016/12/23	天皇誕生日
2017/1/1	元日
2017/1/2	休日
2017/1/9	成人の日
2017/2/11	建国記念の日
2017/3/20	春分の日
2017/4/29	昭和の日
2017/5/3	憲法記念日
2017/5/4	みどりの日
2017/5/5	こどもの日
2017/7/17	海の日
2017/8/11	山の日
2017/9/18	敬老の日
2017/9/23	秋分の日
2017/10/9	体育の日
2017/11/3	文化の日
2017/11/23	勤労感謝の日
2017/12/23	天皇誕生日
2018/1/1	元日
2018/1/8	成人の日
2018/2/11	建国記念の日
2018/2/12	休日
2018/3/21	春分の日
2018/4/29	昭和の日
2018/4/30	休日
2018/5/3	憲法記念日
2018/5/4	みどりの日
2018/5/5	こどもの日
2018/7/16	海の日
2018/8/11	山の日
2018/9/17	敬老の日
2018/9/23	秋分の日
2018/9/24	休日
2018/10/8	体育の日
2018/11/3	文化の日
2018/11/23	勤労感謝の日
2018/12/23	天皇誕生日
2018/12/24	休日
2019/1/1	元日
2019/1/14	成人の日
2019/2/11	建国記念の日
2019/3/21	春分の日
2019/4/29	昭和の日
2019/4/30	休日
2019/5/1	休日（祝日扱い）
2019/5/2	休日
2019/5/3	憲法記念日
2019/5/4	みどりの日
2019/5/5	こどもの日
2019/5/6	休日
2019/7/15	海の日
2019/8/11	山の日
2019/8/12	休日
2019/9/16	敬老の日
2019/9/23	秋分の日
2019/10/14	体育の日（スポーツの日）
2019/10/22	休日（祝日扱い）
2019/11/3	文化の日
2019/11/4	休日
2019/11/23	勤労感謝の日
2020/1/1	元日
2020/1/13	成人の日
2020/2/11	建国記念の日
2020/2/23	天皇誕生日
2020/2/24	休日
2020/3/20	春分の日
2020/4/29	昭和の日
2020/5/3	憲法記念日
2020/5/4	みどりの日
2020/5/5	こどもの日
2020/5/6	休日
2020/7/23	海の日
2020/7/24	スポーツの日
2020/8/10	山の日
2020/9/21	敬老の日
2020/9/22	秋分の日
2020/11/3	文化の日
2020/11/23	勤労感謝の日
2021/1/1	元日
2021/1/11	成人の日
2021/2/11	建国記念の日
2021/2/23	天皇誕生日
2021/3/20	春分の日
2021/4/29	昭和の日
2021/5/3	憲法記念日
2021/5/4	みどりの日
2021/5/5	こどもの日
2021/7/22	海の日
2021/7/23	スポーツの日
2021/8/8	山の日
2021/8/9	休日
2021/9/20	敬老の日
2021/9/23	秋分の日
2021/11/3	文化の日
2021/11/23	勤労感謝の日
2022/1/1	元日
2022/1/10	成人の日
2022/2/11	建国記念の日
2022/2/23	天皇誕生日
2022/3/21	春分の日
2022/4/29	昭和の日
2022/5/3	憲法記念日
2022/5/4	みどりの日
2022/5/5	こどもの日
2022/7/18	海の日
2022/8/11	山の日
2022/9/19	敬老の日
2022/9/23	秋分の日
2022/10/10	スポーツの日
2022/11/3	文化の日
2022/11/23	勤労感謝の日
2023/1/1	元日
2023/1/2	休日
2023/1/9	成人の日
2023/2/11	建国記念の日
2023/2/23	天皇誕生日
2023/3/21	春分の日
2023/4/29	昭和の日
2023/5/3	憲法記念日
2023/5/4	みどりの日
2023/5/5	こどもの日
2023/7/17	海の日
2023/8/11	山の日
2023/9/18	敬老の日
2023/9/23	秋分の日
2023/10/9	スポーツの日
2023/11/3	文化の日
2023/11/23	勤労感謝の日
2024/1/1	元日
2024/1/8	成人の日
2024/2/11	建国記念の日
2024/2/12	休日
2024/2/23	天皇誕生日
2024/3/20	春分の日
2024/4/29	昭和の日
2024/5/3	憲法記念日
2024/5/4	みどりの日
2024/5/5	こどもの日
2024/5/6	休日
2024/7/15	海の日
2024/8/11	山の日
2024/8/12	休日
2024/9/16	敬老の日
2024/9/22	秋分の日
2024/9/23	休日
2024/10/14	スポーツの日
2024/11/3	文化の日
2024/11/4	休日
2024/11/23	勤労感謝の日
2025/1/1	元日
2025/1/13	成人の日
2025/2/11	建国記念の日
2025/2/23	天皇誕生日
2025/2/24	休日
2025/3/20	春分の日
2025/4/29	昭和の日
2025/5/3	憲法記念日
2025/5/4	みどりの日
2025/5/5	こどもの日
2025/5/6	休日
2025/7/21	海の日
2025/8/11	山の日
2025/9/15	敬老の日
2025/9/23	秋分の日
2025/10/13	スポーツの日
2025/11/3	文化の日
2025/11/23	勤労感謝の日
2025/11/24	休日
2026/1/1	元日
2026/1/12	成人の日
2026/2/11	建国記念の日
2026/2/23	天皇誕生日
2026/3/20	春分の日
2026/4/29	昭和の日
2026/5/3	憲法記念日
2026/5/4	みどりの日
2026/5/5	こどもの日
2026/5/6	休日
2026/7/20	海の日
2026/8/11	山の日
2026/9/21	敬老の日
2026/9/22	休日
2026/9/23	秋分の日
2026/10/12	スポーツの日
2026/11/3	文化の日
2026/11/23	勤労感謝の日





