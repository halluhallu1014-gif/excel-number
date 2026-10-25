Option Explicit

'============================================================
' 18桁番号の目くらまし / Windows版Excel VBA7専用（32/64bit）
' 元コード互換：左から1 / 2222 / 3333 / 4444 / 55555を加算。
' 各桁は10で循環。復元は減算。暗号化を目的としない。
'
' コピーは半角数字18文字のみ。先頭0を保持する。
' ハイフン・改行・空白・アポストロフィは追加しない。
' 1-4-4-4-5のハイフン挿入は会社システム側に任せる。
'
' 入力：半角数字1～18文字のString、または非負整数Decimal。
' 短い入力は左を0埋め。空白・全角・符号・ハイフンは拒否。
' Double等は既に失われた桁を検出できないため受け付けない。
' 推奨：EncodeNumberText("012345678901234567")
' 互換：EncodeNumber(CDec("012345678901234567"))
' Excelセルも文字列で保存する。数値で丸めた18桁は復元不可。
'
' Text版は18桁String、旧名の関数は従来どおりDecimalを返す。
' この.basをVBEでインポートし、旧モジュールと置き換える。
'============================================================

Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal fmt As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal flags As Long, ByVal size As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal destination As LongPtr, ByVal source As LongPtr, ByVal size As LongPtr)
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Private Declare PtrSafe Sub SetLastError Lib "kernel32" (ByVal errorCode As Long)

Private Const NUMBER_LENGTH As Long = 18
Private Const OFFSETS As String = "122223333444455555"
Private Const CF_UNICODETEXT As Long = 13
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const CLIPBOARD_RETRIES As Long = 10
Private Const RETRY_DELAY_MS As Long = 25
Private Const ERR_INPUT As Long = vbObjectError + 1010
Private Const ERR_CLIPBOARD As Long = vbObjectError + 1100
Private Const ERR_TEST As Long = vbObjectError + 1200

'検証とゼロ埋めを集約。黙って空白除去や桁の切り捨てはしない。
Public Function FormatNumber18(ByVal numberValue As Variant) As String
    Dim digits As String
    Dim i As Long
    Dim code As Long
    Select Case VarType(numberValue)
        Case vbString
            digits = numberValue
        Case vbDecimal
            If numberValue < 0 Or numberValue >= CDec("1000000000000000000") Then
                Err.Raise ERR_INPUT, "FormatNumber18", "0以上18桁以内で指定してください。"
            End If
            If numberValue <> Fix(numberValue) Then
                Err.Raise ERR_INPUT, "FormatNumber18", "整数を指定してください。"
            End If
            '固定書式で桁区切りや指数表記を避ける。
            digits = Format$(numberValue, "0")
        Case Else
            Err.Raise ERR_INPUT, "FormatNumber18", "半角数字の文字列、またはDecimalで指定してください。"
    End Select
    If Len(digits) < 1 Or Len(digits) > NUMBER_LENGTH Then
        Err.Raise ERR_INPUT, "FormatNumber18", "半角数字1～18文字で指定してください。"
    End If
    'IsNumericは符号・指数表記なども許すので使わない。
    For i = 1 To Len(digits)
        code = AscW(Mid$(digits, i, 1))
        If code < 48 Or code > 57 Then
            Err.Raise ERR_INPUT, "FormatNumber18", "半角数字以外が含まれています。"
        End If
    Next i
    FormatNumber18 = String$(NUMBER_LENGTH - Len(digits), "0") & digits
End Function

Public Function EncodeNumberText(ByVal originalNumber As Variant) As String
    EncodeNumberText = TransformText(originalNumber, True)
End Function

Public Function DecodeNumberText(ByVal encodedNumber As Variant) As String
    DecodeNumberText = TransformText(encodedNumber, False)
End Function

'既存の呼び出し元との互換用。先頭0が必要な表示にはText版を使う。
Public Function EncodeNumber(ByVal originalNumber As Variant) As Variant
    EncodeNumber = CDec(EncodeNumberText(originalNumber))
End Function

Public Function DecodeNumber(ByVal encodedNumber As Variant) As Variant
    DecodeNumber = CDec(DecodeNumberText(encodedNumber))
End Function

Private Function TransformText(ByVal numberValue As Variant, ByVal isEncoding As Boolean) As String
    Dim sourceText As String
    Dim resultText As String
    Dim i As Long
    Dim digit As Long
    Dim offset As Long
    sourceText = FormatNumber18(numberValue)
    resultText = String$(NUMBER_LENGTH, "0")
    For i = 1 To NUMBER_LENGTH
        digit = AscW(Mid$(sourceText, i, 1)) - 48
        offset = AscW(Mid$(OFFSETS, i, 1)) - 48
        If isEncoding Then
            digit = (digit + offset) Mod 10
        Else
            '負数に対するVBAのModを避けるため10を加算する。
            digit = (digit - offset + 10) Mod 10
        End If
        Mid$(resultText, i, 1) = Chr$(48 + digit)
    Next i
    TransformText = resultText
End Function

Public Sub CopyDecodedNumber(ByVal encodedNumber As Variant)
    '入力検証・復元が成功してからクリップボードを変更する。
    CopyTextToClipboard DecodeNumberText(encodedNumber)
End Sub

'汎用Unicodeコピー。空文字は許可。途中NULは切り捨て防止で拒否。
Public Sub CopyTextToClipboard(ByVal textValue As String)
    Dim hMem As LongPtr
    Dim pMem As LongPtr
    Dim ownerWindow As LongPtr
    Dim byteCount As LongPtr
    Dim clipboardOpen As Boolean
    Dim memoryLocked As Boolean
    Dim attempt As Long
    Dim apiResult As Long
    Dim apiError As Long
    Dim savedNumber As Long
    Dim savedDescription As String
    Dim savedSource As String
    On Error GoTo Failed
    If InStr(1, textValue, vbNullChar, vbBinaryCompare) > 0 Then
        Err.Raise ERR_INPUT, "CopyTextToClipboard", "途中にNUL文字を含む文字列はコピーできません。"
    End If
    ownerWindow = Application.hWnd
    If ownerWindow = 0 Then
        Err.Raise ERR_CLIPBOARD, "CopyTextToClipboard", "Excelのウィンドウを取得できません。"
    End If
    '先にメモリを準備し、クリップボードの占有時間を短くする。
    'LongPtrに代入した後で終端2バイトを足す。
    byteCount = LenB(textValue)
    byteCount = byteCount + 2
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, byteCount)
    If hMem = 0 Then
        apiError = Err.LastDllError
        RaiseApiError "GlobalAlloc", apiError
    End If
    pMem = GlobalLock(hMem)
    If pMem = 0 Then
        apiError = Err.LastDllError
        RaiseApiError "GlobalLock", apiError
    End If
    memoryLocked = True
    If LenB(textValue) > 0 Then CopyMemory pMem, StrPtr(textValue), LenB(textValue)
    '終端2バイトはGlobalAllocでゼロ初期化済み。

    'GlobalUnlockは最終ロック解除の成功時にも0を返す。
    'LastErrorを先に0にし、戻り値0かつエラー非0だけを失敗扱い。
    SetLastError 0
    apiResult = GlobalUnlock(hMem)
    apiError = Err.LastDllError
    If apiResult = 0 And apiError <> 0 Then RaiseApiError "GlobalUnlock", apiError
    memoryLocked = False
    pMem = 0
    For attempt = 1 To CLIPBOARD_RETRIES
        apiResult = OpenClipboard(ownerWindow)
        If apiResult <> 0 Then
            clipboardOpen = True
            Exit For
        End If
        apiError = Err.LastDllError
        If attempt < CLIPBOARD_RETRIES Then Sleep RETRY_DELAY_MS
    Next attempt
    If Not clipboardOpen Then RaiseApiError "OpenClipboard", apiError

    'ここから既存内容を置き換える。EmptyClipboard成功後に
    'SetClipboardDataが失敗すると元内容は戻らない。エラーで通知する。
    If EmptyClipboard() = 0 Then
        apiError = Err.LastDllError
        RaiseApiError "EmptyClipboard", apiError
    End If
    If SetClipboardData(CF_UNICODETEXT, hMem) = 0 Then
        apiError = Err.LastDllError
        RaiseApiError "SetClipboardData", apiError
    End If
    '成功後はWindowsの所有物になる。以後GlobalFreeは禁止。
    hMem = 0
    If CloseClipboard() = 0 Then
        apiError = Err.LastDllError
        '開いた状態を維持し、Failedで閉じ直す。
        RaiseApiError "CloseClipboard（データ設定後）", apiError
    End If
    clipboardOpen = False
    Exit Sub
Failed:
    '後始末で元のエラーが消えないよう退避し、最後に再通知する。
    savedNumber = Err.Number
    savedDescription = Err.Description
    savedSource = Err.Source
    On Error Resume Next
    If memoryLocked Then apiResult = GlobalUnlock(hMem)
    If hMem <> 0 Then pMem = GlobalFree(hMem)
    If clipboardOpen Then apiResult = CloseClipboard()
    On Error GoTo 0
    Err.Raise savedNumber, savedSource, savedDescription
End Sub

Private Sub RaiseApiError(ByVal apiName As String, ByVal errorCode As Long)
    Err.Raise ERR_CLIPBOARD, "CopyTextToClipboard", _
        apiName & " に失敗しました。Win32エラー: " & CStr(errorCode)
End Sub

'自動テスト：クリップボードは変更しない。失敗はエラーで通知する。
Public Sub TestNumberSystem()
    Dim cases As Variant
    Dim item As Variant
    Dim encoded As String
    Dim digits As String
    Dim i As Long
    Dim d As Long
    Dim invalidObject As Object
    AssertText EncodeNumberText("0"), "122223333444455555", "ゼロの変換"
    AssertText EncodeNumberText("999999999999999999"), "011112222333344444", "加算の循環"
    AssertText EncodeNumberText("012345678901234567"), "134568901345689012", "固定期待値"
    AssertText DecodeNumberText("134568901345689012"), "012345678901234567", "固定復元"
    AssertText DecodeNumberText("000000000000000000"), "988887777666655555", "減算の循環"
    AssertText FormatNumber18("1"), "000000000000000001", "ゼロ埋め"
    AssertText FormatNumber18(CDec("999999999999999999")), "999999999999999999", "上限"
    AssertText FormatNumber18(CDec(1)), "000000000000000001", "Decimal整形"
    cases = Array("000000000000000000", "000000000000000001", _
        "012345678901234567", "123456789012345678", _
        "789012345678901234", "999999999999999999")
    For Each item In cases
        encoded = EncodeNumberText(item)
        AssertText DecodeNumberText(encoded), CStr(item), "文字列の往復"
        AssertText FormatNumber18(EncodeNumber(CDec(item))), encoded, "変換互換"
        AssertText FormatNumber18(DecodeNumber(CDec(encoded))), CStr(item), "復元互換"
        AssertText EncodeNumberText(DecodeNumberText(item)), CStr(item), "逆順の往復"
    Next item
    '全18位置に全数字を配置し、180通りで循環を確認する。
    For i = 1 To NUMBER_LENGTH
        For d = 0 To 9
            digits = String$(NUMBER_LENGTH, "0")
            Mid$(digits, i, 1) = Chr$(48 + d)
            AssertText DecodeNumberText(EncodeNumberText(digits)), digits, "各桁の往復"
        Next d
    Next i
    ExpectInvalid ""
    ExpectInvalid "1000000000000000000"
    ExpectInvalid " 123"
    ExpectInvalid "123 "
    ExpectInvalid "1-2345-6789-0123-45678"
    ExpectInvalid "１２３"
    ExpectInvalid "+1"
    ExpectInvalid "1.0"
    ExpectInvalid "1E2"
    ExpectInvalid "12" & vbCrLf
    ExpectInvalid "12" & vbNullChar & "3"
    ExpectInvalid CDec("-1")
    ExpectInvalid CDec(1) / CDec(10)
    ExpectInvalid CDec("1000000000000000000")
    ExpectInvalid CDbl(1)
    ExpectInvalid CSng(1)
    ExpectInvalid CLng(1)
    ExpectInvalid CCur(1)
    ExpectInvalid True
    ExpectInvalid Null
    ExpectInvalid Empty
    ExpectInvalid CVErr(2015)
    ExpectInvalid Array("123")
    Set invalidObject = New Collection
    ExpectInvalid invalidObject
    MsgBox "PASS：変換・復元・入力検証の自動テスト完了。" & vbCrLf & _
        "クリップボードと会社システムへの貼り付けは別途確認してください。", vbInformation
End Sub

Private Sub AssertText(ByVal actual As String, ByVal expected As String, ByVal testName As String)
    If StrComp(actual, expected, vbBinaryCompare) <> 0 Then
        Err.Raise ERR_TEST, "TestNumberSystem", testName & " に失敗しました。"
    End If
End Sub

Private Sub ExpectInvalid(ByVal value As Variant)
    Dim ignored As String
    Dim actualError As Long
    On Error Resume Next
    ignored = FormatNumber18(value)
    actualError = Err.Number
    On Error GoTo 0
    If actualError <> ERR_INPUT Then
        Err.Raise ERR_TEST, "TestNumberSystem", "入力エラーの検出に失敗しました。"
    End If
End Sub

'手動テスト：現在のクリップボードをテスト番号で置き換える。
'メモ帳などに貼り付け、先頭0を含め18桁であることを確認する。
'会社システム側の受入条件と自動ハイフン挿入も実際に確認する。
Public Sub TestClipboardManually()
    CopyDecodedNumber "134568901345689012"
    MsgBox "次の18桁をコピーしました。貼り付けて確認してください。" & vbCrLf & _
        "012345678901234567" & vbCrLf & _
        "システムの想定表示：0-1234-5678-9012-34567", vbInformation
End Sub
