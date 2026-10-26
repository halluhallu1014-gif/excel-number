Option Explicit

'============================================================
' 18桁番号の目くらまし / MSForms参照設定不要版・Windows版Excel（32/64bit）
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
' このファイル全文を標準モジュールへ貼り付け、旧コードを置き換える。
' API宣言は使用しない。会社環境での保存・コピー動作は要確認。
' Alt+F8：番号を変換する / 番号を復元してコピーする を実行する。
'============================================================



'参照設定・ユーザーフォームの追加は不要。
'MSFormsのDataObjectを実行時に生成する（遅延バインディング）。
'MSForms自体が未導入・登録不良・会社の制限で使えない場合は、
'生成時にエラーを通知する。参照設定不要でも部品への依存は残る。
'クラスIDはMSForms.DataObjectを識別する公開のIDであり、
'セキュリティ設定の変更やDLLの登録処理は行わない。
'コピーを設定したDataObjectをモジュール変数で保持する。
'次のコピー成功時に置き換える。Excel終了時に解放される。
Private mClipboard As Object

Private Const NUMBER_LENGTH As Long = 18
Private Const OFFSETS As String = "122223333444455555"
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


'汎用テキストコピー。数字18桁の経路はCopyDecodedNumberから呼ぶ。
'セルのCopyを使わないので、セル由来の区切りや改行を追加しない。
'失敗時は成功メッセージを出さず、元のエラー番号と説明を通知する。
Public Sub CopyTextToClipboard(ByVal textValue As String)
    Dim nextClipboard As Object
    Dim savedNumber As Long
    Dim savedDescription As String

    'NULを含む文字列は途中で切れる恐れがあるためコピー前に拒否。
    If InStr(1, textValue, vbNullChar, vbBinaryCompare) > 0 Then
        Err.Raise ERR_INPUT, "CopyTextToClipboard", _
            "途中にNUL文字を含む文字列はコピーできません。"
    End If

    On Error GoTo Failed
    Set nextClipboard = CreateClipboardObject()
    nextClipboard.SetText textValue, 1
    nextClipboard.PutInClipboard

    'PutInClipboardが成功してから保持先を置き換える。
    Set mClipboard = nextClipboard
    Exit Sub

Failed:
    savedNumber = Err.Number
    savedDescription = Err.Description
    Err.Raise ERR_CLIPBOARD, "CopyTextToClipboard", _
        "MSFormsでのコピーに失敗しました。" & vbCrLf & _
        "元のエラー番号：" & CStr(savedNumber) & vbCrLf & savedDescription
End Sub


'生成方法を一か所に集約。型宣言にも参照ライブラリを使わない。
Private Function CreateClipboardObject() As Object
    Dim savedNumber As Long
    Dim savedDescription As String
    On Error GoTo Failed
    Set CreateClipboardObject = CreateObject( _
        "new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    Exit Function
Failed:
    savedNumber = Err.Number
    savedDescription = Err.Description
    Err.Raise ERR_CLIPBOARD, "CreateClipboardObject", _
        "このPCでMSForms.DataObjectを生成できません。" & vbCrLf & _
        "元のエラー番号：" & CStr(savedNumber) & vbCrLf & savedDescription
End Function

'Alt+F8から使う入口。入力は文字列で受け、数値への変換を避ける。
Public Sub 番号を変換する()
    RunNumberDialog True
End Sub

Public Sub 番号を復元してコピーする()
    RunNumberDialog False
End Sub

Private Sub RunNumberDialog(ByVal isEncoding As Boolean)
    Dim inputText As String
    Dim outputText As String
    Dim promptText As String
    Dim titleText As String
    On Error GoTo Failed

    If isEncoding Then
        promptText = "元の番号を半角数字18桁・ハイフンなしで入力してください。"
        titleText = "番号の変換"
    Else
        promptText = "変換済みの番号を半角数字18桁・ハイフンなしで入力してください。"
        titleText = "番号の復元とコピー"
    End If
    inputText = InputBox(promptText, titleText)
    '標準InputBoxはキャンセルも空文字になる。どちらも何も変更しない。
    If Len(inputText) = 0 Then Exit Sub

    '手入力では入力漏れ防止のため18桁ちょうどを要求する。
    '既存APIの短い入力に対するゼロ埋め互換性はそのまま維持する。
    If Len(inputText) <> NUMBER_LENGTH Then
        Err.Raise ERR_INPUT, titleText, "半角数字18桁で入力してください。"
    End If
    If isEncoding Then
        outputText = EncodeNumberText(inputText)
    Else
        outputText = DecodeNumberText(inputText)
    End If
    CopyTextToClipboard outputText

    'コピー処理が正常終了した場合のみ成功を表示する。
    If isEncoding Then
        MsgBox "変換した18桁をコピーしました。" & vbCrLf & _
            "文字列として保存してください。", vbInformation, titleText
    Else
        MsgBox "元の18桁をコピーしました。" & vbCrLf & _
            "会社システムの入力欄に貼り付けてください。", vbInformation, titleText
    End If
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "処理できませんでした"
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


'実機確認：クリップボードをテスト番号で上書きする。
'MSFormsで読み戻して、先頭0・桁数・余分な改行がないことを確認。
'この検査は会社システムの貼り付け受入までは保証しない。
'実行後はメモ帳と会社システムの確認用入力欄でも確認する。
Public Sub TestClipboardManually()
    Dim readback As Object
    Dim actual As String
    On Error GoTo Failed

    CopyDecodedNumber "134568901345689012"
    Set readback = CreateClipboardObject()
    readback.GetFromClipboard
    If Not readback.GetFormat(1) Then
        Err.Raise ERR_TEST, "TestClipboardManually", "コピーしたテキストを取得できません。"
    End If
    actual = readback.GetText(1)
    AssertText actual, "012345678901234567", "コピー内容の読み戻し"

    MsgBox "PASS：コピー内容が18桁で一致しました。" & vbCrLf & _
        "貼り付け確認用：012345678901234567" & vbCrLf & _
        "システムの想定表示：0-1234-5678-9012-34567", vbInformation
    Exit Sub
Failed:
    MsgBox "FAIL：" & Err.Description, vbExclamation, "コピーの確認"
End Sub
Public Sub 組み込みテスト()
'**********************************
'Dim 元番号 As String
'Dim 変換番号 As String
'Dim 復元番号 As String
'
''好きな番号に変更して試せる
'元番号 = "012345678901234567"
'
''元番号 → 目くらまし用番号
'変換番号 = EncodeNumberText(元番号)
'
''目くらまし用番号 → 元番号
'復元番号 = DecodeNumberText(変換番号)
'
''復元して会社システム用にコピー
'CopyDecodedNumber 変換番号
'**********************************
    Dim originalText As String
    Dim encodedText As String
    Dim decodedText As String

    'ここを好きな番号に変更する。必ず引用符で囲む。
    originalText = "099998888777700001"

    '変換結果を取得する
    encodedText = EncodeNumberText(originalText)

    '復元結果を取得する
    decodedText = DecodeNumberText(encodedText)

    'VBEでCtrl+Gを押すと、結果を確認できる
    Debug.Print "元番号："; originalText
    Debug.Print "変換後："; encodedText
    Debug.Print "復元後："; decodedText

    '復元した番号を会社システム用にコピーする
    CopyDecodedNumber encodedText

End Sub

