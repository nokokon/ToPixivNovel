Attribute VB_Name = "ToPixvNovel"
Option Explicit

'** Document **
'
'# ToPixvNovel
'
'## ライセンス
' MIT License 2026 (c) C
'
'## 概要
'  Word文書をpixiv小説記法に変換します。
'
'## 想定環境
'  Microsoft(c) Word 2019, Microsoft 365 Word
'  ※ Word 2013 以降なら動くはずです。
'
'## 初期設定
'  参照設定で "Microsoft Forms 2.0 Object Library" をチェックしてください。
'
'## スタイル名
'Word文書に以下のスタイルを用意してください。
'- 見出し           ... 段落スタイル
'- 標準             ... 段落スタイル
'- 強調太字         ... 文字スタイル
'- 斜体             ... 文字スタイル
'- 強調太字＋斜体   ... 文字スタイル
'
'## 傍点
'以下のルビを傍点として取り扱います。
'- 傍点             ... ’'‘・●丶、UNICODEの傍点文字
'- 中抜き傍点       ... ◯○。゜UNICODEの中抜き傍点文字
'
'## 改ページ
'以下を改ページとして取り扱います。
'- 改ページ
'- セクション
'- 改段
'
'## 改段落
'以下を改段落として取り扱います。
'- 改段落
'- 改行 (Shift + Enter で入力したもの)
'
'## 使い方
'1. "Run" プロシージャーを実行してください。
'2. pixiv小説投稿用URLが開きます。
'3. クリップボードに格納された文字列を貼り付けてください。
'4. 必要に応じて、小説を投稿してください。
'
'**************

Private pPara As Paragraph
Private pChar As Range
Private pLineString As String
Private pSubjectRange As Boolean
Private pBoldRange As Boolean
Private pItalicRange As Boolean
Private pBoldItalicRange As Boolean

Private Sub getSubjectStyle()
    If pPara.style = "見出し" Then
        pLineString = "[chapter:" & pLineString & "]": Exit Sub
    End If

    pLineString = pLineString & vbLf
End Sub

Private Function getStyleTag( _
    ByVal style_name As String, _
    ByVal bool As Boolean, _
    ByVal start_tag As String, _
    ByVal end_tag As String _
) As Boolean
    getStyleTag = bool
    If (bool) And (Not pChar.style = style_name) Then
        getStyleTag = False
        pLineString = pLineString & end_tag: Exit Function
    ElseIf (Not bool) And (pChar.style = style_name) Then
        getStyleTag = True
        pLineString = pLineString & start_tag: Exit Function
    End If
End Function

Private Sub getBoldStyle()
    pBoldRange = getStyleTag("強調太字", pBoldRange, "[b:", "]")
End Sub

Private Sub getItalicStyle()
    pItalicRange = getStyleTag("斜体", pItalicRange, "[i:", "]")
End Sub

Private Sub getBoldItalicStyle()
    pBoldItalicRange = getStyleTag("強調太字＋斜体", pBoldItalicRange, "[b:[i:", "]]")
End Sub

Private Sub getNormalText()
    Select Case pChar.Text
        Case vbLf
            Call clearTextFormat: Exit Sub
        Case Chr(11) '改行
            Call clearTextFormat: Exit Sub
        Case Chr(12) '改ページ/改セクション
            pLineString = pLineString & "[newpage]": Call clearTextFormat: Exit Sub
        Case vbCr '改段落
            Call clearTextFormat: Exit Sub
        Case Chr(14) '段区切り
            pLineString = pLineString & vbLf & "[newpage]": Call clearTextFormat: Exit Sub
        Case Else
            pLineString = pLineString & pChar.Text: Exit Sub
    End Select
End Sub

Private Sub clearTextFormat()
    If pBoldRange Then
        pLineString = pLineString & "]": pBoldRange = False
    End If

    If pItalicRange Then
        pLineString = pLineString & "]": pItalicRange = False
    End If

    If pBoldItalicRange Then
        pLineString = pLineString & "]]": pBoldItalicRange = False
    End If
End Sub

Private Function getRubyMatches(field_code As String) As Object
    With CreateObject("VBScript.RegExp")
        .IgnoreCase = False
        .Global = True
        .Pattern = "EQ .+? jc[0-9]{0,} .+? "".+?"" \\\* hps[0-9]{0,} .*?\(.*?\((.*?)\),(.*?)\)"
        Set getRubyMatches = .Execute(field_code)
    End With
End Function

Private Sub getRubyText()
    Call checkRubyTextFormat

    Dim regMatches As Object: _
        Set regMatches = getRubyMatches(pChar.Fields(1).Code.Text)

    Select Case regMatches(0).SubMatches(0)
        Case "’", "'", "‘", "・", "●", "丶", "、", ChrW(65093)
            pLineString = pLineString & "[[emphasismark:" & regMatches(0).SubMatches(1) & ">" & _
                ChrW(65093) & "]]"
        Case "◯", "○", "。", "゜", ChrW(65094)
            pLineString = pLineString & "[[emphasismark:" & regMatches(0).SubMatches(1) & ">" & _
                ChrW(65094) & "]]"
        Case Else
            pLineString = pLineString & "[[rb:" & regMatches(0).SubMatches(1) & " > " & _
                regMatches(0).SubMatches(0) & "]]"
    End Select
End Sub

Private Sub checkRubyTextFormat()
    Dim memChar As Range: Set memChar = pChar

    Set pChar = pChar.Fields(1).Code.Characters(1)
    Call convStyleRange

    Set pChar = memChar
End Sub

Private Sub convStyleRange()
    Call getBoldStyle
    Call getItalicStyle
    Call getBoldItalicStyle
End Sub

Private Function getCharactersText() As String
    Dim retString As String: retString = ""
    Dim paras As Paragraphs: Set paras = ActiveDocument.Paragraphs
    Dim chars As Characters

    For Each pPara In paras
        pLineString = ""
        Set chars = pPara.Range.Characters

        For Each pChar In chars
            If pChar.Text = Chr(21) Then
                Call getRubyText 'ルビ文字
            Else
                Call convStyleRange
                Call getNormalText
            End If
        Next pChar
        Call getSubjectStyle

        retString = retString & pLineString
    Next pPara

    getCharactersText = _
        Replace(Left(retString, Len(retString) - 1), vbLf, vbCrLf)
End Function

Private Sub setClip(ByVal clip_string As String)
    Dim cb As New DataObject
    With cb
        .SetText clip_string
        .PutInClipboard
    End With
End Sub

Private Sub openURL(ByVal url_strings As String)
    With CreateObject("Wscript.Shell")
        .Run url_strings, 3
    End With
End Sub

Private Function checkFields() As Boolean
    checkFields = False
    Dim fieldsList As Fields: Set fieldsList = ActiveDocument.Fields
    Dim fieldItem As Field
    For Each fieldItem In fieldsList
        If getRubyMatches(fieldItem.Code) Is Nothing Then Exit Function
    Next fieldItem
    checkFields = True
End Function

Private Function checkTables() As Boolean
    checkTables = False
    If ActiveDocument.Tables.Count > 0 Then Exit Function
    checkTables = True
End Function

Public Sub Run()
    If Not checkFields Then _
        MsgBox "ルビ以外のフィールドコードが含まれています。", vbExclamation, "ご確認ください": _
        Exit Sub

    If Not checkTables Then _
        MsgBox "罫表が含まれています。", vbExclamation, "ご確認ください": _
        Exit Sub

    setClip getCharactersText
    openURL "https://www.pixiv.net/novel/upload.php"
End Sub
