Attribute VB_Name = "Module1"
Option Explicit

Sub main()

    Application.ScreenUpdating = False
    
    Dim objIE As InternetExplorer
    Set objIE = New InternetExplorer
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("list")
    
    Dim shell As Object
    Set shell = CreateObject("Shell.Application") 'シェルオブジェクト生成
    
    Dim win As Object
    For Each win In shell.Windows '起動中のウィンドウを順番にチェック
    
        If win.Name = "Internet Explorer" Then '起動してるIEを取得
        
            Set objIE = win
            Exit For
            
        End If
    
    Next
    
    Call WriteShareholderBenefitsData(objIE)
    
    MsgBox "終了しました"

End Sub

Sub waitIE(objIE)
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '読み込み待ち
        DoEvents
    Loop
    
End Sub

'株主優待の検索結果をシートに書き出す
Sub WriteShareholderBenefitsData(objIE As InternetExplorer)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("list")
    
    '前回書き込みデータのクリア
    Range(ws.Cells(2, 1), ws.Cells(Rows.Count, 5)).ClearContents
    
    Dim r As Long, c As Long '書き込み先のセルの行列番号
    Dim cnt As Long
    
    Dim isLastPageFlag As Boolean
    isLastPageFlag = False

    '検索結果の最終ページに達するまで処理を繰り返す
    Do While isLastPageFlag = False
        
        '検索結果のHTMLを読み込む
        Dim htmlDoc As HTMLDocument
        Set htmlDoc = objIE.document
        
        Dim yuutaiTbl As IHTMLElement
        Set yuutaiTbl = htmlDoc.getElementById("item01") 'id名=item01
        
        Dim tdTags As IHTMLElementCollection
        Set tdTags = yuutaiTbl.getElementsByTagName("td")
        
        Dim tdTag As IHTMLElement
        For Each tdTag In tdTags
        
            r = cnt \ 5 + 2 '書き込み先の行番号
            c = cnt Mod 5 + 1 '書き込み先の列番号
            
            ws.Cells(r, c).Value = tdTag.innerText
            
            cnt = cnt + 1
        
        Next tdTag
        
        ' class名は同一ページ内で重複可能なので、何番目かを指定する
        ' ここでは1番目のnextを取得したいので、(0)を指定する
        Dim nextPageLink As IHTMLElement '「次の10件」のリンク
        Set nextPageLink = htmlDoc.getElementsByClassName("next")(0) 'class名=next
        
        If nextPageLink Is Nothing = False Then 'もし次のページがあるなら
            
            nextPageLink.getElementsByTagName("a")(0).Click 'aタグをクリック
            Call waitIE(objIE) '画面遷移を待機する
        
        Else '検索結果の最終ページならフラグを立てる
            
           isLastPageFlag = True
        
        End If
        
        Set htmlDoc = Nothing 'このページのHTML参照をいったん破棄
        
    Loop
    
End Sub
