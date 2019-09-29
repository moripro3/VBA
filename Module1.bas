Attribute VB_Name = "Module1"
Option Explicit

Sub main()

    Application.ScreenUpdating = False
    
    Dim objIE As InternetExplorer
    Set objIE = New InternetExplorer
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("list")
    
    Dim shell As Object
    Set shell = CreateObject("Shell.Application") '�V�F���I�u�W�F�N�g����
    
    Dim win As Object
    For Each win In shell.Windows '�N�����̃E�B���h�E�����ԂɃ`�F�b�N
    
        If win.Name = "Internet Explorer" Then '�N�����Ă�IE���擾
        
            Set objIE = win
            Exit For
            
        End If
    
    Next
    
    Call WriteShareholderBenefitsData(objIE)
    
    MsgBox "�I�����܂���"

End Sub

Sub waitIE(objIE)
    
    Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE '�ǂݍ��ݑ҂�
        DoEvents
    Loop
    
End Sub

'����D�҂̌������ʂ��V�[�g�ɏ����o��
Sub WriteShareholderBenefitsData(objIE As InternetExplorer)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("list")
    
    '�O�񏑂����݃f�[�^�̃N���A
    Range(ws.Cells(2, 1), ws.Cells(Rows.Count, 5)).ClearContents
    
    Dim r As Long, c As Long '�������ݐ�̃Z���̍s��ԍ�
    Dim cnt As Long
    
    Dim isLastPageFlag As Boolean
    isLastPageFlag = False

    '�������ʂ̍ŏI�y�[�W�ɒB����܂ŏ������J��Ԃ�
    Do While isLastPageFlag = False
        
        '�������ʂ�HTML��ǂݍ���
        Dim htmlDoc As HTMLDocument
        Set htmlDoc = objIE.document
        
        Dim yuutaiTbl As IHTMLElement
        Set yuutaiTbl = htmlDoc.getElementById("item01") 'id��=item01
        
        Dim tdTags As IHTMLElementCollection
        Set tdTags = yuutaiTbl.getElementsByTagName("td")
        
        Dim tdTag As IHTMLElement
        For Each tdTag In tdTags
        
            r = cnt \ 5 + 2 '�������ݐ�̍s�ԍ�
            c = cnt Mod 5 + 1 '�������ݐ�̗�ԍ�
            
            ws.Cells(r, c).Value = tdTag.innerText
            
            cnt = cnt + 1
        
        Next tdTag
        
        ' class���͓���y�[�W���ŏd���\�Ȃ̂ŁA���Ԗڂ����w�肷��
        ' �����ł�1�Ԗڂ�next���擾�������̂ŁA(0)���w�肷��
        Dim nextPageLink As IHTMLElement '�u����10���v�̃����N
        Set nextPageLink = htmlDoc.getElementsByClassName("next")(0) 'class��=next
        
        If nextPageLink Is Nothing = False Then '�������̃y�[�W������Ȃ�
            
            nextPageLink.getElementsByTagName("a")(0).Click 'a�^�O���N���b�N
            Call waitIE(objIE) '��ʑJ�ڂ�ҋ@����
        
        Else '�������ʂ̍ŏI�y�[�W�Ȃ�t���O�𗧂Ă�
            
           isLastPageFlag = True
        
        End If
        
        Set htmlDoc = Nothing '���̃y�[�W��HTML�Q�Ƃ���������j��
        
    Loop
    
End Sub
