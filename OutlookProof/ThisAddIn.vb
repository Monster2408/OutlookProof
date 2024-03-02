Imports Microsoft.Office.Interop.Outlook

Public Class ThisAddIn

    ' 外部メールで.co.jpで終わるアドレスから届いたメールを格納するフォルダの名前
    Const COMPANY_DIR_NAME = "外部メール(co.jp含)"
    ' 外部メールを格納するフォルダの名前
    Const UNKNOWN_DIR_NAME = "外部メール"
    ' どのドメインを許可するか
    ' オフライン時のアイテムを何件検索するか
    Const CHECK_OFFLINE_ITEM_SIZE = 500

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    '
    ' MODULE        : Application_NewMailEx
    ' ID            :
    ' VERSION       : v2.0.0
    ' ABSTRACT      : 新規メール受信時に発火するイベント
    ' FUNCTION      : Outlook Event
    ' NOTE          :
    ' RETURN        :
    ' CREATE        : 2024-02-23    Monster2408
    ' UPDATE        :
    '
    Private Sub Application_NewMailEx(ByVal EntryIDCollection As String) Handles Application.NewMailEx
        'Outlookの機能にアクセスするためのMAPIオブジェクトを取得
        Dim ns As Outlook.NameSpace
        Dim objOL As Outlook.Application
        objOL = New Outlook.Application
        ns = objOL.GetNamespace("MAPI")

        '受信トレイのフォルダーオブジェクトを取得
        Dim myFld As Outlook.Folder
        myFld = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
        Dim objMail As Object
        ' 受信アイテムを取得
        objMail = ns.GetItemFromID(EntryIDCollection)
        ' 受信アイテムがメールの場合のみ処理
        If objMail.MessageClass = "IPM.Note" Then
            If objMail.Parent Is myFld Then
                CheckMailDomain(objMail, False)
            End If
        End If
    End Sub

    '
    ' MODULE        : Application_Startup
    ' ID            :
    ' VERSION       : v2.0.0
    ' ABSTRACT      : Outlook起動時に発火するイベント
    ' FUNCTION      : Outlook Event
    ' NOTE          :
    ' RETURN        :
    ' CREATE        : 2024-02-23    Monster2408
    ' UPDATE        :
    '
    Private Sub Application_Startup()
        ' オフライン時のアイテムを再確認する
        Call checkOfflineItems(CHECK_OFFLINE_ITEM_SIZE)
    End Sub

    '
    ' MODULE        : isEnableDomain
    ' ID            :
    ' VERSION       : v2.0.0
    ' ABSTRACT      : ホワイトリストドメインチェッカー
    ' FUNCTION      : ホワイトリストドメインを確認します
    ' NOTE          :
    ' RETURN        : ホワイトリストドメインかどうか
    ' CREATE        : 2024-02-23    Monster2408
    ' UPDATE        :
    '
    Private Function isEnableDomain(ByVal address As String) As Boolean
        isEnableDomain = False
        If address Like "*@*.example.co.jp" Then
            isEnableDomain = True
        End If
        If address Like "*@example.co.jp" Then
            isEnableDomain = True
        End If
        If address Like "*@another.co.jp" Then
            isEnableDomain = True
        End If
        If address Like "*@*.another.co.jp" Then
            isEnableDomain = True
        End If
        If address Like "*@another1.co.jp" Then
            isEnableDomain = True
        End If
        If address Like "*@*.another1.co.jp" Then
            isEnableDomain = True
        End If
    End Function

    '
    ' MODULE        : getCompanyMailBox
    ' PARAMETER     : isApl<Boolean>: アプリケーションかどうか
    ' ID            :
    ' VERSION       : v2.0.0
    ' ABSTRACT      : 企業メール(.co.jp)を格納するディレクトリ
    ' FUNCTION      : 企業メール(.co.jp)を格納するディレクトリを取得します
    ' NOTE          :
    ' RETURN        : 企業メール(.co.jp)を格納するディレクトリ
    ' CREATE        : 2024-02-23    Monster2408
    ' UPDATE        : 2024-03-02    Monster2s408: エラーが発生するため修正
    '
    Private Function getCompanyMailBox(ByVal isApl As Boolean) As Outlook.Folder
        'Outlookの機能にアクセスするためのMAPIオブジェクトを取得
        Dim ns As Outlook.NameSpace
        Dim objOL As Outlook.Application
        objOL = New Outlook.Application
        ns = objOL.GetNamespace("MAPI")

        '受信トレイのフォルダーオブジェクトを取得
        Dim myFld As Outlook.Folder
        myFld = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)

        Dim missing_dir As Boolean
        missing_dir = False

        Dim sub_folder As Outlook.Folder

        For Each sub_folder In myFld.Folders
            If sub_folder.Name = COMPANY_DIR_NAME Then
                missing_dir = True
                getCompanyMailBox = sub_folder
            End If
        Next

        If missing_dir = False Then
            sub_folder = myFld.Folders.Add(COMPANY_DIR_NAME)
            getCompanyMailBox = sub_folder
        End If
    End Function

    '
    ' MODULE        : getUnknownMailBox
    ' PARAMETER     : isApl<Boolean>: アプリケーションかどうか
    ' ID            :
    ' VERSION       : v2.0.0
    ' ABSTRACT      : 外部メールを格納するディレクトリ
    ' FUNCTION      : 外部メールを格納するディレクトリを取得します
    ' NOTE          :
    ' RETURN        : 外部メールを格納するディレクトリ
    ' CREATE        : 2024-02-23    Monster2408
    ' UPDATE        : 2024-03-02    Monster2408: エラーが発生するため修正
    '
    Private Function getUnknownMailBox(ByVal isApl As Boolean) As Outlook.Folder
        'Outlookの機能にアクセスするためのMAPIオブジェクトを取得
        Dim ns As Outlook.NameSpace
        Dim objOL As Outlook.Application
        objOL = New Outlook.Application
        ns = objOL.GetNamespace("MAPI")

        '受信トレイのフォルダーオブジェクトを取得
        Dim myFld As Outlook.Folder
        myFld = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)

        Dim missing_dir As Boolean
        missing_dir = False

        Dim sub_folder As Outlook.Folder

        For Each sub_folder In myFld.Folders
            If sub_folder.Name = UNKNOWN_DIR_NAME Then
                missing_dir = True
                getUnknownMailBox = sub_folder
            End If
        Next

        If missing_dir = False Then
            sub_folder = myFld.Folders.Add(UNKNOWN_DIR_NAME)
            getUnknownMailBox = sub_folder
        End If
    End Function

    '
    ' MODULE        : checkOfflineItems
    ' ID            :
    ' VERSION       : v2.0.0
    ' ABSTRACT      : メールフォルダの確認
    ' FUNCTION      : 受信フォルダを対象に最新メッセージを確認する
    ' NOTE          :
    ' RETURN        :
    ' CREATE        : 2024-02-23    Monster2408
    ' UPDATE        :
    '
    Sub checkOfflineItems(ByVal amount As Integer)
        Dim objOL As Object
        Dim objNAMESPC As Outlook.NameSpace
        Dim myfolders As Object
        Dim myInbox As Object
        Dim myfolder As Object
        Dim myItems As Outlook.Items

        'Outlookの機能にアクセスするためのMAPIオブジェクトを取得
        Dim ns As Outlook.NameSpace
        objOL = New Outlook.Application
        ns = objOL.GetNamespace("MAPI")
        '受信トレイのフォルダーオブジェクトを取得
        Dim myFld As Outlook.Folder
        myFld = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
        'メールアイテムを取得する
        Dim myItem As Object
        myItems = myFld.Items
        myItems.Sort("[受信日時]", True)

        Dim num As Integer
        num = 0
        For Each myItem In myItems
            'サンプルとしてイミディエイトウィンドウに件名を表示
            num = num + 1
            If num > amount Then
                Exit For
            End If
            If myItem.UnRead Then
                If myItem.MessageClass = "IPM.Note" Then
                    CheckMailDomain(myItem, True)
                End If
            End If
        Next myItem
    End Sub

    '
    ' MODULE        : CheckMailDomain
    ' ID            :
    ' VERSION       : v2.0.0
    ' ABSTRACT      : メールアイテムの仕分け
    ' FUNCTION      : 送信者アドレスを基に社内メール/企業メール/外部メールに仕分けする
    ' NOTE          :
    ' RETURN        :
    ' CREATE        : 2024-02-23    Monster2408
    ' UPDATE        :
    '
    Private Sub CheckMailDomain(ByVal objMail As MailItem, ByVal isApl As Boolean)
        ' 差出人アドレスのドメインのチェック
        ' MsgBox objMail.SenderEmailAddress
        ' メールアドレスを受信トレイに残していい場合は処理を中止する
        If isEnableDomain(GetSenderEmailAddress(objMail)) Then
            Exit Sub
        End If
        ' メールアドレスが.co.jpで終わる物
        If GetSenderEmailAddress(objMail) Like "*@*.co.jp" Then
            ' 外部メール(co.jp含)へ移動
            MoveToCompanyMail(objMail, isApl)
        Else
            ' 外部メールへ移動
            MoveToUnknownMail(objMail, isApl)
        End If
    End Sub

    '
    ' MODULE        : GetSenderEmailAddress
    ' ID            :
    ' VERSION       : v2.0.0
    ' ABSTRACT      : 送信者アドレスを取得
    ' FUNCTION      : Exchangeアカウントを含む送信者アドレスを取得する
    ' NOTE          :
    ' RETURN        :
    ' CREATE        : 2024-02-23    Monster2408
    ' UPDATE        :
    '
    Private Function GetSenderEmailAddress(ByRef oItem As MailItem) As String

        Dim PR_SMTP_ADDRESS As String
        Dim oSender As AddressEntry
        Dim oExUser As ExchangeUser


        PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

        ' Exchange 以外
        If oItem.SenderEmailType <> "EX" Then
            GetSenderEmailAddress = oItem.SenderEmailAddress
            Exit Function
        End If

        oSender = oItem.Sender

        If oSender.AddressEntryUserType = Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry _
            Or oSender.AddressEntryUserType = Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry Then

            oExUser = oSender.GetExchangeUser
            GetSenderEmailAddress = oExUser.PrimarySmtpAddress
            Exit Function
        End If

        GetSenderEmailAddress = oSender.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)

    End Function

    '
    ' MODULE        : MoveToCompanyMail
    ' PARAMETER     : objMail<MailItem>: メールアイテム
    '               : isApl<Boolean>:    アプリケーションかどうか
    ' ID            :
    ' VERSION       : v2.0.0
    ' ABSTRACT      : 企業メールフォルダへ移動させる
    ' FUNCTION      : 企業メールフォルダへ移動させる
    ' NOTE          :
    ' RETURN        :
    ' CREATE        : 2024-02-23    Monster2408
    ' UPDATE        :
    '
    Private Sub MoveToCompanyMail(ByVal objMail As MailItem, ByVal isApl As Boolean)
        On Error Resume Next 'エラー無視

        Dim subFld As Outlook.Folder
        subFld = getCompanyMailBox(isApl)
        objMail.Move(subFld)
        On Error GoTo 0  'エラー無視を解除
    End Sub
    '
    ' MODULE        : MoveToUnknownMail
    ' PARAMETER     : objMail<MailItem>: メールアイテム
    '               : isApl<Boolean>:    アプリケーションかどうか
    ' ID            :
    ' VERSION       : v2.0.0
    ' ABSTRACT      : 外部メールフォルダへ移動させる
    ' FUNCTION      : 外部メールフォルダへ移動させる
    ' NOTE          :
    ' RETURN        :
    ' CREATE        : 2024-02-23    Monster2408
    ' UPDATE        :
    '
    Private Sub MoveToUnknownMail(ByVal objMail As MailItem, ByVal isApl As Boolean)
        On Error Resume Next 'エラー無視

        Dim subFld As Outlook.Folder
        subFld = getUnknownMailBox(isApl)
        objMail.Move(subFld)
        On Error GoTo 0  'エラー無視を解除
    End Sub
End Class
