Imports System
Imports System.IO
Imports System.Net
Imports System.Security.Cryptography
Imports System.Text
Imports System.Globalization
Imports System.Web
Imports System.Text.RegularExpressions
Imports System.Xml
Imports Microsoft.VisualBasic

Public Class Form1

    'Form Load
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        oauth_consumer_key = TextBox7.Text

        oauth_consumer_secret = TextBox8.Text

        Dim enumValues As Array = System.[Enum].GetValues(GetType(User_Lookup_Type))

        For Each resourcez As User_Lookup_Type In enumValues
            ComboBox2.Items.Add(resourcez.ToString.Replace("_", ""))
        Next

        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 6
        ComboBox3.SelectedIndex = 1
        Label1.Text = String.Empty
        Label7.Text = String.Empty

    End Sub
    'Pause/Slow/Wait Function for Webbrowser
    Private Sub WBNavigatingWait(ByVal WB As WebBrowser, Optional Sec As Integer = 0)

        WebTimer = New Timer

        WebTimer.Interval = 1000

        WebTimerTick = 0
        AddHandler WebTimer.Tick, AddressOf WebTimer_Tick

        WebTimer.Start()

        Do Until (WB.ReadyState = WebBrowserReadyState.Complete AndAlso WebTimerTick > Sec)

            Application.DoEvents()

        Loop

        WebTimer.Stop()

        WebTimer.Dispose()

    End Sub

    'Authorize Account
    Private Sub AuthorizeAccount(ByVal UserName As String, Password As String)

        WB = New WebBrowser

        WB.ScriptErrorsSuppressed = True

        Label1.Text = String.Empty

        WB.Navigate("https://api.twitter.com/oauth/authorize?" & request_Token())

        WBNavigatingWait(WB)

        If Not WB.Document.GetElementById("username_or_email") = Nothing Then ' Checks for UserName Text Box

            WB.Document.GetElementById("username_or_email").SetAttribute("value", UserName)

            WB.Document.GetElementById("password").SetAttribute("value", Password)

            WB.Document.GetElementById("allow").InvokeMember("click")

            WBNavigatingWait(WB, 1)

            If PinVerify(WB) = True Then Label1.Text = "Successfully Authenticated" Else Label1.Text = "Authentication Failed"

        Else

            'Signs Out
            For Each Ele As HtmlElement In WB.Document.GetElementsByTagName("input")

                If (Not Ele.GetAttribute("value") = Nothing AndAlso Ele.GetAttribute("value").Equals("Sign out")) Then

                    Ele.InvokeMember("click")

                    WBNavigatingWait(WB, 0.5)

                    AuthorizeAccount(UserName, Password)

                    Exit For

                End If

            Next

        End If

        WB.Dispose()

    End Sub

    'Authorize Button
    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click

        AuthorizeAccount(TextBox2.Text, TextBox3.Text)

    End Sub
    'Send Tweet Button
    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click

        Label7.Text = String.Empty

        AuthenticateWith(oauth_consumer_key, My.Settings.okey, My.Settings.okeysecret)

        UpdateStatus(TextBox1.Text)

    End Sub
    'Timeline Lookup Button
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        AuthenticateWith(oauth_consumer_key, My.Settings.okey, My.Settings.okeysecret)

        ListView1.Items.Clear()

        For Each TStatus As TwitterStatus In User_TimeLine(TextBox4.Text, ComboBox1.SelectedItem)

            Dim NLVItem As New ListViewItem(TStatus.Text)

            ListView1.Items.Add(NLVItem)

        Next

    End Sub
    'User Info Lookup Button
    Private Sub Button2_Click_1(sender As System.Object, e As System.EventArgs) Handles Button2.Click

        TextBox5.Clear()

        AuthenticateWith(oauth_consumer_key, My.Settings.okey, My.Settings.okeysecret)

        Dim pair As KeyValuePair(Of String, String)

        For Each pair In Show_User_Info(TextBox6.Text, ComboBox2.SelectedIndex)

            TextBox5.Text += String.Format("{0}, " & ComboBox2.SelectedItem & ": {1}", pair.Key, pair.Value) & vbCrLf

        Next

    End Sub


    'Change to Lookup by Value
    Private Sub ComboBox3_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        Select Case ComboBox3.SelectedIndex
            Case 0
                infoLookupType = "user_id"
            Case 1
                infoLookupType = "screen_name"
        End Select
    End Sub

    'Tweet Textbox Length Display
    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged
        Label8.Text = TextBox1.TextLength & "/140"
    End Sub

End Class



Module TwitterConstants
    Public Const oauth_version As String = "1.0"
    Public Const oauth_signature_method As String = "HMAC-SHA1"
    Public Const oauth_callback As String = "oob"
    Public Const oauth_request_token_url As String = "https://api.twitter.com/oauth/request_token"
    Public Const oauth_access_token_request_url As String = "https://api.twitter.com/oauth/access_token"
    Public Const update_url As String = "http://api.twitter.com/1/statuses/update.json"
    Public Const timeline_url As String = "https://api.twitter.com/1.1/statuses/user_timeline.json"
    Public Const user_lookup_url As String = "https://api.twitter.com/1.1/users/lookup.json"

    Public oauth_consumer_key, oauth_consumer_secret As String
    Public oauth_signature, access_token, oauth_token, oauth_token_secret, user_name, header, baseString, request_URL, response_String, Method_String As String
    Public infoLookupType As String = "screen_name"
    Public response As WebResponse
    Public Doc As XmlDocument


    'Get Nonce
    Public Function Get_Nonce() As String
        Return New Random().Next(123400, Integer.MaxValue).ToString("X", CultureInfo.InvariantCulture)
    End Function
    'Get Time Stamp
    Public Function Get_Timestamp() As String
        Return Convert.ToInt64((DateTime.UtcNow - New DateTime(1970, 1, 1, 0, 0, 0, 0)).TotalSeconds, CultureInfo.CurrentCulture).ToString(CultureInfo.CurrentCulture)
    End Function
    'Make CompositeKey
    Public Function MakeCompositeKey() As String

        Return String.Concat(OAuthUrlEncode(oauth_consumer_secret), "&", OAuthUrlEncode(oauth_token_secret))

    End Function

    'Parse JSON Status
    Public Function ParseStatuses(Xml As String) As List(Of TwitterStatus)
        Dim enumerator As IEnumerator = Nothing
        Dim xmlDocument As New XmlDocument()
        xmlDocument.LoadXml(Xml)
        Dim xmlNamespaceStripper As New XmlNamespaceStripper(xmlDocument)
        xmlDocument = xmlNamespaceStripper.StrippedXmlDocument
        Dim twitterStatuses As New List(Of TwitterStatus)()
        Using TryCast(enumerator, IDisposable)
            If Not TypeOf (enumerator) Is IDisposable Then
                enumerator = xmlDocument.SelectNodes("//status").GetEnumerator()
                While enumerator.MoveNext()
                    Dim current As XmlNode = DirectCast(enumerator.Current, XmlNode)
                    twitterStatuses.Add(New TwitterStatus(current))
                End While
            End If
        End Using
        Return twitterStatuses
    End Function
    'Parse User JSON
    Public Function ParseUsers(Xml As String) As List(Of TwitterUser)
        Dim enumerator As IEnumerator = Nothing
        Dim xmlDocument As New XmlDocument()
        xmlDocument.LoadXml(Xml)
        Dim twitterUsers As New List(Of TwitterUser)()
        Using TryCast(enumerator, IDisposable)
            If Not TypeOf (enumerator) Is IDisposable Then
                enumerator = xmlDocument.SelectNodes("//user").GetEnumerator()
                While enumerator.MoveNext()
                    Dim current As XmlNode = DirectCast(enumerator.Current, XmlNode)
                    twitterUsers.Add(New TwitterUser(current))
                End While
            End If
        End Using
        Return twitterUsers
    End Function


    'Construct Headers
    Public Function MakeHeader(ByVal Header_Type As HeaderType, Optional PinCode As String = "") As String

        header = String.Empty

        Select Case Header_Type

            Case HeaderType.AuthHeader

                header =
                "OAuth" + " " + _
                "oauth_consumer_key=""" + OAuthUrlEncode(oauth_consumer_key) + """, " + _
                "oauth_nonce=""" + OAuthUrlEncode(Get_Nonce()) + """, " + _
                "oauth_signature_method=""" + OAuthUrlEncode(oauth_signature_method) + """," + _
                "oauth_timestamp=""" + OAuthUrlEncode(Get_Timestamp()) + """, " + _
                "oauth_version=""" + OAuthUrlEncode(oauth_version) + """, " + _
                "oauth_signature=""" + OAuthUrlEncode(oauth_signature) + """"

            Case HeaderType.GeneralHeader

                header = "OAuth" + " " + _
                "oauth_consumer_key=""" + OAuthUrlEncode(oauth_consumer_key) + """, " + _
                "oauth_nonce=""" + OAuthUrlEncode(Get_Nonce()) + """, " + _
                "oauth_signature=""" + OAuthUrlEncode(oauth_signature) + """, " + _
                "oauth_signature_method=""" + OAuthUrlEncode(oauth_signature_method) + """, " + _
                "oauth_timestamp=""" + OAuthUrlEncode(Get_Timestamp()) + """, " + _
                "oauth_token=""" + OAuthUrlEncode(oauth_token) + """, " + _
                "oauth_version=""" + OAuthUrlEncode(oauth_version) + """"

            Case HeaderType.VerifyHeader

                header = "OAuth" + " " + _
                "realm=""" + "Twitter API" + "," + _
                "oauth_token=""" + OAuthUrlEncode(access_token) + """," + _
                "oauth_verifier=""" + OAuthUrlEncode(PinCode) + """," + _
                "oauth_nonce=""" + OAuthUrlEncode(Get_Nonce) + """," + _
                "oauth_signature_method=""" + OAuthUrlEncode(oauth_signature_method) + """," + _
                "oauth_timestamp=""" + OAuthUrlEncode(Get_Timestamp) + """," + _
                "oauth_version=""" + OAuthUrlEncode(oauth_version) + """," + _
                "oauth_signature=""" + OAuthUrlEncode(oauth_signature) + """"

        End Select

        Return header

    End Function
    'Construct Signatures
    Public Function Construct_Signatures(ByVal SigType As Signature_Types,
                                          Request_URL As String,
                                          MethodType As Method_Type,
                                          Optional Status As String = "None",
                                          Optional user_nameORid As String = "None",
                                          Optional count_param As String = "None",
                                          Optional PinCode As String = "None")

        baseString = String.Empty

        Select Case SigType

            Case Signature_Types.TokenRequestSignature

                baseString =
                  "oauth_callback=" + oauth_callback +
                  "&oauth_consumer_key=" + oauth_consumer_key +
                  "&oauth_nonce=" + Get_Nonce() +
                  "&oauth_signature_method=" + oauth_signature_method +
                  "&oauth_timestamp=" + Get_Timestamp() +
                  "&oauth_version=" + oauth_version

            Case Signature_Types.VerificationSignature

                baseString =
                "oauth_callback=" + oauth_callback +
                "&oauth_token=" + access_token +
                "&oauth_verifier=" + PinCode +
                "&oauth_nonce=" + Get_Nonce() +
                "&oauth_signature_method=" + oauth_signature_method +
                "&oauth_timestamp=" + Get_Timestamp() +
                "&oauth_version=" + oauth_version

            Case Signature_Types.ShowUserSignatureByID

                baseString =
                "oauth_consumer_key=" + oauth_consumer_key + _
                "&oauth_nonce=" + Get_Nonce() + _
                "&oauth_signature_method=" + oauth_signature_method + _
                "&oauth_timestamp=" + Get_Timestamp() + _
                "&oauth_token=" + oauth_token + _
                "&oauth_version=" + oauth_version + _
                "&user_id=" + user_nameORid

            Case Signature_Types.ShowUserSignatureByScreenName

                baseString =
                "oauth_consumer_key=" + oauth_consumer_key + _
                "&oauth_nonce=" + Get_Nonce() + _
                "&oauth_signature_method=" + oauth_signature_method + _
                "&oauth_timestamp=" + Get_Timestamp() + _
                "&oauth_token=" + oauth_token + _
                "&oauth_version=" + oauth_version + _
                "&screen_name=" + user_nameORid

            Case Signature_Types.TimeLineSignature

                baseString =
                "count=" + count_param + _
                "&oauth_consumer_key=" + oauth_consumer_key + _
                "&oauth_nonce=" + Get_Nonce() + _
                "&oauth_signature_method=" + oauth_signature_method + _
                "&oauth_timestamp=" + Get_Timestamp() + _
                "&oauth_token=" + oauth_token + _
                "&oauth_version=" + oauth_version + _
                "&screen_name=" + user_nameORid

            Case Signature_Types.UpdateSignature

                baseString =
                "oauth_consumer_key=" + oauth_consumer_key + _
                "&oauth_nonce=" + Get_Nonce() + _
                "&oauth_signature_method=" + oauth_signature_method + _
                "&oauth_timestamp=" + Get_Timestamp() + _
                "&oauth_token=" + oauth_token + _
                "&oauth_version=" + oauth_version + _
                "&status=" + Status

        End Select

        baseString = String.Concat(MType(MethodType), "&", Uri.EscapeDataString(Request_URL), "&", Uri.EscapeDataString(baseString))

        Return Convert.ToBase64String((New HMACSHA1(Encoding.ASCII.GetBytes(MakeCompositeKey)).ComputeHash(Encoding.ASCII.GetBytes(baseString))))

    End Function

    'Generate WebRequest
    Public Function PerformWebRequest(ByVal MethodType As Method_Type,
                                      ByVal Request_URL As String,
                                      ByVal request_Type As Request_Types,
                                      Optional VarString As String = "None") As String

        response_String = String.Empty

        Dim request As HttpWebRequest = HttpWebRequest.Create(Request_URL)
        With request
            .Headers.Add("Authorization", header)
            .ContentType = "application/x-www-form-urlencoded"
            .ServicePoint.Expect100Continue = False
            .MaximumAutomaticRedirections = 4
            .MaximumResponseHeadersLength = 4
            .Method = MType(MethodType)
        End With

        Try

            Select Case request_Type



                Case Request_Types.Token_Request

                    response = request.GetResponse()

                    Dim oauthToken As String() = Split(New StreamReader(response.GetResponseStream()).ReadToEnd().ToString, "&")

                    access_token = oauthToken(0).Replace("oauth_token=", "")

                    response_String = oauthToken(0)

                Case Request_Types.Verify_Request

                    response = request.GetResponse()

                    response_String = New StreamReader(response.GetResponseStream()).ReadToEnd.ToString



                Case Request_Types.ShowUser_Request

                    response = request.GetResponse()

                    response_String = "{'?xml': {'@version': '1.0', '@standalone': 'no' },'root': { user:" & New StreamReader(response.GetResponseStream()).ReadToEnd.ToString & "}}"




                Case Request_Types.TimeLine_Request

                    response = request.GetResponse()

                    response_String = "{'?xml': {'@version': '1.0', '@standalone': 'no' },'root': { status:" & New StreamReader(response.GetResponseStream()).ReadToEnd.ToString & "}}"



                Case Request_Types.Update_Request

                    Dim postBody As String = "status=" + VarString

                    Using Stream As Stream = request.GetRequestStream()

                        Dim content As Byte() = ASCIIEncoding.ASCII.GetBytes(postBody)

                        Stream.Write(content, 0, content.Length)

                    End Using

                    response = request.GetResponse()

                    Form1.Label7.Text = "Sent!"

            End Select


        Catch ex As Exception

            Form1.Label1.Text = ex.Message

        End Try

        response.Close()

        request = Nothing

        Return response_String

    End Function

    'Encode Text
    Public Function OAuthUrlEncode(value As String) As String

        If value = Nothing Then Return Nothing

        Dim stringBuilder As New StringBuilder()
        Dim str As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_.~"
        Dim str1 As String = value
        Dim num As Integer = 0
        Dim length As Integer = str1.Length
        While num < length
            Dim chr As Char = str1(num)
            If str.IndexOf(chr) = -1 Then
                Dim bytes As Byte() = Encoding.UTF8.GetBytes(chr.ToString())
                Dim numArray As Byte() = bytes
                For i As Integer = 0 To CInt(numArray.Length) - 1
                    Dim num1 As Byte = numArray(i)
                    stringBuilder.AppendFormat("%{0:X2}", num1)
                Next
            Else
                stringBuilder.Append(chr)
            End If
            num += 1
        End While
        Return stringBuilder.ToString()
    End Function

    Property MType(ByVal MTyper As Method_Type) As String
        Get
            Return MTyper.ToString.Replace("_", "")
        End Get
        Set(ByVal Value As String)
            Method_String = Value
        End Set
    End Property

End Module
Module TwitterRequests

    'Authenticate Credentials
    Public Sub AuthenticateWith(ByVal Consumer As String, ByVal AuthToken As String, ByVal AuthTokenSecret As String)

        oauth_consumer_key = Consumer

        oauth_token = AuthToken

        oauth_token_secret = AuthTokenSecret

    End Sub

    'Get Request Tokens
    Public Function request_Token() As String

        oauth_token = String.Empty

        oauth_token_secret = String.Empty

        Try

            oauth_signature = Construct_Signatures(Signature_Types.TokenRequestSignature,
                                       oauth_request_token_url,
                                       Method_Type._POST)

            header = MakeHeader(HeaderType.AuthHeader)

            Return PerformWebRequest(Method_Type._POST,
                              oauth_request_token_url + "?" + "oauth_callback" + "=" + oauth_callback,
                              Request_Types.Token_Request)

        Catch ex As Exception

            Console.WriteLine(ex.Message)

            Return Nothing

        End Try

    End Function
    'Verify Pin and Pull Auth Tokens Etc..
    Public Function PinVerify(ByVal WB As WebBrowser) As Boolean

        Dim AccountAuthInfo As String()

        Try

            request_URL = oauth_access_token_request_url + "?" + "oauth_callback" + "=" + oauth_callback

            oauth_signature = Construct_Signatures(Signature_Types.VerificationSignature,
                                       String.Concat(oauth_access_token_request_url, "?", "oauth_callback", "=", oauth_callback),
                                       Method_Type._POST,
                                       "None", "None", "None",
                                       WB.Document.GetElementsByTagName("code").Item(0).OuterText)

            header = MakeHeader(HeaderType.VerifyHeader, WB.Document.GetElementsByTagName("code").Item(0).OuterText)

            AccountAuthInfo = Split(PerformWebRequest(Method_Type._POST, request_URL, Request_Types.Verify_Request), "&")

            oauth_token = AccountAuthInfo(0).Replace("oauth_token=", "")

            oauth_token_secret = AccountAuthInfo(1).Replace("oauth_token_secret=", "")

            My.Settings.okey = oauth_token

            My.Settings.okeysecret = oauth_token_secret

            Return True

        Catch ex As Exception

            Console.WriteLine("Ran into an error.  Check Password and Username")

            Return False

        End Try

    End Function

    'Send Tweet
    Public Sub UpdateStatus(status As String)

        AuthenticateWith(oauth_consumer_key, My.Settings.okey, My.Settings.okeysecret)

        oauth_signature = Construct_Signatures(Signature_Types.UpdateSignature,
                                   update_url,
                                   Method_Type._POST,
                                   OAuthUrlEncode(status))

        header = MakeHeader(HeaderType.GeneralHeader)

        PerformWebRequest(Method_Type._POST,
                          update_url,
                          Request_Types.Update_Request,
                          OAuthUrlEncode(status))

    End Sub
    'Timeline Lookup
    Public Function User_TimeLine(ByVal screen_name As String, countparam As String) As List(Of TwitterStatus)

        Dim Timeline_Results As New List(Of TwitterStatus)

        AuthenticateWith(oauth_consumer_key, My.Settings.okey, My.Settings.okeysecret)

        Try
            oauth_signature = Construct_Signatures(Signature_Types.TimeLineSignature, timeline_url, Method_Type._GET, "None", screen_name, countparam)

            header = MakeHeader(HeaderType.GeneralHeader)

            Dim Doc As XmlDocument = Newtonsoft.Json.JsonConvert.DeserializeXmlNode(PerformWebRequest(Method_Type._GET, _
                                                                                                     timeline_url + "?" + "screen_name" + "=" + screen_name + "&" + "count=" + countparam, _
                                                                                                     Request_Types.TimeLine_Request, _
                                                                                                     screen_name))

            For Each TStatus As TwitterStatus In ParseStatuses(Doc.InnerXml)

                Timeline_Results.Add(TStatus)

            Next

        Catch ex As Exception

            Console.WriteLine(ex.Message)

        End Try

        Return Timeline_Results

    End Function
    'UserLookup - Follwer Count
    Public Function Show_User_Info(ByVal screen_nameORid As String, Lookup_Type As Integer) As Dictionary(Of String, String)

        Dim UserInfoString As New Dictionary(Of String, String)

        AuthenticateWith(oauth_consumer_key, My.Settings.okey, My.Settings.okeysecret)

        Try

            Select Case infoLookupType

                Case "screen_name"

                    oauth_signature = Construct_Signatures(Signature_Types.ShowUserSignatureByScreenName,
                                       user_lookup_url,
                                       Method_Type._GET,
                                       "None",
                                       Uri.EscapeDataString(screen_nameORid),
                                       "None")

                Case "user_id"

                    oauth_signature = Construct_Signatures(Signature_Types.ShowUserSignatureByID,
                                       user_lookup_url,
                                       Method_Type._GET,
                                       "None",
                                       Uri.EscapeDataString(screen_nameORid),
                                       "None")

            End Select

            header = MakeHeader(HeaderType.GeneralHeader)

            Doc = Newtonsoft.Json.JsonConvert.DeserializeXmlNode(PerformWebRequest(Method_Type._GET,
                                                                                   user_lookup_url + "?" + infoLookupType + "=" + Uri.EscapeDataString(screen_nameORid),
                                                                                   Request_Types.ShowUser_Request,
                                                                                   Uri.EscapeDataString(screen_nameORid)))


           
            For Each TUser As TwitterUser In ParseUsers(Doc.InnerXml)
                Select Case Lookup_Type
                    Case User_Lookup_Type._CreatedAt
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.CreatedAt.ToShortDateString)
                    Case User_Lookup_Type._Description
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.Description.ToString)
                    Case User_Lookup_Type._FavoritesCount
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.FavoritesCount.ToString)
                    Case User_Lookup_Type._FollowersCount
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.FollowersCount.ToString)
                    Case User_Lookup_Type._FriendsCount
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.FriendsCount.ToString)
                    Case User_Lookup_Type._ID
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.ID.ToString)
                    Case User_Lookup_Type._Location
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.Location.ToString)
                    Case User_Lookup_Type._Name
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.Name.ToString)
                    Case User_Lookup_Type._Notifications
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.Notifications.ToString)
                    Case User_Lookup_Type._ProfileBackgroundColor
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.ProfileBackgroundColor.ToString)
                    Case User_Lookup_Type._ProfileBackgroundImageUrl
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.ProfileBackgroundImageUrl.ToString)
                    Case User_Lookup_Type._ProfileImageUrl
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.ProfileImageUrl.ToString)
                    Case User_Lookup_Type._ProfileLinkColor
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.ProfileLinkColor.ToString)
                    Case User_Lookup_Type._ProfileSidebarBorderColor
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.ProfileSidebarBorderColor.ToString)
                    Case User_Lookup_Type._ProfileSidebarFillColor
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.ProfileSidebarFillColor.ToString)
                    Case User_Lookup_Type._ProfileTextColor
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.ProfileTextColor.ToString)
                    Case User_Lookup_Type._Protected
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.Protected.ToString)
                    Case User_Lookup_Type._ScreenName
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.ScreenName.ToString)
                    Case User_Lookup_Type._Status
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.Status.Text)
                    Case User_Lookup_Type._StatusesCount
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.StatusesCount.ToString)
                    Case User_Lookup_Type._TimeZone
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.TimeZone.ToString)
                    Case User_Lookup_Type._Url
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.Url.ToString)
                    Case User_Lookup_Type._UTCOffset
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.UTCOffset.ToString)
                    Case User_Lookup_Type._Verified
                        UserInfoString.Add(TUser.ScreenName.ToString, TUser.Verified.ToString)
                End Select

            Next

        Catch ex As Exception

            Form1.TextBox5.Text = ex.Message

            Console.WriteLine(ex.Message)

        End Try

        Return UserInfoString

    End Function

End Module
Module ControlItems

    Public WebTimer As Timer
    Public WB As WebBrowser

    Public WebTimerTick As Integer

    'Web Timer Sub
    Public Sub WebTimer_Tick(sender As System.Object, e As System.EventArgs)

        WebTimerTick = WebTimerTick + 1

    End Sub

End Module



Public Class TwitterStatus
    Inherits XmlObjectBase
    Private _ID As Long

    Private _CreatedAt As DateTime

    Private _Text As String

    Private _Favorited As Boolean

    Private _InReplyToStatusID As Long

    Private _InReplyToUserID As String

    Private _InReplyToScreenName As String

    Private _IsDirectMessage As Boolean

    Private _Source As String

    Private _Truncated As Boolean

    Private _User As TwitterUser

    Private _RetweetedStatus As TwitterStatus

    Private _GeoLat As String

    Private _GeoLong As String

    Public Property CreatedAt() As DateTime
        Get
            Return Me._CreatedAt
        End Get
        Set(value As DateTime)
            Me._CreatedAt = value
        End Set
    End Property

    Public ReadOnly Property CreatedAtLocalTime() As DateTime
        Get
            Dim createdAt As DateTime = Me.CreatedAt
            Return createdAt.ToLocalTime()
        End Get
    End Property

    Public Property Favorited() As Boolean
        Get
            Return Me._Favorited
        End Get
        Set(value As Boolean)
            Me._Favorited = value
        End Set
    End Property

    Public Property GeoLat() As String
        Get
            Return Me._GeoLat
        End Get
        Set(value As String)
            Me._GeoLat = value
        End Set
    End Property

    Public Property GeoLong() As String
        Get
            Return Me._GeoLong
        End Get
        Set(value As String)
            Me._GeoLong = value
        End Set
    End Property

    Public Property ID() As Long
        Get
            Return Me._ID
        End Get
        Set(value As Long)
            Me._ID = value
        End Set
    End Property

    Public Property InReplyToScreenName() As String
        Get
            Return Me._InReplyToScreenName
        End Get
        Set(value As String)
            Me._InReplyToScreenName = value
        End Set
    End Property

    Public Property InReplyToStatusID() As Long
        Get
            Return Me._InReplyToStatusID
        End Get
        Set(value As Long)
            Me._InReplyToStatusID = value
        End Set
    End Property

    Public Property InReplyToUserID() As String
        Get
            Return Me._InReplyToUserID
        End Get
        Set(value As String)
            Me._InReplyToUserID = value
        End Set
    End Property

    Public Property IsDirectMessage() As Boolean
        Get
            Return Me._IsDirectMessage
        End Get
        Set(value As Boolean)
            Me._IsDirectMessage = value
        End Set
    End Property

    Public Property RetweetedStatus() As TwitterStatus
        Get
            Return Me._RetweetedStatus
        End Get
        Set(value As TwitterStatus)
            Me._RetweetedStatus = value
        End Set
    End Property

    Public Property Source() As String
        Get
            Return Me._Source
        End Get
        Set(value As String)
            Me._Source = value
        End Set
    End Property

    Public Property Text() As String
        Get
            Return Me._Text
        End Get
        Set(value As String)
            Me._Text = value
        End Set
    End Property

    Public Property Truncated() As Boolean
        Get
            Return Me._Truncated
        End Get
        Set(value As Boolean)
            Me._Truncated = value
        End Set
    End Property

    Public Property User() As TwitterUser
        Get
            Return Me._User
        End Get
        Set(value As TwitterUser)
            Me._User = value
        End Set
    End Property

    Public Sub New(StatusNode As XmlNode)
        Me._Text = String.Empty
        Me._InReplyToUserID = String.Empty
        Me._InReplyToScreenName = String.Empty
        Me._IsDirectMessage = False
        Me._Source = String.Empty
        Me._User = Nothing
        Me._RetweetedStatus = Nothing
        Me._GeoLat = String.Empty
        Me._GeoLong = String.Empty
        Me.CreatedAt = Me.XmlDate_Get(StatusNode("created_at"))
        Me.Favorited = Me.XmlBoolean_Get(StatusNode("favorited"))
        Me.ID = Me.XmlInt64_Get(StatusNode("id"))
        Me.InReplyToScreenName = Me.XmlString_Get(StatusNode("in_reply_to_screen_name"))
        Me.InReplyToStatusID = Me.XmlInt64_Get(StatusNode("in_reply_to_status_id"))
        Me.InReplyToUserID = Me.XmlString_Get(StatusNode("in_reply_to_user_id"))
        Me.Source = Me.XmlString_Get(StatusNode("source"))
        Me.Text = Me.XmlString_Get(StatusNode("text"))
        Me.Truncated = Me.XmlBoolean_Get(StatusNode("truncated"))
        If StatusNode("user") IsNot Nothing Then
            Me.User = New TwitterUser(StatusNode("user"))
        End If
        If StatusNode("retweeted_status") IsNot Nothing Then
            Me.RetweetedStatus = New TwitterStatus(StatusNode("retweeted_status"))
        End If
        Dim str As String = Me.XmlString_Get(StatusNode("geo"))
        If Not String.IsNullOrEmpty(str) Then
            Dim chrArray As Char() = New Char(0) {}
            chrArray(0) = " "c
            Dim strArrays As String() = str.Split(chrArray)
            If CInt(strArrays.Length) = 2 Then
                Me.GeoLat = strArrays(0)
                Me.GeoLong = strArrays(1)
            End If
        End If
    End Sub
End Class
Public Class XmlNamespaceStripper
    Public StrippedXmlDocument As XmlDocument

    Public Sub New(SourceDocument As XmlDocument)
        Dim enumerator As IEnumerator = Nothing
        Me.StrippedXmlDocument = New XmlDocument()
        Me.StrippedXmlDocument.PreserveWhitespace = True
        Using TryCast(enumerator, IDisposable)
            If Not TypeOf (enumerator) Is IDisposable Then
                enumerator = SourceDocument.ChildNodes.GetEnumerator()
                While enumerator.MoveNext()
                    Dim current As XmlNode = DirectCast(enumerator.Current, XmlNode)
                    Me.StrippedXmlDocument.AppendChild(Me.StripNamespace(current))
                End While
            End If
        End Using
    End Sub

    Private Function StripNamespace(inputNode As XmlNode) As XmlNode
        Dim enumerator As IEnumerator
        Dim enumerator1 As IEnumerator = Nothing
        Dim value As XmlNode = Me.StrippedXmlDocument.CreateNode(inputNode.NodeType, inputNode.LocalName, Nothing)
        If inputNode.Attributes IsNot Nothing Then
            Try
                enumerator = inputNode.Attributes.GetEnumerator()
                While enumerator.MoveNext()
                    Dim current As XmlAttribute = DirectCast(enumerator.Current, XmlAttribute)
                    If Operators.CompareString(current.NamespaceURI, "http://www.w3.org/2000/xmlns/", False) = 0 OrElse Operators.CompareString(current.LocalName, "xmlns", False) = 0 Then
                        Continue While
                    End If
                    Dim xmlAttribute As XmlAttribute = Me.StrippedXmlDocument.CreateAttribute(current.LocalName)
                    xmlAttribute.Value = current.Value
                    value.Attributes.Append(xmlAttribute)
                End While
            Finally
                If TypeOf (enumerator) Is IDisposable Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try
        End If
        Using TryCast(enumerator1, IDisposable)
            If Not TypeOf (enumerator) Is IDisposable Then
                enumerator1 = inputNode.ChildNodes.GetEnumerator()
                While enumerator1.MoveNext()
                    Dim xmlNodes As XmlNode = DirectCast(enumerator1.Current, XmlNode)
                    value.AppendChild(Me.StripNamespace(xmlNodes))
                End While
            End If
        End Using
        If inputNode.Value IsNot Nothing Then
            value.Value = inputNode.Value
        End If
        Return value
    End Function
End Class
Public Class TwitterUser
    Inherits XmlObjectBase
    Private _ID As Long

    Private _ScreenName As String

    Private _CreatedAt As DateTime

    Private _Description As String

    Private _FavoritesCount As Long

    Private _FriendsCount As Long

    Private _FollowersCount As Long

    Private _Location As String

    Private _Name As String

    Private _Notifications As Boolean

    Private _ProfileBackgroundColor As String

    Private _ProfileBackgroundImageUrl As String

    Private _ProfileImageUrl As String

    Private _ProfileLinkColor As String

    Private _ProfileSidebarBorderColor As String

    Private _ProfileSidebarFillColor As String

    Private _ProfileTextColor As String

    Private _Protected As Boolean

    Private _Status As TwitterStatus

    Private _StatusesCount As Long

    Private _TimeZone As String

    Private _Url As String

    Private _UTCOffset As String

    Private _Verified As Boolean

    Public Property CreatedAt() As DateTime
        Get
            Return Me._CreatedAt
        End Get
        Set(value As DateTime)
            Me._CreatedAt = value
        End Set
    End Property

    Public ReadOnly Property CreatedAtLocalTime() As DateTime
        Get
            Return Me._CreatedAt.ToLocalTime()
        End Get
    End Property

    Public Property Description() As String
        Get
            Return Me._Description
        End Get
        Set(value As String)
            Me._Description = value
        End Set
    End Property

    Public Property FavoritesCount() As Long
        Get
            Return Me._FavoritesCount
        End Get
        Set(value As Long)
            Me._FavoritesCount = value
        End Set
    End Property

    Public Property FollowersCount() As Long
        Get
            Return Me._FollowersCount
        End Get
        Set(value As Long)
            Me._FollowersCount = value
        End Set
    End Property

    Public Property FriendsCount() As Long
        Get
            Return Me._FriendsCount
        End Get
        Set(value As Long)
            Me._FriendsCount = value
        End Set
    End Property

    Public Property ID() As Long
        Get
            Return Me._ID
        End Get
        Set(value As Long)
            Me._ID = value
        End Set
    End Property

    Public Property Location() As String
        Get
            Return Me._Location
        End Get
        Set(value As String)
            Me._Location = value
        End Set
    End Property

    Public Property Name() As String
        Get
            Return Me._Name
        End Get
        Set(value As String)
            Me._Name = value
        End Set
    End Property

    Public Property Notifications() As Boolean
        Get
            Return Me._Notifications
        End Get
        Set(value As Boolean)
            Me._Notifications = value
        End Set
    End Property

    Public Property ProfileBackgroundColor() As String
        Get
            Return Me._ProfileBackgroundColor
        End Get
        Set(value As String)
            Me._ProfileBackgroundColor = value
        End Set
    End Property

    Public Property ProfileBackgroundImageUrl() As String
        Get
            Return Me._ProfileBackgroundImageUrl
        End Get
        Set(value As String)
            Me._ProfileBackgroundImageUrl = value
        End Set
    End Property

    Public Property ProfileImageUrl() As String
        Get
            Return Me._ProfileImageUrl
        End Get
        Set(value As String)
            Me._ProfileImageUrl = value
        End Set
    End Property

    Public Property ProfileLinkColor() As String
        Get
            Return Me._ProfileLinkColor
        End Get
        Set(value As String)
            Me._ProfileLinkColor = value
        End Set
    End Property

    Public Property ProfileSidebarBorderColor() As String
        Get
            Return Me._ProfileSidebarBorderColor
        End Get
        Set(value As String)
            Me._ProfileSidebarBorderColor = value
        End Set
    End Property

    Public Property ProfileSidebarFillColor() As String
        Get
            Return Me._ProfileSidebarFillColor
        End Get
        Set(value As String)
            Me._ProfileSidebarFillColor = value
        End Set
    End Property

    Public Property ProfileTextColor() As String
        Get
            Return Me._ProfileTextColor
        End Get
        Set(value As String)
            Me._ProfileTextColor = value
        End Set
    End Property

    Public Property [Protected]() As Boolean
        Get
            Return Me._Protected
        End Get
        Set(value As Boolean)
            Me._Protected = value
        End Set
    End Property

    Public Property ScreenName() As String
        Get
            Return Me._ScreenName
        End Get
        Set(value As String)
            Me._ScreenName = value
        End Set
    End Property

    Public Property Status() As TwitterStatus
        Get
            Return Me._Status
        End Get
        Set(value As TwitterStatus)
            Me._Status = value
        End Set
    End Property

    Public Property StatusesCount() As Long
        Get
            Return Me._StatusesCount
        End Get
        Set(value As Long)
            Me._StatusesCount = value
        End Set
    End Property

    Public Property TimeZone() As String
        Get
            Return Me._TimeZone
        End Get
        Set(value As String)
            Me._TimeZone = value
        End Set
    End Property

    Public Property Url() As String
        Get
            Return Me._Url
        End Get
        Set(value As String)
            Me._Url = value
        End Set
    End Property

    Public Property UTCOffset() As String
        Get
            Return Me._UTCOffset
        End Get
        Set(value As String)
            Me._UTCOffset = value
        End Set
    End Property

    Public Property Verified() As Boolean
        Get
            Return Me._Verified
        End Get
        Set(value As Boolean)
            Me._Verified = value
        End Set
    End Property

    Public Sub New(UserNode As XmlNode)
        Me._ScreenName = String.Empty
        Me._Description = String.Empty
        Me._Location = String.Empty
        Me._Name = String.Empty
        Me._ProfileBackgroundColor = String.Empty
        Me._ProfileBackgroundImageUrl = String.Empty
        Me._ProfileImageUrl = String.Empty
        Me._ProfileLinkColor = String.Empty
        Me._ProfileSidebarBorderColor = String.Empty
        Me._ProfileSidebarFillColor = String.Empty
        Me._ProfileTextColor = String.Empty
        Me._Status = Nothing
        Me._TimeZone = String.Empty
        Me._Url = String.Empty
        Me._UTCOffset = String.Empty
        Me.ID = Me.XmlInt64_Get(UserNode("id"))
        Me.ScreenName = Me.XmlString_Get(UserNode("screen_name"))
        Me.CreatedAt = Me.XmlDate_Get(UserNode("created_at"))
        Me.Description = Me.XmlString_Get(UserNode("description"))
        Me.FavoritesCount = Me.XmlInt64_Get(UserNode("favourites_count"))
        Me.FollowersCount = Me.XmlInt64_Get(UserNode("followers_count"))
        Me.FriendsCount = Me.XmlInt64_Get(UserNode("friends_count"))
        Me.Location = Me.XmlString_Get(UserNode("location"))
        Me.Name = Me.XmlString_Get(UserNode("name"))
        Me.Notifications = Me.XmlBoolean_Get(UserNode("notifications"))
        Me.ProfileBackgroundColor = Me.XmlString_Get(UserNode("profile_background_color"))
        Me.ProfileBackgroundImageUrl = Me.XmlString_Get(UserNode("profile_background_image_url"))
        Me.ProfileImageUrl = Me.XmlString_Get(UserNode("profile_image_url"))
        Me.ProfileLinkColor = Me.XmlString_Get(UserNode("profile_link_color"))
        Me.ProfileSidebarBorderColor = Me.XmlString_Get(UserNode("profile_sidebar_border_color"))
        Me.ProfileSidebarFillColor = Me.XmlString_Get(UserNode("profile_sidebar_fill_color"))
        Me.ProfileTextColor = Me.XmlString_Get(UserNode("profile_text_color"))
        Me.[Protected] = Me.XmlBoolean_Get(UserNode("protected"))
        If UserNode("status") IsNot Nothing Then
            Me.Status = New TwitterStatus(UserNode("status"))
        End If
        Me.StatusesCount = Me.XmlInt64_Get(UserNode("statuses_count"))
        Me.TimeZone = Me.XmlString_Get(UserNode("time_zone"))
        Me.Url = Me.XmlString_Get(UserNode("url"))
        Me.UTCOffset = Me.XmlString_Get(UserNode("utc_offset"))
        Me.Verified = Me.XmlBoolean_Get(UserNode("verified"))
    End Sub

    Public Sub New()
        Me._ScreenName = String.Empty
        Me._Description = String.Empty
        Me._Location = String.Empty
        Me._Name = String.Empty
        Me._ProfileBackgroundColor = String.Empty
        Me._ProfileBackgroundImageUrl = String.Empty
        Me._ProfileImageUrl = String.Empty
        Me._ProfileLinkColor = String.Empty
        Me._ProfileSidebarBorderColor = String.Empty
        Me._ProfileSidebarFillColor = String.Empty
        Me._ProfileTextColor = String.Empty
        Me._Status = Nothing
        Me._TimeZone = String.Empty
        Me._Url = String.Empty
        Me._UTCOffset = String.Empty
    End Sub
End Class
Public MustInherit Class XmlObjectBase
    Public Sub New()
    End Sub

    Protected Function XmlBoolean_Get(Node As XmlNode) As Boolean
        Dim flag As Boolean = False
        If Node IsNot Nothing Then
            Dim upper As String = Node.InnerText.ToUpper()
            If Operators.CompareString(upper, "FALSE", False) <> 0 Then
                If Operators.CompareString(upper, "TRUE", False) <> 0 Then
                    If Operators.CompareString(upper, String.Empty, False) <> 0 Then
                        Return flag
                    Else
                        Return False
                    End If
                Else
                    Return True
                End If
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    Protected Function XmlDate_Get(Node As XmlNode) As DateTime
        If Node IsNot Nothing Then
            Dim innerText As String = Node.InnerText
            Dim regex As New System.Text.RegularExpressions.Regex("(?<DayName>[^ ]+) (?<MonthName>[^ ]+) (?<Day>[^ ]{1,2}) (?<Hour>[0-9]{1,2}):(?<Minute>[0-9]{1,2}):(?<Second>[0-9]{1,2}) (?<TimeZone>[+-][0-9]{4}) (?<Year>[0-9]{4})")
            Dim match As System.Text.RegularExpressions.Match = regex.Match(innerText)
            Dim value As Object() = New Object(5) {}
            value(0) = match.Groups("MonthName").Value
            value(1) = match.Groups("Day").Value
            value(2) = match.Groups("Year").Value
            value(3) = match.Groups("Hour").Value
            value(4) = match.Groups("Minute").Value
            value(5) = match.Groups("Second").Value
            Dim dateTime__1 As DateTime = DateTime.Parse(String.Format("{0} {1} {2} {3}:{4}:{5}", value))
            Return dateTime__1
        Else
            Return DateTime.MinValue
        End If
    End Function

    Protected Function XmlInt64_Get(Node As XmlNode) As Long
        Dim num As Long = 0L
        If Node IsNot Nothing Then
            If Not Long.TryParse(Node.InnerText, num) Then
                num = CLng(0)
            End If
        Else
            num = CLng(0)
        End If
        Return num
    End Function

    Protected Function XmlString_Get(Node As XmlNode) As String
        If Node IsNot Nothing Then
            Return Node.InnerText
        Else
            Return String.Empty
        End If
    End Function
End Class



Public Enum HeaderType
    AuthHeader
    GeneralHeader
    VerifyHeader
End Enum
Public Enum Request_Types
    Auth_Request
    ShowUser_Request
    TimeLine_Request
    Token_Request
    Update_Request
    Verify_Request
End Enum
Public Enum Signature_Types
    ShowUserSignatureByScreenName
    ShowUserSignatureByID
    TimeLineSignature
    TokenRequestSignature
    UpdateSignature
    VerificationSignature
End Enum
Public Enum Method_Type
    _GET
    _POST
End Enum
Public Enum User_Lookup_Type
    _ID
    _ScreenName
    _CreatedAt
    _Description
    _FavoritesCount
    _FriendsCount
    _FollowersCount
    _Location
    _Name
    _Notifications
    _ProfileBackgroundColor
    _ProfileBackgroundImageUrl
    _ProfileImageUrl
    _ProfileLinkColor
    _ProfileSidebarBorderColor
    _ProfileSidebarFillColor
    _ProfileTextColor
    _Protected
    _Status
    _StatusesCount
    _TimeZone
    _Url
    _UTCOffset
    _Verified
End Enum
