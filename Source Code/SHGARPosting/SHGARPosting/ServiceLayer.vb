Imports System.IO
Imports System.Net
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports Newtonsoft
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class ServiceLayer


    Private userid As String = String.Empty
    Private password As String = String.Empty
    Private companydb As String = String.Empty
    Private userData As String = String.Empty
    Private url As String = String.Empty

    Private mrsFieldsCookies As CookieCollection = New CookieCollection()

    Public JSONBody As String = String.Empty
    Public SLMethod As String = String.Empty

    Private ErrorCode As String = String.Empty
    Private ErrorMessage As String = String.Empty

    Private DocKey As String = String.Empty



    Public Sub New(ByVal as_userid As String, ByVal as_password As String, ByVal as_companydb As String, ByVal as_url As String)
        Me.userid = as_userid
        Me.password = as_password
        Me.companydb = as_companydb
        Me.url = as_url

        Me.JSONBody = ""
        Me.SLMethod = ""

    End Sub

    Private Shared Function RemoteSSLTLSCertificateValidate(ByVal sender As Object, ByVal cert As X509Certificate, ByVal chain As X509Chain, ByVal ssl As SslPolicyErrors) As Boolean
        Return True
    End Function

    Public Function ConstructUserData() As String
        Return "{""UserName"" : """ + Me.userid + """, ""Password"" : """ + Me.password + """, ""CompanyDB"" : """ + Me.companydb + """}"
    End Function


    Public Function Login(ByVal as_relog As Boolean) As Dictionary(Of Boolean, String)
        If Not as_relog Then
            Return Nothing
        Else
            If UserLogged Then
                Dim logout As Dictionary(Of Boolean, String) = Me.Logout()
                If logout.ContainsKey("False") Then
                    Throw New Exception(Me.GetErrorMessage(logout("False")))
                Else
                    UserLogged = False
                End If
            End If
        End If
        ProcessAppendNew("Logging In to B1 via Service Layer. Company DB: " + SAPB1CompanyName)
        Return Me.Interact(Me.url + "/Login", "POST", True, Me.ConstructUserData())
    End Function


    Public Function Logout() As Dictionary(Of Boolean, String)

        ProcessAppendNew("Logging Out to Service Layer. Company DB: " + SAPB1CompanyName)

        Return Me.Interact(Me.url + "/Logout", "POST", False)
    End Function


    Public Function SendInvoice() As Dictionary(Of Boolean, String)
        ProcessAppendNew("Sending Invoice... " + DateTime.Now.ToString("hh:mm:ss fff"))
        Return Me.Interact(Me.url + "/Invoices", Me.SLMethod, False, Me.JSONBody)
    End Function

    Public Function SendDelivery() As Dictionary(Of Boolean, String)
        ProcessAppendNew("Sending Delivery... " + DateTime.Now.ToString("hh:mm:ss fff"))
        'ProcessAppendNew("Logging Out to Service Layer. Company DB: " + SAPB1CompanyName)
        Return Me.Interact(Me.url + "/DeliveryNotes", Me.SLMethod, False, Me.JSONBody)
    End Function

    Public Function SendDraft() As Dictionary(Of Boolean, String)
        'ProcessAppendNew("Logging Out to Service Layer. Company DB: " + SAPB1CompanyName)
        Return Me.Interact(Me.url + "/Drafts", Me.SLMethod, False, Me.JSONBody)
    End Function

    Public Function SendTender() As Dictionary(Of Boolean, String)
        'ProcessAppendNew("Logging Out to Service Layer. Company DB: " + SAPB1CompanyName)
        Return Me.Interact(Me.url + "/IncomingPayments", Me.SLMethod, False, Me.JSONBody)
    End Function

    Public Function SendCM() As Dictionary(Of Boolean, String)
        ' ProcessAppendNew("Sending Outgoing Payment To Service Layer. Company DB: " + SAPB1CompanyName)
        Return Me.Interact(Me.url + "/CreditNotes", Me.SLMethod, False, Me.JSONBody)
    End Function

    Public Function SendOutgoing() As Dictionary(Of Boolean, String)
        'ProcessAppendNew("Sending Outgoing Payment to Service Layer. Company DB: " + SAPB1CompanyName)
        Return Me.Interact(Me.url + "/VendorPayments", Me.SLMethod, False, Me.JSONBody)
    End Function

    Public Function UpdateCM(ByVal as_docentry As String) As Dictionary(Of Boolean, String)
        'ProcessAppendNew("Updating ARCM " + as_docentry + ". Company DB: " + SAPB1CompanyName)
        Return Me.Interact(Me.url + "/CreditNotes(" + as_docentry + ")", Me.SLMethod, False, Me.JSONBody)
    End Function


    Private Function Interact(ByVal as_url As String, ByVal as_method As String, ByVal as_saveCookie As Boolean, ByVal Optional as_body As String = "") As Dictionary(Of Boolean, String)
        Dim actionReturn As Boolean = True
        Dim result = New Dictionary(Of Boolean, String)()

        Dim httpResponse As HttpWebResponse = Nothing

        Dim encoding As New ASCIIEncoding()
        'Dim byte1 As Byte() = encoding.GetBytes(jsondata)

        Dim success As Boolean = False
        'Declare a web request based on the parameter sent
        Dim request As HttpWebRequest = DirectCast(WebRequest.Create(New Uri(as_url)), HttpWebRequest)




        'Method is POST based on the discussion.
        request.Method = as_method.ToUpper()
        'Content Type must be json char set utf-8 is for special characters only.
        'request.ContentType = "application/json; charset=utf-8"
        'Just setting the length
        'request.ContentLength = byte1.Length

        'request.Timeout = 20000 ' //A request that didn't get respond within 20 seconds is unacceptable, and we would rather just retry.
        request.KeepAlive = False
        'request.ProtocolVersion = HttpVersion.Version10
        request.ServicePoint.Expect100Continue = False
        ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf RemoteSSLTLSCertificateValidate)
        request.CookieContainer = New CookieContainer()


        request.ContentType = "application/json; odata=minimalmetadata; charset=utf8"
        request.Timeout = 10000000

        If mrsFieldsCookies.Count > 0 Then
            For Each nibblers As Cookie In mrsFieldsCookies
                request.CookieContainer.Add(New Uri(as_url), New Cookie(nibblers.Name, nibblers.Value))
            Next
        End If

        If Not String.IsNullOrEmpty(as_body) Then
            Dim byte1 As Byte() = encoding.GetBytes(as_body)
            request.ContentLength = byte1.Length
            'ProcessAppendNew("Content:" + as_body)
            Try
                Dim requestStream As Stream = request.GetRequestStream()
                requestStream.Write(byte1, 0, byte1.Length)
                requestStream.Close()
            Catch ex As Exception
                actionReturn = False
                result.Add(False, ex.Message())
            End Try
        End If


        If actionReturn Then
            Try
                Try
                    httpResponse = DirectCast(request.GetResponse(), HttpWebResponse)
                Catch wex As WebException
                    httpResponse = wex.Response
                    actionReturn = False
                End Try
                Using streamReader = New StreamReader(httpResponse.GetResponseStream())

                    Dim response As String = streamReader.ReadToEnd.ToString()
                    If response.Contains("error") Then
                        'ErrorAppend(response)
                        actionReturn = False
                    End If

                    result.Add(actionReturn, response)
                    'ProcessAppendNew("Response: " + response)
                End Using

                If as_saveCookie Then
                    mrsFieldsCookies = httpResponse.Cookies
                End If

                Dim doctype As String = ""
                If as_url.Contains("DeliveryNotes") Then doctype = "Delivery"
                If as_url.Contains("Invoice") Then doctype = "Invoice"

                If doctype <> "" Then
                    ProcessAppendNew("Sending " + doctype + " ends... " + DateTime.Now.ToString("hh:mm:ss fff"))
                End If
            Catch ex As Exception
                result.Add(False, ex.Message)
            End Try
        End If


        Me.JSONBody = ""
        Me.SLMethod = ""
        Return result
    End Function


    Public Function PostMethod() As String
        Return "POST"
    End Function


    Public Function GetMethod() As String
        Return "GET"
    End Function

    Public Function PatchMethod() As String
        Return "PATCH"
    End Function

    Public Function DeleteMethod() As String
        Return "DELETE"
    End Function

    Public Function GetErrorMessage() As String
        Return Me.ErrorMessage
    End Function

    Public Function GetErrorCode() As String
        Return Me.ErrorCode
    End Function

    Public Function GetDocKey() As String
        Return Me.DocKey
    End Function


    Public Function OperationSuccess(ByVal as_result As Dictionary(Of Boolean, String)) As Boolean

        Dim jsonstring As String = ""

        Me.ErrorCode = ""
        Me.ErrorMessage = ""
        Me.DocKey = ""
        Dim jObj As JObject

        Try
            If as_result Is Nothing Then
                Return True
            End If


            If as_result.ContainsKey("False") Then
                jsonstring = as_result("False")
                'ErrorAppendNew("false: " + jsonstring)
                jObj = JObject.Parse(jsonstring)
                Me.ErrorCode = jObj("error")("code").ToString()
                Me.ErrorMessage = jObj("error")("message")("value").ToString()
                Return False
            Else


                jsonstring = as_result("True")
                'MessageBox.Show(jsonstring)
                'ErrorAppendNew("true: " + jsonstring)
                If jsonstring <> "" Then

                    Try
                        jObj = JObject.Parse(jsonstring)
                        If jsonstring.Contains("DocEntry") Then
                            Me.DocKey = jObj("DocEntry").ToString()
                        Else
                            If jsonstring.Contains("DocNum") Then
                                Me.DocKey = jObj("DocNum").ToString()
                            End If
                        End If
                    Catch ex As Exception
                        Throw New Exception(jsonstring)
                    End Try
                End If
            End If
        Catch ex As Exception
            Me.ErrorCode = "-999"
            Me.ErrorMessage = "Exception Msg: " + ex.Message + " Response Msg:" + jsonstring
            Return False
        End Try
        Return True
    End Function
End Class
