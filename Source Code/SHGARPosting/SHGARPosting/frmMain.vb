Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.IO
Imports System.IO.DirectoryInfo
Imports System.IO.Directory


Imports System.Globalization

Public Class frmMain

    Private oDataSales, oDataSalesKeys, oDataDiscounts, oDataDelivery, oDataFreight, oDataSL, oDataDD As DataTable

    Dim batchStart, batchEnd, processDocStart, processDocEnd As String

    Dim relogServiceLayer = False
    Dim sLayerLogged As Boolean = False

    Dim SLayer As ServiceLayer
    Dim SLayerRslt As Dictionary(Of Boolean, String)


    Dim invoicestarts, deliverystarts As String

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Control.CheckForIllegalCrossThreadCalls = False
        SAPB1Company = New SAPbobsCOM.Company

        ProgramName = "FTSHGARPosting"
        Me.Text = "FastTrack AR Posting"
        Today = DateTime.Now.ToString("yyyy-MM-dd")

        batchStart = ""
        batchEnd = ""


        processDocStart = ""
        processDocEnd = ""

        relogServiceLayer = True

        lblStatus.Text = ""
        lblToday.Text = DateTime.Now.ToString("MMMM dd, yyyy hh:mm:ss tt")

        ReConnect = True

        initDataTable()

        timerToday.Start()
        timerStarter.Start()
        bgwInitialize.RunWorkerAsync()

    End Sub


    Sub initDataTable()

        Try

            oDataSL = New DataTable("SL")
            With oDataSL.Columns

            End With


            oDataSales = New DataTable("Sales")
            With oDataSales.Columns
                .Add("Transation_date", GetType(System.String))
                .Add("Store_Code", GetType(System.String))
                .Add("Line_ID", GetType(System.String))
                .Add("Sales_Type", GetType(System.String))
                .Add("Item_Code", GetType(System.String))
                .Add("Item_Description", GetType(System.String))
                .Add("Sales_Quantity", GetType(System.String))
                .Add("Price", GetType(System.String))
                .Add("Sales_Amount", GetType(System.String))
                .Add("Tax_Code", GetType(System.String))
                .Add("Tax_Amount", GetType(System.String))
                .Add("Discount_Quantity", GetType(System.String))
                .Add("Discount_Amount", GetType(System.String))
                .Add("Total_Guest", GetType(System.String))
                .Add("Service_charge", GetType(System.String))
                .Add("Item_Group", GetType(System.String))
                .Add("IsBatch", GetType(System.String))
                .Add("CostingCode3", GetType(System.String))
            End With

            oDataSalesKeys = New DataTable("DataKey")
            oDataSalesKeys.Columns.Add("Key", GetType(System.String))
            oDataSalesKeys.Columns.Add("FilePath", GetType(System.String))

            oDataDelivery = New DataTable("Delivery")
            With oDataDelivery.Columns
                .Add("Transation_date", GetType(System.String))
                .Add("Store_Code", GetType(System.String))
                .Add("Discount_Type", GetType(System.String))
                .Add("Document", GetType(System.String))
            End With

            oDataDiscounts = New DataTable("Discounts")
            With oDataDiscounts.Columns
                .Add("Transation_date", GetType(System.String))
                .Add("Store_Code", GetType(System.String))
                .Add("Line_ID", GetType(System.String))
                .Add("Base_Line_ID", GetType(System.String))
                .Add("Item_Code", GetType(System.String))
                .Add("Item_Description", GetType(System.String))
                .Add("Discount_Type", GetType(System.String))
                .Add("Discount_Quantity", GetType(System.String))
                .Add("Discount_Amount", GetType(System.String))
                .Add("Employee_Code", GetType(System.String))
                .Add("Cost_Code", GetType(System.String))
                .Add("Document", GetType(System.String))
                .Add("Freight", GetType(System.String))
                .Add("Item_Group", GetType(System.String))
                .Add("IsBatch", GetType(System.String))
                .Add("CostingCode3", GetType(System.String))
            End With

            oDataFreight = New DataTable("Freight")
            With oDataFreight.Columns
                .Add("Code", GetType(System.String))
                .Add("Amount", GetType(System.String))
                .Add("VatGroup", GetType(System.String))
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message + " Application will exit!")
            ExitProgram()
        End Try

    End Sub

    Private Sub frmMain_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        If MessageBox.Show("Closing this application will stop all AR being processed. Do you want to continue?", ProgramName, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            e.Cancel = False
        Else
            'If OnProcess = True Then
            '    MessageBox.Show("A document is being processed. Application will close after processing!")
            'End If
        End If
    End Sub

    Private Sub bgwInitialize_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwInitialize.DoWork
        Try

            lblStatus.Text = "Checking existing instance..."
            CheckForExistingInstance()
            Threading.Thread.Sleep(1000)

            lblStatus.Text = "Reading configuration file..."
            ReadConfig()
            Threading.Thread.Sleep(1000)


            lblStatus.Text = "Creating system folders..."
            FileWrite()
            Threading.Thread.Sleep(1000)

            lblStatus.Text = "Connecting to " + SAPB1CompanyName + " via DI API..."
            B1Connect(SAPB1CompanyName)
            Threading.Thread.Sleep(1000)

            lblStatus.Text = "Connecting to " + SAPB1CompanyName + " via Service Layer..."
            Threading.Thread.Sleep(1000)
            If relogServiceLayer Then
                Dim bool As Boolean = SLayer.OperationSuccess(SLayer.Login(relogServiceLayer))

                'SuccessAppendNew(bool.ToString())

                If Not bool Then
                    ErrMsg = "Service layer can't logged in ! Error:" + SLayer.GetErrorCode() + " " + SLayer.GetErrorMessage() + ". Application will close."
                    Throw New Exception(ErrMsg)
                Else
                    relogServiceLayer = False
                    UserLogged = True
                    'timerLoginSession.Start()
                    lblStatus.Text = "Successfully logged via Service Layer..."
                    ProcessAppendNew(lblStatus.Text)

                End If

            End If

            Threading.Thread.Sleep(1000)

            lblStatus.Text = "Initializing tables for " + SAPB1CompanyName + "..."
            InitializeB1Tables()
            Threading.Thread.Sleep(1000)

            lblStatus.Text = "Initializing fields for " + SAPB1CompanyName + "..."
            InitializeB1Fields()
            Threading.Thread.Sleep(1000)

            lblStatus.Text = "Loading default setups from " + SAPB1CompanyName + "..."
            B1Setups()
            Threading.Thread.Sleep(1000)

            lblStatus.Text = "Starting memory optimizer..."
            MemoryOptimizeStart()
            Threading.Thread.Sleep(1000)



            lblStatus.Text = ""

        Catch ex As Exception
            MessageBox.Show(ex.Message, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ExitProgram()
        End Try
    End Sub
    Private Sub ReadConfig()
        Try
            Dim stringline As String = ""
            Dim sr As StreamReader = Nothing
            FilePath = System.Windows.Forms.Application.StartupPath + "\\Settings.ini"

            Try
                sr = New StreamReader(FilePath)
            Catch ex As Exception
                Throw New Exception("Error reading config! " + ex.Message)
            End Try

            stringline = sr.ReadLine
            ServerName = stringline.Substring(stringline.IndexOf("=") + 1)

            stringline = sr.ReadLine
            SQLServerType = stringline.Substring(stringline.IndexOf("=") + 1)


            stringline = sr.ReadLine
            SAPB1UseTrusted = Convert.ToBoolean(stringline.Substring(stringline.IndexOf("=") + 1))

            stringline = sr.ReadLine
            SQLUser = stringline.Substring(stringline.IndexOf("=") + 1)

            stringline = sr.ReadLine
            SQLPassword = stringline.Substring(stringline.IndexOf("=") + 1)

            stringline = sr.ReadLine
            SAPB1CompanyName = stringline.Substring(stringline.IndexOf("=") + 1)

            stringline = sr.ReadLine
            UserId = stringline.Substring(stringline.IndexOf("=") + 1)

            stringline = sr.ReadLine
            UserPassword = stringline.Substring(stringline.IndexOf("=") + 1)

            stringline = sr.ReadLine
            FileOriginDump = stringline.Substring(stringline.IndexOf("=") + 1)

            If Not Directory.Exists(FileOriginDump) Then
                Throw New Exception(FileOriginDump + " does not exists!")
            End If

            stringline = sr.ReadLine
            ProcessTime = stringline.Substring(stringline.IndexOf("=") + 1)
            If ProcessTime <> "0" Then
                If ProcessTime.Length <> "4" Then
                    Throw New Exception("Invalid Process Time!")
                End If
            End If

            stringline = sr.ReadLine
            SAPB1UserIdtoMessage = stringline.Substring(stringline.IndexOf("=") + 1)


            stringline = sr.ReadLine
            SLayerUrl = stringline.Substring(stringline.IndexOf("=") + 1)

            SLayer = New ServiceLayer(UserId, UserPassword, SAPB1CompanyName, SLayerUrl)

            If SAPB1CompanyName.Trim <> "" Then
                Me.Text += String.Format("({0})", SAPB1CompanyName.ToUpper())
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Function InitializeB1Tables() As String
        Try
            B1Connect(SAPB1CompanyName)

            If isUDTExist("FT_DISCTYPE") = False Then
                If Not createUDT("FT_DISCTYPE", "Discount Type Setup", SAPbobsCOM.BoUTBTableType.bott_NoObject) Then
                    Throw New Exception(ErrMsg)
                End If
            End If



            If isUDTExist("FT_SALETYPE") = False Then
                If Not createUDT("FT_SALETYPE", "Sales Type Setup", SAPbobsCOM.BoUTBTableType.bott_NoObject) Then
                    Throw New Exception(ErrMsg)
                End If
            End If

            If isUDTExist("POSEXITEM") = False Then
                If Not createUDT("POSEXITEM", "POSEXITEM", SAPbobsCOM.BoUTBTableType.bott_NoObject) Then
                    Throw New Exception(ErrMsg)
                End If
            End If

            If isUDTExist("POSITEM") = False Then
                If Not createUDT("POSITEM", "POSITEM", SAPbobsCOM.BoUTBTableType.bott_NoObject) Then
                    Throw New Exception(ErrMsg)
                End If
            End If

            If isUDTExist("WHSEMAPPING") = False Then
                If Not createUDT("WHSEMAPPING", "WHSEMAPPING", SAPbobsCOM.BoUTBTableType.bott_NoObject) Then
                    Throw New Exception(ErrMsg)
                End If
            End If

            'Change Request 04202022

            If isUDTExist("FT_CIBGL") = False Then
                If Not createUDT("FT_CIBGL", "Cash In-Bank GL Account", SAPbobsCOM.BoUTBTableType.bott_NoObject) Then
                    Throw New Exception(ErrMsg)
                End If
            End If

        Catch ex As Exception
            ErrorAppend(ex.Message)
            Throw New Exception(ex.Message)
        End Try
        Return ""
    End Function
    Private Function B1Setups() As Boolean
        Dim actionReturn As Boolean = True

        Try
            B1Connect(SAPB1CompanyName)



        Catch ex As Exception
            ErrorAppend(ex.Message)
            Throw New Exception(ex.Message)
        End Try
        Return actionReturn
    End Function
    Private Function InitializeB1Fields() As String

        Try
            B1Connect(SAPB1CompanyName)

            'OINV
            If isUDFExist("OINV", "FileName") = False Then
                If Not createUDF("OINV", "FileName", "File Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If
            If isUDFExist("INV1", "SalesTyp") = False Then
                If Not createUDF("INV1", "SalesTyp", "Sales Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 2, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If
            'Item Master
            If isUDFExist("OITM", "MCItem") = False Then
                If Not createUDF("OITM", "MCItem", "MC ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If
            If isUDFExist("OITM", "CWItem") = False Then
                If Not createUDF("OITM", "CWItem", "CW ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If

            'Business Partners
            If isUDFExist("OCRD", "MCEmp") = False Then
                If Not createUDF("OCRD", "MCEmp", "MC Employee", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If
            If isUDFExist("OCRD", "CWEmp") = False Then
                If Not createUDF("OCRD", "CWEmp", "CW Employee", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If
            If isUDFExist("OCRD", "CostCode") = False Then
                If Not createUDF("OCRD", "CostCode", "Customer Cost Center", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If
            If isUDFExist("OCRD", "StoreID") = False Then
                If Not createUDF("OCRD", "StoreID", "Store Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If

            If isUDFExist("OCRD", "Store") = False Then
                If Not createUDF("OCRD", "Store", "Store", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If

            If isUDFExist("OCRD", "WhsCode") = False Then
                If Not createUDF("OCRD", "WhsCode", "Warehouse Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 8, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If
            'Discount Type
            If isUDFExist("@FT_DISCTYPE", "DiscCode") = False Then
                If Not createUDF("@FT_DISCTYPE", "DiscCode", "POS Discount Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If
            If isUDFExist("@FT_DISCTYPE", "DiscName") = False Then
                If Not createUDF("@FT_DISCTYPE", "DiscName", "POS Discount Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If
            If isUDFExist("@FT_DISCTYPE", "Document") = False Then
                If Not createUDF("@FT_DISCTYPE", "Document", "POS Discount Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 4, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If
            If isUDFExist("@FT_DISCTYPE", "CostCode") = False Then
                If Not createUDF("@FT_DISCTYPE", "CostCode", "Cogs Costing Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If
            If isUDFExist("@FT_DISCTYPE", "Freight") = False Then
                If Not createUDF("@FT_DISCTYPE", "Freight", "Freight Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If
            If isUDFExist("@FT_DISCTYPE", "Type") = False Then
                If Not createUDF("@FT_DISCTYPE", "Type", "Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If

            If isUDFExist("@FT_SALETYPE", "VatCode") = False Then
                If Not createUDF("@FT_SALETYPE", "VatCode", "Vat Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If

            If isUDFExist("OITB", "Type") = False Then
                If Not createUDF("OITB", "Type", "Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If

            If isUDFExist("OVTG", "VatGroup") = False Then
                If Not createUDF("OVTG", "VatGroup", "Vat Group", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If


            If isUDFExist("@POSITEM", "POS") = False Then
                If Not createUDF("@POSITEM", "POS", "Point of Sale", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If

            If isUDFExist("@POSITEM", "ItemCode") = False Then
                If Not createUDF("@POSITEM", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If


            If isUDFExist("@WHSEMAPPING", "Store") = False Then
                If Not createUDF("@WHSEMAPPING", "Store", "Store", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If

            If isUDFExist("@WHSEMAPPING", "ItemGroup") = False Then
                If Not createUDF("@WHSEMAPPING", "ItemGroup", "ItemGroup", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If
            If isUDFExist("@WHSEMAPPING", "Whse") = False Then
                If Not createUDF("@WHSEMAPPING", "Whse", "Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If

            'PAYMENT AND CHANGE REQUEST
            If isUDFExist("OPRC", "U_Dim1") = False Then
                If Not createUDF("OPRC", "U_Dim1", "Brand Cost Center", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If
            If isUDFExist("OCRC", "POSPay") = False Then
                If Not createUDF("OCRC", "POSPay", "POS Payment Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If

            '04202022

            If isUDFExist("@FT_CIBGL", "AcctCode") = False Then
                If Not createUDF("@FT_CIBGL", "AcctCode", "G/L Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If

            If isUDFExist("@FT_CIBGL", "AcctName") = False Then
                If Not createUDF("@FT_CIBGL", "AcctName", "G/L Account Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "", "") Then
                    Throw New Exception(ErrMsg)
                End If
            End If


        Catch ex As Exception
            ErrorAppend(ex.Message)
            Throw New Exception(ex.Message)
        End Try

        Return ""
    End Function
    Private Function GetB1DataType(ByVal as_type As String) As Long
        Dim ltReturnValue As Long = Nothing
        Select Case as_type
            Case "ALPHA"
                ltReturnValue = SAPbobsCOM.BoFieldTypes.db_Alpha
            Case "DATE"
                ltReturnValue = SAPbobsCOM.BoFieldTypes.db_Date
            Case "NUMERIC"
                ltReturnValue = SAPbobsCOM.BoFieldTypes.db_Numeric
            Case "FLOAT"
                ltReturnValue = SAPbobsCOM.BoFieldTypes.db_Float
            Case "PRICE"
                ltReturnValue = SAPbobsCOM.BoFldSubTypes.st_Price
            Case "QUANTITY"
                ltReturnValue = SAPbobsCOM.BoFldSubTypes.st_Quantity
            Case "PERCENTAGE"
                ltReturnValue = SAPbobsCOM.BoFldSubTypes.st_Percentage
        End Select


        Return ltReturnValue
    End Function
    Private Sub bgwIntegrate_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwIntegrate.DoWork
        If DateTime.Now.ToString("yyyy-MM-dd") <> Today Then
            FileWrite()
            Today = DateTime.Now.ToString("yyyy-MM-dd")
        End If
        lblStatus.Text = "Waiting for file to process..."
        ProcessStart()
        Threading.Thread.Sleep(3000)
    End Sub

    Private Sub ProcessStart()
        Dim actionReturn As Boolean = False
        Try



            If ProcessTime <> "0" Then
                Dim testtime As String = DateTime.Now.ToString("HHmm")
                If ProcessTime = testtime Then
                    actionReturn = True
                End If
            Else
                actionReturn = True
            End If

            If actionReturn Then

                If relogServiceLayer Then
                    Dim bool As Boolean = SLayer.OperationSuccess(SLayer.Login(relogServiceLayer))

                    'SuccessAppendNew(bool.ToString())

                    If Not bool Then
                        ErrMsg = "Service layer can't logged in ! Error:" + SLayer.GetErrorCode() + " " + SLayer.GetErrorMessage() + ". Application will close."
                        ErrorAppendNew(ErrMsg)
                        ExitProgram()
                    Else
                        relogServiceLayer = False
                        UserLogged = True
                        'timerLoginSession.Start()
                        ProcessAppendNew("Successfully logged via Service Layer")
                    End If

                End If


                If Not ReadTextFiles() Then
                    Throw New Exception(ErrMsg)
                End If

                If Not ProcessTransferedTextFiles() Then
                    Throw New Exception(ErrMsg)
                End If

                ProcessPaymentMethod()

            End If
            GC.Collect()
        Catch ex As Exception
            ErrorAppend("FATAL ERROR ! " + ex.Message + " Application will exit!")

            'If SAPB1UserIdtoMessage <> "" Then
            '    If Not SendAlert(ex.Message) Then
            '        ErrorAppend("SENDING ALERT EXCEPTION!! " + ErrMsg + ".")
            '    End If
            'End If


            ExitProgram()
        End Try

    End Sub

    Private Sub ProcessPaymentMethod()

        Dim strLine, strCardCode, strFileName, strQuery, strCIBGL,
            strPayType, strDocEntry, strObjType, strVoucher,
            strTransFile(), strValue() As String

        Dim intInvCnt, intCCCount As Integer
        Dim decAmount As Decimal

        Dim dteDocDate As DateTime

        Dim oPayments As SAPbobsCOM.Payments
        Dim oRecordset, oCIBGLRS As SAPbobsCOM.Recordset

        Dim blWithErr As Boolean

        strFileName = ""

        Try

            For Each strFile In Directory.GetFiles(FileDump, "PM*.txt")

                intCCCount = 0
                blWithErr = False
                strFileName = Path.GetFileName(strFile)
                strTransFile = strFileName.Split("_")

                strVoucher = strFileName.Substring(0, 20).Replace("PM_CW_", "")

                strCardCode = strTransFile(2)
                dteDocDate = System.Convert.ToDateTime(strTransFile(3).Replace(".txt", "").Substring(0, 2) + "/" + strTransFile(3).Replace(".txt", "").Substring(2, 2) + "/" + strTransFile(3).Replace(".txt", "").Substring(4, 2))

                If Not SAPB1Company.InTransaction() Then
                    SAPB1Company.StartTransaction()
                End If

                oPayments = Nothing
                oPayments = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)

                oPayments.DocDate = dteDocDate
                oPayments.CardCode = strCardCode

                strQuery = String.Format("SELECT ""DocEntry"", ""ObjType"" FROM ""OINV"" WHERE ""CardCode"" = '{0}'  AND  ""DocStatus"" = 'O' AND " +
                                         """DocDate"" = " + " to_date('{1}', 'MM/DD/YYYY')  ", strCardCode, dteDocDate.ToShortDateString())

                oRecordset = Nothing
                oRecordset = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordset.DoQuery(strQuery)

                If Not oRecordset.RecordCount > 0 Then

                    SetLabelText("Error Processing File [" & strFileName & "]. Please check Error Log.")
                    ErrorAppend("Sales Transaction not Found - FileName [" & strFileName & "].")

                    If File.Exists(ErrorFileDumpPath & "\" & strFileName) Then
                        File.Delete(ErrorFileDumpPath & "\" & strFileName)
                    End If

                    If SAPB1Company.InTransaction() Then
                        SAPB1Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    End If

                    File.Move(strFile, ErrorFileDumpPath & "\" & strFileName)

                    Continue For

                Else

                    intInvCnt = 0

                    While Not oRecordset.EoF

                        strDocEntry = oRecordset.Fields.Item("DocEntry").Value.ToString()
                        strObjType = oRecordset.Fields.Item("ObjType").Value.ToString()

                        If intInvCnt > 0 Then
                            oPayments.Invoices.Add()
                        End If

                        oPayments.Invoices.DocEntry = strDocEntry
                        oPayments.Invoices.InvoiceType = strObjType


                        intInvCnt += 1
                        oRecordset.MoveNext()

                    End While

                End If

                Dim sr = New StreamReader(strFile)

                Do While sr.Peek() <> -1

                    strLine = sr.ReadLine
                    strValue = strLine.Split(vbTab)

                    strPayType = strValue(2)
                    decAmount = Convert.ToDecimal(strValue(3))

                    If strPayType = "CASH" Or strPayType = "Cash" Or strPayType = "cash" Then

                        strQuery = String.Format("SELECT OACT.""AcctCode"" " +
                                                 "From OACT INNER JOIN ""@FT_CIBGL"" CIBGL ON OACT.""FormatCode"" = REPLACE(CIBGL.""U_AcctCode"", '-', '') " +
                                                 "WHERE CIBGL.""Code"" = '{0}' ", strCardCode)

                        oCIBGLRS = Nothing
                        oCIBGLRS = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oCIBGLRS.DoQuery(strQuery)

                        If oCIBGLRS.RecordCount > 0 Then
                            oPayments.CashAccount = oCIBGLRS.Fields.Item("AcctCode").Value.ToString()
                        End If

                        oPayments.CashSum = decAmount

                        Else

                        strQuery = String.Format("SELECT OCRC.""AcctCode"", OCRC.""CreditCard"", OCRP.""CrTypeCode"" " +
                                                "FROM OCRC LEFT JOIN OCRP ON OCRC.""CreditCard"" = OCRP.""CreditCard"" " +
                                                "WHERE ""U_POSPay"" = '{0}' ", strPayType)


                        oRecordset = Nothing
                        oRecordset = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordset.DoQuery(strQuery)

                        If Not oRecordset.RecordCount > 0 Then

                            blWithErr = True

                            SetLabelText("Error Processing File [" & strFileName & "]. Please check Error Log.")
                            ErrorAppend("POS Payment Type not Found - FileName [" & strFileName & "].")

                            Exit Do

                        Else

                            If intCCCount > 0 Then
                                oPayments.CreditCards.Add()
                            End If

                            oPayments.CreditCards.CreditAcct = oRecordset.Fields.Item("AcctCode").Value.ToString()
                            oPayments.CreditCards.CreditCard = oRecordset.Fields.Item("CreditCard").Value.ToString()
                            oPayments.CreditCards.PaymentMethodCode = oRecordset.Fields.Item("CrTypeCode").Value.ToString()
                            oPayments.CreditCards.CreditCardNumber = "1234"
                            oPayments.CreditCards.CardValidUntil = "12.31.2099"
                            oPayments.CreditCards.VoucherNum = strVoucher
                            oPayments.CreditCards.CreditSum = decAmount

                            intCCCount += 1

                        End If

                    End If

                Loop

                sr.Close()

                If blWithErr = False Then

                    If oPayments.Add <> 0 Then

                        ErrorAppend("Error adding Payment Method - FileName [" & strFileName & "]. " & SAPB1Company.GetLastErrorCode.ToString() & " - " & SAPB1Company.GetLastErrorDescription)
                        SetLabelText("Error Processing File [" & strFileName & "]. Please check Error Log.")

                        If File.Exists(ErrorFileDumpPath & "\" & strFileName) Then
                            File.Delete(ErrorFileDumpPath & "\" & strFileName)
                        End If

                        If SAPB1Company.InTransaction() Then
                            SAPB1Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If

                        File.Move(strFile, ErrorFileDumpPath & "\" & strFileName)

                    Else

                        Try

                            If SAPB1Company.InTransaction() Then
                                SAPB1Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                            End If

                            SetLabelText("File [" & strFileName & "] Processed Successfully")
                            SuccessAppend("File [" & strFileName & "] Processed Successfully")

                            If File.Exists(SuccessFileDumpPath & "\" & strFileName) Then
                                File.Delete(SuccessFileDumpPath & "\" & strFileName)
                            End If

                            File.Move(strFile, SuccessFileDumpPath & "\" & strFileName)

                        Catch ex As Exception

                            If SAPB1Company.InTransaction() Then
                                SAPB1Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            End If

                            ErrorAppend("Error adding Payment Method - FileName [" & strFileName & "]. " & SAPB1Company.GetLastErrorCode.ToString() & " - " & SAPB1Company.GetLastErrorDescription)
                            SetLabelText("Error Processing File [" & strFileName & "]. Please check Error Log.")

                            If File.Exists(ErrorFileDumpPath & "\" & strFileName) Then
                                File.Delete(ErrorFileDumpPath & "\" & strFileName)
                            End If

                            File.Move(strFile, ErrorFileDumpPath & "\" & strFileName)

                        End Try

                    End If

                Else

                    If SAPB1Company.InTransaction() Then
                        SAPB1Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    End If

                    If File.Exists(ErrorFileDumpPath & "\" & strFileName) Then
                        File.Delete(ErrorFileDumpPath & "\" & strFileName)
                    End If

                    File.Move(strFile, ErrorFileDumpPath & "\" & strFileName)

                End If

            Next

        Catch ex As Exception

            ErrorAppend("Error adding Payment Method - FileName [" & strFileName & "]. " & ex.Message.ToString())
            SetLabelText("Error Processing File [" & strFileName & "] " & ex.Message.ToString())

        End Try

    End Sub
    Private Function ProcessTransferedTextFiles() As Boolean
        Try
            Dim cofPath As New DirectoryInfo(FileDump)
            Dim textFiles As FileInfo() = cofPath.GetFiles("SL*.txt")
            Dim strfile, strFileName, strFileDumpDest As String
            Dim iFileTransferedCount As Integer = 0
            Dim lineadd As Boolean = False
            Dim actionReturn As Boolean = True

            Dim oRS As SAPbobsCOM.Recordset

            Dim codecheck As String = ""

            Dim ls_sales_file, ls_discount_file, ls_query As String
            Dim lb_Error As Boolean
            lb_Error = False
            Dim dir As DirectoryInfo
            Dim fi_sales, fi_Disc As FileInfo()
            Dim Fi_sl, Fi_dd As FileInfo

            Dim invoiceprocess, deliveryprocess, invnow, delnow As Boolean



            dir = New DirectoryInfo(FileDump)
            fi_sales = dir.GetFiles("SL*.txt")

            B1Connect(SAPB1CompanyName)

            If fi_sales.Length > 0 Then
                ProcessAppendNew("Processing files starts " + DateTime.Now.ToString("hh:mm:ss fff"))

                For Each Fi_sl In fi_sales
                    ls_sales_file = Fi_sl.FullName
                    ls_discount_file = "DD" & Fi_sl.Name.Substring(2)
                    fi_Disc = dir.GetFiles("DD" & Fi_sl.Name.Substring(2))


                    'If fi_Disc.Length > 0 Then

                    Fi_dd = Nothing
                    For Each Fi_dd In fi_Disc
                        ls_discount_file = Fi_dd.FullName
                    Next Fi_dd
                    oDataFreight.Rows.Clear()
                    oDataDelivery.Clear()
                    oDataDiscounts.Clear()
                    oDataSales.Clear()

                    invoiceprocess = False
                    deliveryprocess = False
                    invnow = False
                    delnow = False

                    ls_query = "SELECT A.""DocEntry"" AS ""DocEntry"", A.""TYPE"" FROM ( SELECT ""DocEntry"",""U_FileName"",""CANCELED"", 'DLRY' AS ""TYPE"" FROM ODLN WHERE IFNULL(""U_FileName"", '') = '" + Fi_dd.Name + "' UNION ALL SELECT ""DocEntry"",""U_FileName"",""CANCELED"", 'ARIN' AS ""TYPE"" FROM OINV WHERE IFNULL(""U_FileName"", '') = '" + Fi_sl.Name + "' ) A WHERE A.""CANCELED"" = 'N'"

                    oRS = Nothing
                    oRS = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRS.DoQuery(ls_query)

                    If oRS.RecordCount > 0 Then
                        'ErrorAppend("FileName [" & Fi_sl.Name & "] already uploaded to DocEntry [" & oRS.Fields.Item("DocEntry").Value & "]")
                        'lb_Error = True
                        'If File.Exists(ErrorFileDumpPath & "\" & Fi_sl.Name) Then File.Delete(ErrorFileDumpPath & "\" & Fi_sl.Name)
                        'Fi_sl.MoveTo(ErrorFileDumpPath & "\" & Fi_sl.Name)
                        'If File.Exists(ErrorFileDumpPath & "\" & Fi_dd.Name) Then File.Delete(ErrorFileDumpPath & "\" & Fi_dd.Name)
                        'Fi_dd.MoveTo(ErrorFileDumpPath & "\" & Fi_dd.Name)
                        'GoTo movenext

                        oRS.MoveFirst()

                        For i As Integer = 0 To oRS.RecordCount - 1
                            If oRS.Fields.Item("TYPE").Value = "DLRY" Then
                                deliveryprocess = True
                            End If
                            If oRS.Fields.Item("TYPE").Value = "ARIN" Then
                                invoiceprocess = True
                            End If
                            oRS.MoveNext()
                        Next i


                    End If
                    oRS = Nothing
                    GC.Collect()

                    SetLabelText("Loading [" & Fi_dd.Name & "]")
                    If Not InsertToDiscount(ls_discount_file, Fi_dd.Name) Then
                        ErrorAppend("FileName [" & Fi_sl.Name & "]. Error In Function [InsertToDiscount]")
                        lb_Error = True
                        If File.Exists(ErrorFileDumpPath & "\" & Fi_sl.Name) Then File.Delete(ErrorFileDumpPath & "\" & Fi_sl.Name)
                        Fi_sl.MoveTo(ErrorFileDumpPath & "\" & Fi_sl.Name)
                        If File.Exists(ErrorFileDumpPath & "\" & Fi_dd.Name) Then File.Delete(ErrorFileDumpPath & "\" & Fi_dd.Name)
                        Fi_dd.MoveTo(ErrorFileDumpPath & "\" & Fi_dd.Name)
                        GoTo movenext
                    End If


                    SetLabelText("Loading [" & Fi_sl.Name & "]")
                    If Not InsertToSales(ls_sales_file, Fi_sl.Name) Then
                        ErrorAppend("FileName [" & Fi_sl.Name & "]. Error In Function [InsertToSales]")
                        lb_Error = True
                        If File.Exists(ErrorFileDumpPath & "\" & Fi_sl.Name) Then File.Delete(ErrorFileDumpPath & "\" & Fi_sl.Name)
                        Fi_sl.MoveTo(ErrorFileDumpPath & "\" & Fi_sl.Name)
                        If File.Exists(ErrorFileDumpPath & "\" & Fi_dd.Name) Then File.Delete(ErrorFileDumpPath & "\" & Fi_dd.Name)
                        Fi_dd.MoveTo(ErrorFileDumpPath & "\" & Fi_dd.Name)
                        GoTo movenext
                    End If


                    If Not invoiceprocess Then
                        SetLabelText("Creating AR Invoice Document " + Fi_sl.Name)
                        If Not Invoices(Fi_sl.Name) Then
                            'If SAPB1Company.InTransaction() Then
                            '    SAPB1Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            'End If


                            ProcessAppendNew("Processing " + Fi_sl.Name + " for invoice ends with error " + DateTime.Now.ToString("hh:mm:ss fff"))

                            ErrorAppend("FileName [" & Fi_sl.Name & "]. Error In Function [Invoice]")
                            lb_Error = True
                            If File.Exists(ErrorFileDumpPath & "\" & Fi_sl.Name) Then File.Delete(ErrorFileDumpPath & "\" & Fi_sl.Name)
                            Fi_sl.MoveTo(ErrorFileDumpPath & "\" & Fi_sl.Name)
                            If File.Exists(ErrorFileDumpPath & "\" & Fi_dd.Name) Then File.Delete(ErrorFileDumpPath & "\" & Fi_dd.Name)
                            Fi_dd.MoveTo(ErrorFileDumpPath & "\" & Fi_dd.Name)
                            GoTo movenext
                        Else

                            ProcessAppendNew("Processing " + Fi_sl.Name + " for invoice ends " + DateTime.Now.ToString("hh:mm:ss fff"))
                            invnow = True
                        End If

                    End If


                    If Not deliveryprocess Then
                        SetLabelText("Creating Delivery Document " + Fi_dd.Name)
                        If Not Delivery(Fi_dd.Name) Then
                            ProcessAppendNew("Processing " + Fi_dd.Name + " for delivery ends with error " + DateTime.Now.ToString("hh:mm:ss fff"))
                            ErrorAppend("FileName [" & Fi_dd.Name & "]. Error In Function [Delivery]")
                            lb_Error = True
                            If File.Exists(ErrorFileDumpPath & "\" & Fi_sl.Name) Then File.Delete(ErrorFileDumpPath & "\" & Fi_sl.Name)
                            Fi_sl.MoveTo(ErrorFileDumpPath & "\" & Fi_sl.Name)
                            If File.Exists(ErrorFileDumpPath & "\" & Fi_dd.Name) Then File.Delete(ErrorFileDumpPath & "\" & Fi_dd.Name)
                            Fi_dd.MoveTo(ErrorFileDumpPath & "\" & Fi_dd.Name)
                            GoTo movenext
                        Else

                            ProcessAppendNew("Processing " + Fi_dd.Name + " for delivery ends " + DateTime.Now.ToString("hh:mm:ss fff"))
                            delnow = True
                        End If
                    End If

                    If File.Exists(SuccessFileDumpPath & "\" & Fi_sl.Name) Then File.Delete(SuccessFileDumpPath & "\" & Fi_sl.Name)
                    Fi_sl.MoveTo(SuccessFileDumpPath & "\" & Fi_sl.Name)
                    If File.Exists(SuccessFileDumpPath & "\" & Fi_dd.Name) Then File.Delete(SuccessFileDumpPath & "\" & Fi_dd.Name)
                    Fi_dd.MoveTo(SuccessFileDumpPath & "\" & Fi_dd.Name)
                    If invoiceprocess = True And deliveryprocess = True Then
                        If delnow = False And invnow = False Then
                            SetLabelText("File [" & ls_sales_file & "] Already Processed")
                            SuccessAppend("File [" & ls_sales_file & "] Already Processed")
                        Else
                            SetLabelText("File [" & ls_sales_file & "] Processed Successfully")
                            SuccessAppend("File [" & ls_sales_file & "] Processed Successfully")
                        End If
                    Else
                        SetLabelText("File [" & ls_sales_file & "] Processed Successfully")
                        SuccessAppend("File [" & ls_sales_file & "] Processed Successfully")
                    End If

                    'Else
                    '    ErrorAppend("Cannot find Discount File[" & ls_discount_file & "].")
                    '    If File.Exists(SuccessFileDumpPath & "\" & Fi_sl.Name) Then File.Delete(SuccessFileDumpPath & "\" & Fi_sl.Name)
                    '    Fi_sl.MoveTo(SuccessFileDumpPath & "\" & Fi_sl.Name)
                    'End If
movenext:
                Next Fi_sl



                ProcessAppendNew("Processing files ends " + DateTime.Now.ToString("hh:mm:ss fff"))
            End If

        Catch ex As Exception
            ProcessAppendNew("Processing files ends with error " + DateTime.Now.ToString("hh:mm:ss fff"))

            ErrMsg = ex.Message
            Return False
        End Try

        Return True
    End Function

    Private Function InsertToSales(ByVal filename As String, ByVal fileN As String) As Boolean
        Dim sr As New StreamReader(filename)
        Dim ls_readline() As String
        Dim ls_TaxCode, ls_query, ls_itemcode, ls_POSCode, ls_freight, ls_itemgroup, ls_isbatch As String
        Dim oRS As SAPbobsCOM.Recordset
        Dim dr_freight As DataRow()
        'Dim ld_Qty, ls_document, dr_disc As Decimal


        ls_isbatch = "N"
        ls_POSCode = fileN.Substring(3, 2)
        ls_query = "select ""U_Freight"" from ""@FT_DISCTYPE"" WHERE ""U_DiscCode"" = 'SRVC'"
        oRS = Nothing
        oRS = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRS.DoQuery(ls_query)
        ls_freight = ""
        If oRS.RecordCount > 0 Then
            ls_freight = oRS.Fields.Item("U_Freight").Value
        Else
            ErrorAppend(fileN & ": No Service Charge Setup.")
            sr.Close()
            Return False
        End If

        oRS = Nothing
        GC.Collect()

        While Not sr.EndOfStream
            ls_readline = sr.ReadLine.Split(vbTab)

            ls_query = "SELECT ""Code"" FROM ""@POSEXITEM"" WHERE ""Code"" = '" & ls_readline(4) & "'"
            oRS = Nothing
            oRS = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery(ls_query)
            If oRS.RecordCount > 0 Then
                GoTo movenext
            End If
            oRS = Nothing
            GC.Collect()

            ls_query = "SELECT ""U_ItemCode"" FROM ""@POSITEM"" WHERE ""Code"" = '" & ls_readline(4) & "' AND U_POS = '" & ls_POSCode & "'"
            oRS = Nothing
            oRS = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ls_itemcode = ""
            oRS.DoQuery(ls_query)
            If oRS.RecordCount > 0 Then
                ls_itemcode = oRS.Fields.Item("U_ItemCode").Value
                If ls_itemcode = "201003" Then
                    ls_itemcode = "201003"
                End If
            Else
                ErrorAppend(fileN & ": Invalid ItemCode[" & ls_readline(4) & "]")
                sr.Close()
                Return False
            End If
            oRS = Nothing
            GC.Collect()

            'ls_query = "select ItemCode,ItmsGrpCod from OITM where U_" & ls_POSCode & "Item = '" & ls_readline(4) & "'"
            ls_query = "select ""ItemCode"",""ItmsGrpCod"", ""ManBtchNum"" from OITM where ""ItemCode"" = '" & ls_itemcode & "'"
            oRS = Nothing
            oRS = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery(ls_query)
            ls_itemcode = ""
            If oRS.RecordCount > 0 Then
                ls_itemcode = oRS.Fields.Item("ItemCode").Value
                ls_itemgroup = oRS.Fields.Item("ItmsGrpCod").Value
                ls_isbatch = oRS.Fields.Item("ManBtchNum").Value
            Else
                ErrorAppend(fileN & ": ItemCode[" & ls_readline(4) & "] Does not Exists on POS Item Mapping.")
                sr.Close()
                Return False
            End If
            oRS = Nothing
            GC.Collect()

            ls_query = "select ""U_VatCode"" from ""@FT_SALETYPE"" where ""Code"" = '" & ls_readline(3) & "'"
            oRS = Nothing
            oRS = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery(ls_query)

            ls_TaxCode = ""

            If oRS.RecordCount > 0 Then
                ls_TaxCode = oRS.Fields.Item("U_VatCode").Value
            Else
                ErrorAppend(fileN & ": Invalid Sales Type[" & ls_readline(3) & "]")
                sr.Close()
                Return False
            End If

            oRS = Nothing

            GC.Collect()

            dr_freight = oDataFreight.Select("Code='" & ls_freight & "'")
            If dr_freight.Length <= 0 Then

                Dim vatgroup As String = ""

                oRS = Nothing
                oRS = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRS.DoQuery("select ""VatGroupI"" from OEXD where ""ExpnsCode"" = '" & ls_freight & "'")

                If oRS.RecordCount > 0 Then
                    vatgroup = oRS.Fields.Item("VatGroupI").Value
                Else
                    ErrorAppend(fileN & ": No VatGroup Found for Frieght [" & ls_freight & "]")
                    sr.Close()
                    Return False
                End If

                oRS = Nothing
                GC.Collect()

                oDataFreight.Rows.Add(ls_freight, ls_readline(14), vatgroup)
                'Else
                'dr_freight(0)("Amount") = dr_freight(0)("Amount") + ls_readline(14)
            End If


            'dr_disc = odt_discount.Select("Base_Line_ID='" & ls_readline(3) & "'")
            'If dr_disc.Length > 0 Then
            '    For li_disc_row As Integer = 0 To dr_disc.Length - 1
            '        ls_document = dr_disc(li_disc_row)("Document").ToString().Trim()
            '        If ls_document = "DLRY" Then
            '            ld_Qty = ld_Qty + dr_disc(li_disc_row)("Discount_Quantity")
            '        End If
            '    Next li_disc_row
            'End If
            oDataSales.Rows.Add(ls_readline(0), ls_readline(1), ls_readline(2), ls_readline(3), ls_itemcode _
                               , ls_readline(5), ls_readline(6), ls_readline(7), ls_readline(8), ls_TaxCode _
                               , ls_readline(10), ls_readline(11), ls_readline(12), ls_readline(13), ls_readline(14), ls_itemgroup, ls_isbatch, ls_readline(15))
movenext:
        End While
        sr.Close()
        Return True
    End Function

    Private Function InsertToDiscount(ByVal filename As String, ByVal FileN As String) As Boolean
        Dim sr As New StreamReader(filename)
        Dim ls_readline() As String
        Dim dt_row, dr_freight As DataRow()
        Dim ls_query, ls_itemcode, ls_POSCode, ls_Type, ls_costcode, ls_document, ls_freight, ls_empCode, ls_itemgroup, ls_isbatch As String

        Dim dblPrice As Double

        Dim oRS As SAPbobsCOM.Recordset
        ls_POSCode = FileN.Substring(3, 2)
        While Not sr.EndOfStream
            ls_readline = sr.ReadLine.Split(vbTab)
            ls_query = "SELECT ""Code"" FROM ""@POSEXITEM"" WHERE ""Code"" = '" & ls_readline(4) & "'"
            oRS = Nothing
            oRS = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery(ls_query)
            If oRS.RecordCount > 0 Then
                GoTo movenext
            End If
            oRS = Nothing
            GC.Collect()

            ls_isbatch = "N"

            ls_query = "SELECT ""U_ItemCode"" FROM ""@POSITEM"" WHERE ""Code"" = '" & ls_readline(4) & "' AND ""U_POS"" = '" & ls_POSCode & "'"
            oRS = Nothing
            oRS = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ls_itemcode = ""
            oRS.DoQuery(ls_query)
            If oRS.RecordCount > 0 Then
                ls_itemcode = oRS.Fields.Item("U_ItemCode").Value
            Else
                ErrorAppend(FileN & ": ItemCode[" & ls_readline(4) & "] Does not Exists on POS Item Mapping.")
                sr.Close()
                Return False
            End If
            oRS = Nothing
            GC.Collect()

            'ls_query = "select OITM.ItemCode,OITB.U_Type,OITM.ItmsGrpCod from OITM LEFT JOIN OITB ON OITM.ItmsGrpCod = OITB.ItmsGrpCod WHERE OITM.U_" & ls_POSCode & "Item = '" & ls_readline(4) & "'"
            ls_query = "select OITM.""ItemCode"",OITB.""U_Type"",OITM.""ItmsGrpCod"", OITM.""ManBtchNum"" from OITM LEFT JOIN OITB ON OITM.""ItmsGrpCod"" = OITB.""ItmsGrpCod"" WHERE OITM.""ItemCode"" = '" & ls_itemcode & "'"
            oRS = Nothing
            oRS = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery(ls_query)
            ls_itemcode = ""
            ls_itemgroup = ""
            If oRS.RecordCount > 0 Then
                ls_itemcode = oRS.Fields.Item("ItemCode").Value
                ls_itemgroup = oRS.Fields.Item("ItmsGrpCod").Value
                ls_Type = oRS.Fields.Item("U_Type").Value
                ls_isbatch = oRS.Fields.Item("ManBtchNum").Value
            Else
                ErrorAppend(FileN & ": Invalid ItemCode[" & ls_readline(4) & " - " & ls_itemcode & "]")
                sr.Close()
                Return False
            End If
            oRS = Nothing
            GC.Collect()

            'ls_query = "select ""CardCode"" from OCRD where ""U_" & ls_POSCode & "Emp"" = '" & ls_readline(9) & "'"
            'oRS = Nothing
            'oRS = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRS.DoQuery(ls_query)
            'ls_empCode = ""
            'If oRS.RecordCount > 0 Then
            '    ls_empCode = oRS.Fields.Item("CardCode").Value
            'Else
            '    ls_empCode = ""
            'End If
            'oRS = Nothing

            GC.Collect()

            ls_query = "select ""U_CostCode"",""U_Document"",""U_Freight"" from ""@FT_DISCTYPE"" " &
                        "where ""U_DiscCode"" = '" & ls_readline(6).Trim() & "' " &
                        "and case ""U_Document"" when 'ARIN' then ""U_Type"" else '" & ls_Type & "' end = '" & ls_Type & "'"
            oRS = Nothing
            oRS = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery(ls_query)
            ls_costcode = ""
            ls_document = ""
            ls_freight = ""
            If oRS.RecordCount > 0 Then
                ls_costcode = oRS.Fields.Item("U_CostCode").Value
                ls_document = oRS.Fields.Item("U_Document").Value
                ls_freight = oRS.Fields.Item("U_Freight").Value
            Else
                ErrorAppend(FileN & ": Invalid Discount Code[" & ls_readline(6) & "] Type [" & ls_Type & "]")
                sr.Close()
                Return False
            End If
            oRS = Nothing
            GC.Collect()

            dr_freight = oDataFreight.Select("Code='" & ls_freight & "'")
            If dr_freight.Length <= 0 Then
                oDataFreight.Rows.Add(ls_freight, 0 - ls_readline(8))
            Else
                dr_freight(0)("Amount") = dr_freight(0)("Amount") - ls_readline(8)
            End If

            If ls_document = "DLRY" Then
                dt_row = oDataDelivery.Select("Discount_Type='" & ls_readline(6) & "'")

                If dt_row.Length <= 0 Then
                    oDataDelivery.Rows.Add(ls_readline(0), ls_readline(1), ls_readline(6), ls_document)
                End If
            End If

            dblPrice = Math.Round(Convert.ToDouble(ls_readline(8)) / Convert.ToDouble(ls_readline(7)), 2)



            oDataDiscounts.Rows.Add(ls_readline(0), ls_readline(1), ls_readline(2), ls_readline(3), ls_itemcode _
                               , ls_readline(5), ls_readline(6), ls_readline(7), dblPrice.ToString(), ls_readline(9) _
                               , ls_costcode, ls_document, ls_freight, ls_itemgroup, ls_isbatch, ls_readline(10))
movenext:
        End While
        sr.Close()
        Return True
    End Function


    Public Function Invoices(ByVal Filename As String) As Boolean
        Dim oDraft As SAPbobsCOM.Documents
        Dim oDraft_Lines As SAPbobsCOM.Document_Lines
        Dim dr_sales, dr_disc As DataRow()
        Dim ls_transdate, ls_storecode, ls_lineid, ls_salestype, ls_itemcode, ls_itemdesc, ls_tax_code, ls_cardcode, ls_query, ls_itemgroup, ls_uomcode As String
        Dim ls_costcode, ls_document, ls_freight, ls_ErrMsg, ls_whscode As String
        Dim ld_quantity, ld_taxamount, ld_price, ld_salesamount, ld_discquantity, ld_discamount, ld_totalguest, ld_servcharge As Decimal
        Dim ld_dlryqty, ld_arinqty, ld_freightamt As Decimal
        Dim lb_add, lb_freight As Boolean
        Dim oRS As SAPbobsCOM.Recordset
        Dim li_result, li_ErrCode As Integer
        Dim DocEntry As Long
        ls_storecode = ""
        ls_freight = ""
        li_ErrCode = 0
        ls_ErrMsg = ""
        ls_itemcode = ""
        ls_tax_code = ""
        ls_itemgroup = ""
        ls_whscode = ""

        Dim treetype As String = "N"

        Dim ls_comp_itemcode, ls_comp_whscode, ls_comp_price, ls_comp_itmgroup, ls_comp_quantity, ls_comp_isbatch, ls_comp_uomcode As String

        Dim jsonstring As String = ""
        Dim jsonstringlines As String = ""
        Dim jsonexpenses As String = ""
        Dim jsonbatches As String = ""

        Dim isbatch As String = ""

        Dim ls_ocrcode3 = ""

        Try
            dr_sales = oDataSales.Select()
            If dr_sales.Length > 0 Then
                lb_add = False

                B1Connect(SAPB1CompanyName)
                ProcessAppendNew("Processing " + Filename + " for invoice starts " + DateTime.Now.ToString("hh:mm:ss fff"))

                For li_sales_row As Integer = 0 To dr_sales.Length - 1

                    ls_transdate = dr_sales(li_sales_row)("Transation_date").ToString().Trim()
                    ls_storecode = dr_sales(li_sales_row)("Store_Code").ToString().Trim()
                    ls_lineid = dr_sales(li_sales_row)("Line_ID").ToString().Trim()
                    ls_salestype = dr_sales(li_sales_row)("Sales_Type").ToString().Trim()
                    ls_itemcode = dr_sales(li_sales_row)("Item_Code").ToString().Trim()
                    ls_itemdesc = dr_sales(li_sales_row)("Item_Description").ToString().Trim()
                    ld_quantity = dr_sales(li_sales_row)("Sales_Quantity").ToString().Trim()
                    ld_price = dr_sales(li_sales_row)("Price").ToString().Trim()
                    ld_salesamount = dr_sales(li_sales_row)("Sales_Amount").ToString().Trim()
                    ls_tax_code = dr_sales(li_sales_row)("Tax_Code").ToString().Trim()
                    ld_taxamount = dr_sales(li_sales_row)("Tax_Amount").ToString().Trim()
                    ld_discquantity = dr_sales(li_sales_row)("Discount_Quantity").ToString().Trim()
                    ld_discamount = dr_sales(li_sales_row)("Discount_Amount").ToString().Trim()
                    ld_totalguest = dr_sales(li_sales_row)("Total_Guest").ToString().Trim()
                    ls_itemgroup = dr_sales(li_sales_row)("Item_Group").ToString().Trim()
                    isbatch = dr_sales(li_sales_row)("IsBatch").ToString().Trim()
                    ld_servcharge = ld_servcharge + dr_sales(li_sales_row)("Service_charge").ToString().Trim()
                    ld_freightamt = 0

                    ls_ocrcode3 = dr_sales(li_sales_row)("CostingCode3").ToString().Trim()

                    treetype = "N"

                    dr_disc = oDataDiscounts.Select("Base_Line_ID='" & ls_lineid & "'")
                    If dr_disc.Length > 0 Then
                        ld_arinqty = 0
                        ld_dlryqty = 0
                        For li_disc_row As Integer = 0 To dr_disc.Length - 1
                            ls_costcode = dr_disc(li_disc_row)("Cost_Code").ToString().Trim()
                            ls_document = dr_disc(li_disc_row)("Document").ToString().Trim()
                            ls_freight = dr_disc(li_disc_row)("Freight").ToString().Trim()
                            If ls_document = "DLRY" Then
                                ld_quantity = ld_quantity - dr_disc(li_disc_row)("Discount_Quantity")
                            Else
                                ld_freightamt = ld_freightamt + dr_disc(li_disc_row)("Discount_Amount")
                            End If
                        Next li_disc_row
                    End If

                    If ld_quantity = 0 Then
                        GoTo movenext
                    End If

                    If lb_add Then
                        'oDraft_Lines.Add()
                        jsonstringlines += ","
                    Else
                        ls_query = "select ""CardCode"",""U_WhsCode"" from OCRD where ""U_Store""='" & ls_storecode & "'"
                        oRS = Nothing
                        oRS = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRS.DoQuery(ls_query)
                        ls_cardcode = ""

                        If oRS.RecordCount > 0 Then
                            ls_cardcode = oRS.Fields.Item("CardCode").Value
                            ls_whscode = oRS.Fields.Item("U_WhsCode").Value
                        Else
                            ErrorAppend(Filename & ": Invalid Store Code [" & ls_storecode & "]")
                            Return False
                        End If

                        oRS = Nothing
                        Dim year As String = (2000 + Convert.ToInt32(ls_transdate.Substring(4, 2))).ToString()
                        jsonstring = "{" +
                                           """CardCode"":""" + ls_storecode + """," +
                                           """DocDate"":""" + year + "-" + ls_transdate.Substring(0, 2) & "-" & ls_transdate.Substring(2, 2) + """," +
                                           """U_TotalGuest"":""" + ld_totalguest.ToString() + """," +
                                           """U_FileName"":""" + Filename + """," +
                                           """DocumentLines"":[@DOCUMENTLINE]," +
                                           """DocumentAdditionalExpenses"":[@EXPENSES]" +
                                           "}"


                    End If

                    Dim isbom As Boolean = False

                    ls_query = "SELECT ""U_Whse"" FROM ""@WHSEMAPPING"" WHERE ""U_Store"" = '" & ls_storecode & "' and ""U_ItemGroup"" ='" & ls_itemgroup & "'"
                    FMSValue = FMS(ls_query)
                    If FMSValue = "-1" Then
                        ErrorAppend(ErrMsg)
                        Return False
                    Else
                        If FMSValue <> "" Then ls_whscode = FMSValue
                    End If

                    jsonbatches = ""

                    jsonstringlines += "{" +
                                           """ItemCode"":""" + ls_itemcode + """," +
                                           """Quantity"":""" + ld_quantity.ToString() + """," +
                                           """UnitPrice"":""" + ld_price.ToString() + """," +
                                           """CostingCode"":""" + ls_storecode + """," +
                                           """CostingCode3"":""" + ls_ocrcode3 + """," +
                                           """WarehouseCode"":""" + ls_whscode + """," +
                                           """COGSCostingCode"":""" + ls_storecode + """," +
                                           """VatGroup"":""" + ls_tax_code + """," +
                                           """U_ItmGrp"":""" + ls_itemgroup + """," +
                                           """U_SalesTyp"":""" + ls_salestype + """," +
                                           """BatchNumbers"":[@BATCHES]" +
                                           "}"


                    If isbatch = "Y" Then
                        jsonbatches = GetBatches(ls_itemcode, ld_quantity, ls_whscode)
                        If jsonbatches = "-1" Then
                            ErrorAppend(ErrMsg + " Getting Batch For Father ItemCode: " + ls_itemcode)
                            Return False
                        End If
                    End If

                    jsonstringlines = jsonstringlines.Replace("@BATCHES", jsonbatches)

                    SQLQuery = "select itt1.""Code"", itt1.""Warehouse"", itt1.""Price"", ""oitm2"".""ItmsGrpCod"", itt1.""Quantity"", ""oitm2"".""ManBtchNum"", IFNULL(""oitm2"".""IUoMEntry"", ""oitm2"".""UgpEntry"") as ""UomCode"" from oitm left join oitt on oitt.""Code"" = oitm.""ItemCode"" left join itt1 on itt1.""Father"" = oitt.""Code"" left join oitm as ""oitm2"" on ""oitm2"".""ItemCode"" = itt1.""Code"" where oitm.""ItemCode"" = '" + ls_itemcode + "' and oitm.""PrchseItem"" = 'N' and oitm.""SellItem"" = 'Y' and oitm.""InvntItem"" = 'N' and oitt.""TreeType"" = 'S'"


                    'Child Items Part
                    oRS = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRS.DoQuery(SQLQuery)
                    If oRS.RecordCount > 0 Then
                        oRS.MoveFirst()
                        For i As Integer = 0 To oRS.RecordCount - 1

                            jsonbatches = ""
                            ls_comp_itemcode = oRS.Fields.Item("Code").Value.ToString().Trim()
                            ls_comp_whscode = oRS.Fields.Item("Warehouse").Value.ToString().Trim()
                            ls_comp_price = oRS.Fields.Item("Price").Value.ToString().Trim()
                            ls_comp_itmgroup = oRS.Fields.Item("ItmsGrpCod").Value.ToString().Trim()
                            ls_comp_quantity = oRS.Fields.Item("Quantity").Value.ToString().Trim()

                            ls_comp_isbatch = oRS.Fields.Item("ManBtchNum").Value.ToString().Trim()
                            ls_comp_uomcode = oRS.Fields.Item("UomCode").Value.ToString().Trim()

                            ls_comp_quantity = (ld_quantity * Convert.ToDecimal(ls_comp_quantity)).ToString()

                            ls_query = "SELECT ""U_Whse"" FROM ""@WHSEMAPPING"" WHERE ""U_Store"" = '" & ls_storecode & "' and ""U_ItemGroup"" ='" & ls_comp_itmgroup & "'"
                            FMSValue = FMS(ls_query)
                            If FMSValue = "-1" Then
                                ErrorAppend(ErrMsg)
                                Return False
                            Else
                                If FMSValue <> "" Then ls_comp_whscode = FMSValue
                            End If
                            treetype = "I"

                            jsonstringlines += ",{" +
                                       """ItemCode"":""" + ls_comp_itemcode + """," +
                                       """Quantity"":""" + ls_comp_quantity + """," +
                                       """CostingCode"":""" + ls_storecode + """," +
                                       """CostingCode3"":""" + ls_ocrcode3 + """," +
                                       """WarehouseCode"":""" + ls_comp_whscode + """," +
                                       """COGSCostingCode"":""" + ls_storecode + """," +
                                       """VatGroup"":""" + ls_tax_code + """," +
                                       """U_ItmGrp"":""" + ls_comp_itmgroup + """," +
                                       """U_SalesTyp"":""" + ls_salestype + """," +
                                       """TreeType"":""" + treetype + """," +
                                       """UoMEntry"":""" + ls_comp_uomcode + """," +
                                       """BatchNumbers"":[@BATCHES]" +
                                       "}"

                            If ls_comp_isbatch = "Y" Then
                                jsonbatches = GetBatches(ls_comp_itemcode, Convert.ToDecimal(ls_comp_quantity), ls_whscode)
                                If jsonbatches = "-1" Then
                                    ErrorAppend(ErrMsg + " Getting Batch For Component ItemCode: " + ls_itemcode)
                                    Return False
                                End If
                            End If
                            jsonstringlines = jsonstringlines.Replace("@BATCHES", jsonbatches)
                            oRS.MoveNext()
                        Next i
                    Else
                        ls_comp_itemcode = ls_itemcode
                    End If
                    oRS = Nothing

                    lb_add = True
movenext:
                Next li_sales_row

                'Add Freight
                Dim dr_freight As DataRow()
                dr_freight = oDataFreight.Select()
                If dr_freight.Length > 0 Then
                    For li_freight_row As Integer = 0 To dr_freight.Length - 1
                        If dr_freight(li_freight_row)("Code") <> "" Then
                            If lb_freight Then jsonexpenses += ","
                            'oDraft.Expenses.ExpenseCode = dr_freight(li_freight_row)("Code")
                            'oDraft.Expenses.LineTotal = dr_freight(li_freight_row)("Amount")
                            'oDraft.Expenses.DistributionRule = ls_storecode

                            jsonexpenses += "{" +
                                """ExpenseCode"":""" + dr_freight(li_freight_row)("Code") + """," +
                                """LineTotal"":""" + dr_freight(li_freight_row)("Amount") + """," +
                                """DistributionRule"":""" + ls_storecode + """" +
                                "}"


                            lb_freight = True
                        End If
                    Next li_freight_row
                End If

                jsonstring = jsonstring.Replace("@DOCUMENTLINE", jsonstringlines)
                jsonstring = jsonstring.Replace("@EXPENSES", jsonexpenses)

                If relogServiceLayer Then

                    If Not SLayer.OperationSuccess(SLayer.Login(relogServiceLayer)) Then
                        Throw New Exception(SLayer.GetErrorCode() + " " + SLayer.GetErrorMessage())
                    Else
                        relogServiceLayer = False
                    End If
                Else
                    'FORCE RELOG BEFORE POSTING FOR YABU
                    If Not SLayer.OperationSuccess(SLayer.Login(True)) Then
                        Throw New Exception(SLayer.GetErrorCode() + " " + SLayer.GetErrorMessage())
                    Else
                        relogServiceLayer = False
                    End If

                End If

                SLayer.JSONBody = jsonstring
                SLayer.SLMethod = SLayer.PostMethod
                If Not SLayer.OperationSuccess(SLayer.SendInvoice()) Then
                    ErrorAppend(Filename + ": Error Code: " + SLayer.GetErrorCode + " Error Message: " + SLayer.GetErrorMessage())
                    Return False
                Else
                    SuccessAppend("Invoice [" + SLayer.GetDocKey() + "] has been created for file [" + Filename + "]")
                End If


                jsonstring = ""
                jsonstringlines = ""
                jsonexpenses = ""
                jsonbatches = ""


            End If

        Catch ex As Exception

            ErrorAppend(Filename + " " + ex.Message)
            Return False
        End Try
        Return True
    End Function

    Function FMS(ByVal as_query As String) As String
        Try

            Dim oRS As SAPbobsCOM.Recordset

            B1Connect(SAPB1CompanyName)

            oRS = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery(as_query)
            If oRS.RecordCount > 0 Then
                Return oRS.Fields.Item(0).Value.ToString().Trim()
            End If
            oRS = Nothing
            GC.Collect()
        Catch ex As Exception
            ErrMsg = ex.Message
            Return "-1"
        End Try
        Return ""
    End Function

    Public Function GetBatches(ByVal as_itemcode As String, ByVal ad_qty As Decimal, ByVal as_whscode As String) As String
        Try
            B1Connect(SAPB1CompanyName)
            Dim oRs, oRs1 As SAPbobsCOM.Recordset
            Dim oDraft As SAPbobsCOM.Documents
            Dim oDraft_Lines As SAPbobsCOM.Document_Lines = Nothing
            Dim oBatches As SAPbobsCOM.BatchNumbers = Nothing
            Dim ls_ItemCode, ls_WhsCode, ls_BatchNum, ls_ErrMsg As String
            Dim ld_Quantity, ld_Remaining_Quantity, ld_BatchQuantity As Decimal
            Dim li_errCode, li_LineId As Integer
            Dim bl_Batch As Boolean = False
            Dim odt_batch As New DataTable("Batch")
            Dim odatarows As DataRow() = Nothing

            Dim jsonstring As String = ""

            odt_batch.Columns.Add("Row", GetType(System.Int32))
            odt_batch.Columns.Add("ItemCode", GetType(System.String))
            odt_batch.Columns.Add("WhsCode", GetType(System.String))
            odt_batch.Columns.Add("BatchNum", GetType(System.String))
            odt_batch.Columns.Add("Quantity", GetType(System.Decimal))

            SQLQuery = ""

            ld_Quantity = ad_qty

            oRs1 = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs1.DoQuery("SELECT OIBT.""BatchNum"",OIBT.""Quantity"" FROM OIBT WHERE OIBT.""ItemCode"" = '" & as_itemcode & "' AND OIBT.""WhsCode"" = '" & as_whscode & "' AND OIBT.""Quantity"" > 0 ORDER BY OIBT.""ExpDate"" Desc")

            If oRs1.RecordCount > 0 Then
                oRs1.MoveFirst()
                While Not oRs1.EoF
                    ls_BatchNum = oRs1.Fields.Item("BatchNum").Value
                    ld_BatchQuantity = Convert.ToDecimal(oRs1.Fields.Item("Quantity").Value)
                    ld_Remaining_Quantity = ld_Quantity - ld_BatchQuantity
                    If ld_Remaining_Quantity <= 0 Then
                        odt_batch.Rows.Add(odt_batch.Rows.Count, as_itemcode, as_whscode, ls_BatchNum, ld_Quantity)
                        ld_Quantity = ld_Remaining_Quantity
                        Exit While
                    Else
                        ld_Quantity = ld_Remaining_Quantity
                        odt_batch.Rows.Add(odt_batch.Rows.Count, as_itemcode, as_whscode, ls_BatchNum, ld_BatchQuantity)
                    End If
                    oRs1.MoveNext()
                End While
                oRs1 = Nothing
                GC.Collect()

                If ld_Quantity > 0 Then
                    Throw New Exception("Item Code [" & as_itemcode & "] Warehouse Code [" & as_whscode & "]" & " Falls Into Negative Inventory")
                End If

                Dim addbatch As Boolean = False
                For Each row As DataRow In odt_batch.Rows
                    If addbatch Then jsonstring += ","

                    ls_BatchNum = row("BatchNum").ToString()
                    ld_BatchQuantity = Convert.ToDecimal(row("Quantity").ToString())

                    jsonstring += "{" +
                        """BatchNumber"":""" + ls_BatchNum + """," +
                        """Quantity"":""" + ld_BatchQuantity.ToString() + """" +
                        "}"

                    addbatch = True
                Next

                Return jsonstring

            Else
                Throw New Exception("Item Code [" & as_itemcode & "] Warehouse Code [" & as_whscode & "]" & " Falls Into Negative Inventory - No available batch found")
            End If

        Catch ex As Exception
            ErrMsg = ex.Message
            Return "-1"
        End Try
        Return ""
    End Function

    Public Function Delivery(ByVal Filename As String) As Boolean
        Dim ors As SAPbobsCOM.Recordset
        Dim oDraft As SAPbobsCOM.Documents
        Dim oDraft_Lines As SAPbobsCOM.Document_Lines
        Dim dr_delivery, dr_disc As DataRow()
        Dim ld_Quantity, ld_price As Decimal
        Dim ls_ItemCode, ls_itemgroup, ls_ErrMsg As String
        Dim ls_disctype, ls_storecode, ls_date, ls_document, ls_empcode, ls_costcode, ls_whscode, ls_query, ls_compsales As String
        Dim li_result, li_errCode As Integer
        Dim DocEntry As Long
        Dim lb_Add As Boolean

        Dim ls_ocrcode3 = ""

        Dim treetype As String = "N"

        Dim ls_comp_itemcode, ls_comp_whscode, ls_ocrcode2, ls_comp_price, ls_comp_itmgroup, ls_comp_quantity, ls_comp_isbatch, ls_comp_uomcode As String

        Dim jsonstring As String = ""
        Dim jsonstringlines As String = ""
        Dim jsonexpenses As String = ""
        Dim jsonbatches As String = ""

        Dim isbatch As String = ""

        ls_ErrMsg = ""
        ls_ItemCode = ""
        ls_itemgroup = ""
        ls_query = ""

        dr_delivery = oDataDelivery.Select()
        Try
            If dr_delivery.Length > 0 Then

                B1Connect(SAPB1CompanyName)

                ProcessAppendNew("Processing " + Filename + " for delivery starts " + DateTime.Now.ToString("hh:mm:ss fff"))

                For li_del_row As Integer = 0 To dr_delivery.Length - 1


                    ls_date = dr_delivery(li_del_row)("Transation_Date").ToString().Trim()
                    ls_storecode = dr_delivery(li_del_row)("Store_Code").ToString().Trim()
                    ls_disctype = dr_delivery(li_del_row)("Discount_Type").ToString().Trim()
                    ls_document = dr_delivery(li_del_row)("Document").ToString().Trim()
                    ls_whscode = ""
                    dr_disc = oDataDiscounts.Select("Discount_Type='" & ls_disctype & "' and Document='" & ls_document & "'")
                    If dr_disc.Length > 0 Then

                        ls_query = "SELECT ""U_CompSales"" FROM ""@FT_DISCTYPE"" WHERE ""U_DiscCode""='" & ls_disctype & "'"
                        ors = Nothing
                        ors = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ors.DoQuery(ls_query)
                        ls_compsales = ""
                        If ors.RecordCount > 0 Then
                            ls_compsales = ors.Fields.Item("U_CompSales").Value
                        Else
                            ls_compsales = ""
                        End If
                        ors = Nothing
                        GC.Collect()
                        Dim year As String = (2000 + Convert.ToInt32(ls_date.Substring(4, 2))).ToString()
                        jsonstring = "{" +
                                           """CardCode"":""" + ls_storecode + """," +
                                           """DocDate"":""" + year + "-" + ls_date.Substring(0, 2) & "-" & ls_date.Substring(2, 2) + """," +
                                           """U_CompSales"":""" + ls_compsales.ToString() + """," +
                                           """U_FileName"":""" + Filename + """," +
                                           """DocumentLines"":[@DOCUMENTLINE]" +
                                           "}"


                        lb_Add = False

                        ls_query = "select ""CardCode"",""U_WhsCode"" from OCRD where ""U_Store""='" & ls_storecode & "'"
                        ors = Nothing
                        ors = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ors.DoQuery(ls_query)
                        If ors.RecordCount > 0 Then
                            ls_whscode = ors.Fields.Item("U_WhsCode").Value
                        Else
                            ErrorAppend(Filename & ": Invalid Store Code [" & ls_storecode & "]")
                            Return False
                        End If
                        ors = Nothing
                        GC.Collect()

                        For li_disc_row As Integer = 0 To dr_disc.Length - 1

                            ld_Quantity = dr_disc(li_disc_row)("Discount_Quantity")
                            ld_price = dr_disc(li_disc_row)("Discount_Amount")
                            ls_ItemCode = dr_disc(li_disc_row)("Item_Code")
                            ls_empcode = dr_disc(li_disc_row)("Employee_Code")
                            ls_costcode = dr_disc(li_disc_row)("Cost_Code")
                            ls_itemgroup = dr_disc(li_disc_row)("Item_Group")
                            isbatch = dr_disc(li_disc_row)("IsBatch")

                            ls_ocrcode3 = dr_disc(li_disc_row)("CostingCode3")

                            If lb_Add Then jsonstringlines += ","



                            ors = Nothing
                            ls_query = "SELECT ""U_Whse"" FROM ""@WHSEMAPPING"" WHERE ""U_Store"" = '" & ls_storecode & "' and ""U_ItemGroup"" ='" & ls_itemgroup & "'"
                            FMSValue = FMS(ls_query)
                            If FMSValue = "-1" Then
                                ErrorAppend(ErrMsg)
                                Return False
                            Else
                                If FMSValue <> "" Then ls_whscode = FMSValue
                            End If

                            jsonbatches = ""

                            jsonstringlines += "{" +
                                           """ItemCode"":""" + ls_ItemCode + """," +
                                           """Quantity"":""" + ld_Quantity.ToString() + """," +
                                           """UnitPrice"":""" + ld_price.ToString() + """," +
                                           """CostingCode"":""" + ls_storecode + """," +
                                           """CostingCode3"":""" + ls_ocrcode3 + """," +
                                           """WarehouseCode"":""" + ls_whscode + """," +
                                           """COGSCostingCode2"":""" + ls_costcode + """," +
                                           """VatGroup"":""" + "X0" + """," +
                                           """U_ItmGrp"":""" + ls_itemgroup + """," +
                                           """U_EmpName"":""" + ls_empcode + """," +
                                           """BatchNumbers"":[@BATCHES]" +
                                           "}"


                            If isbatch = "Y" Then
                                jsonbatches = GetBatches(ls_ItemCode, ld_Quantity, ls_whscode)
                                If jsonbatches = "-1" Then
                                    ErrorAppend(ErrMsg + " Getting Batch For Father ItemCode: " + ls_ItemCode)
                                    Return False
                                End If
                            End If

                            jsonstringlines = jsonstringlines.Replace("@BATCHES", jsonbatches)

                            'SQLQuery = "select itt1.""Code"", itt1.""Warehouse"", itt1.""Price"", ""oitm2"".""ItmsGrpCod"", itt1.""Quantity"", ""oitm2"".""ManBtchNum"" from oitm left join oitt on oitt.""Code"" = oitm.""ItemCode"" left join itt1 on itt1.""Father"" = oitt.""Code"" left join oitm as ""oitm2"" on ""oitm2"".""ItemCode"" = itt1.""Code"" where oitm.""ItemCode"" = '" + ls_ItemCode + "' and oitm.""PrchseItem"" = 'N' and oitm.""SellItem"" = 'Y' and oitm.""InvntItem"" = 'N' and oitt.""TreeType"" = 'S'"

                            SQLQuery = "select itt1.""Code"", itt1.""Warehouse"", itt1.""Price"", ""oitm2"".""ItmsGrpCod"", itt1.""Quantity"", ""oitm2"".""ManBtchNum"", IFNULL(""oitm2"".""IUoMEntry"", ""oitm2"".""UgpEntry"") as ""UomCode"" from oitm left join oitt on oitt.""Code"" = oitm.""ItemCode"" left join itt1 on itt1.""Father"" = oitt.""Code"" left join oitm as ""oitm2"" on ""oitm2"".""ItemCode"" = itt1.""Code"" where oitm.""ItemCode"" = '" + ls_ItemCode + "' and oitm.""PrchseItem"" = 'N' and oitm.""SellItem"" = 'Y' and oitm.""InvntItem"" = 'N' and oitt.""TreeType"" = 'S'"

                            ors = SAPB1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            ors.DoQuery(SQLQuery)

                            If ors.RecordCount > 0 Then
                                ors.MoveFirst()
                                For i As Integer = 0 To ors.RecordCount - 1

                                    jsonbatches = ""
                                    ls_comp_itemcode = ors.Fields.Item("Code").Value.ToString().Trim()
                                    ls_comp_whscode = ors.Fields.Item("Warehouse").Value.ToString().Trim()
                                    ls_comp_price = ors.Fields.Item("Price").Value.ToString().Trim()
                                    ls_comp_itmgroup = ors.Fields.Item("ItmsGrpCod").Value.ToString().Trim()
                                    ls_comp_quantity = ors.Fields.Item("Quantity").Value.ToString().Trim()
                                    ls_comp_uomcode = ors.Fields.Item("UomCode").Value.ToString().Trim()

                                    ls_comp_isbatch = ors.Fields.Item("ManBtchNum").Value.ToString().Trim()

                                    ls_comp_quantity = (ld_Quantity * Convert.ToDecimal(ls_comp_quantity)).ToString()

                                    ls_query = "SELECT ""U_Whse"" FROM ""@WHSEMAPPING"" WHERE ""U_Store"" = '" & ls_storecode & "' and ""U_ItemGroup"" ='" & ls_comp_itmgroup & "'"
                                    FMSValue = FMS(ls_query)
                                    If FMSValue = "-1" Then
                                        ErrorAppend(ErrMsg)
                                        Return False
                                    Else
                                        If FMSValue <> "" Then ls_comp_whscode = FMSValue
                                    End If
                                    treetype = "I"

                                    jsonstringlines += ",{" +
                                           """ItemCode"":""" + ls_comp_itemcode + """," +
                                           """Quantity"":""" + ls_comp_quantity.ToString() + """," +
                                           """CostingCode"":""" + ls_storecode + """," +
                                           """CostingCode3"":""" + ls_ocrcode3 + """," +
                                           """WarehouseCode"":""" + ls_comp_whscode + """," +
                                           """COGSCostingCode2"":""" + ls_costcode + """," +
                                           """VatGroup"":""" + "X0" + """," +
                                           """U_ItmGrp"":""" + ls_itemgroup + """," +
                                           """U_EmpName"":""" + ls_empcode + """," +
                                           """TreeType"":""" + treetype + """," +
                                           """UoMEntry"":""" + ls_comp_uomcode + """," +
                                           """BatchNumbers"":[@BATCHES]" +
                                           "}"

                                    If ls_comp_isbatch = "Y" Then
                                        jsonbatches = GetBatches(ls_comp_itemcode, Convert.ToDecimal(ls_comp_quantity), ls_whscode)
                                        If jsonbatches = "-1" Then
                                            ErrorAppend(ErrMsg + " Getting Batch For Component ItemCode: " + ls_ItemCode)
                                            Return False
                                        End If
                                    End If
                                    jsonstringlines = jsonstringlines.Replace("@BATCHES", jsonbatches)
                                    ors.MoveNext()
                                Next i
                            End If
                            ors = Nothing

                            lb_Add = True
                        Next li_disc_row


                        jsonstring = jsonstring.Replace("@DOCUMENTLINE", jsonstringlines)

                        If relogServiceLayer Then
                            If Not SLayer.OperationSuccess(SLayer.Login(relogServiceLayer)) Then
                                Throw New Exception(SLayer.GetErrorCode() + " " + SLayer.GetErrorMessage())
                            Else
                                relogServiceLayer = False
                            End If
                        Else
                            'FORCE RELOG BEFORE POSTING FOR YABU
                            If Not SLayer.OperationSuccess(SLayer.Login(True)) Then
                                Throw New Exception(SLayer.GetErrorCode() + " " + SLayer.GetErrorMessage())
                            Else
                                relogServiceLayer = False
                            End If
                        End If


                        SLayer.JSONBody = jsonstring
                        SLayer.SLMethod = SLayer.PostMethod
                        If Not SLayer.OperationSuccess(SLayer.SendDelivery()) Then
                            ErrorAppend(Filename + ": Error Code: " + SLayer.GetErrorCode + " Error Message: " + SLayer.GetErrorMessage())
                            Return False
                        Else
                            SuccessAppend("Delivery [" + SLayer.GetDocKey() + "] has been created for file [" + Filename + "]")
                        End If
                        jsonstring = ""
                        jsonstringlines = ""
                        jsonexpenses = ""
                        jsonbatches = ""
                        GC.Collect()

                    End If

                Next li_del_row
            End If
        Catch ex As Exception
            ErrorAppend(Filename + " " + ex.Message)
            Return False
        End Try
        Return True
    End Function

    Private Sub frmMain_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        ExitProgram()
    End Sub

    Sub SetLabelText(ByVal as_string As String)
        lblStatus.Text = as_string
    End Sub

    Private Function SendAlert(ByVal as_errormsg As String) As Boolean

        Try
            Dim oCompService As SAPbobsCOM.CompanyService
            Dim oMsgService As SAPbobsCOM.MessagesService
            Dim oMessage As SAPbobsCOM.Message

            Dim users As String
            Dim ctr As Integer = 0
            users = "'" + SAPB1UserIdtoMessage.Replace(",", "','") + "'"

            Dim dtUsers As DataTable = ExecuteReader("SELECT userid, USER_CODE, E_Mail  FROM OUSR WHERE USER_CODE IN (" + users + ")")

            If dtUsers.Rows.Count > 0 Then

                B1Connect(SAPB1CompanyName)

                oCompService = SAPB1Company.GetCompanyService()
                oMsgService = oCompService.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService)
                oMessage = oMsgService.GetDataInterface(SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage)

                oMessage.User = Convert.ToInt32(FormattedSearch("SELECT userid FROM OUSR WHERE USER_CODE='" + UserId + "'"))
                oMessage.Subject = "AR Posting Integration Fatal Error!"
                oMessage.Text = "Integration has encountered an fatal error while processing. Please see the error below " + vbNewLine + vbNewLine + as_errormsg + vbNewLine + vbNewLine + "COF Integration has been stopped."

                Dim cEmailAddress As SAPbobsCOM.RecipientCollection = oMessage.RecipientCollection


                For Each row In dtUsers.Rows
                    If row("E_Mail").ToString().Trim <> "" Then
                        cEmailAddress.Add()
                        cEmailAddress.Item(ctr).SendInternal = SAPbobsCOM.BoYesNoEnum.tYES
                        cEmailAddress.Item(ctr).SendEmail = SAPbobsCOM.BoYesNoEnum.tYES
                        cEmailAddress.Item(ctr).UserCode = row("USER_CODE").ToString()
                        cEmailAddress.Item(ctr).EmailAddress = row("E_Mail").ToString()
                        ctr += 1
                    End If
                Next


                Try
                    'Dim oMsgHead As SAPbobsCOM.MessageHeader
                    oMsgService.SendMessage(oMessage)
                Catch ex_alert As Exception
                    Throw New Exception(ex_alert.Message)
                End Try

            End If


        Catch ex As Exception
            ErrMsg = ex.Message
            Return False
        End Try

        Return True
    End Function


    'Function GetExcelColumnRowValue(ByVal as_xcelws As Excel.Worksheet, ByVal as_row As Integer, ByVal as_column As Integer) As String
    '    Try
    '        Return as_xcelws.Cells(as_row, as_column).Value.ToString().Trim()
    '    Catch ex As Exception

    '    End Try
    '    Return ""
    'End Function

    Function ReadTextFiles() As Boolean
        ErrMsg = ""
        Try
            Dim cofPath As New DirectoryInfo(FileOriginDump)
            Dim excelFiles As FileInfo() = cofPath.GetFiles("*.txt")
            Dim strfile, strFileName, strFileDumpDest As String
            Dim iFileTransferedCount As Integer = 0

            If excelFiles.Length > 0 Then
                For i As Integer = 0 To excelFiles.Length - 1
                    strFileName = excelFiles(i).ToString()
                    strfile = FileOriginDump + "\" + strFileName

                    If FileIsDone(FileOriginDump + "\" + strFileName) Then
                        If FileTransfer(strfile, FileDump + "\" + strFileName) Then
                            If File.Exists(strfile) Then
                                File.Delete(strfile)
                            End If
                            iFileTransferedCount += 1
                        Else
                            Throw New Exception(ErrMsg)
                        End If
                    Else
                        'Throw New Exception(ErrMsg)
                    End If
                Next

                If iFileTransferedCount > 0 Then
                    lblStatus.Text = iFileTransferedCount.ToString() + " file" + (IIf(iFileTransferedCount = 1, "", "s")) + " has been transfered for processing successfully!"
                    SuccessAppend(lblStatus.Text)
                End If
            End If

        Catch ex As Exception
            ErrMsg = ex.Message
            Return False
        End Try
        Return True
    End Function
    Private Sub timerReconnect_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles timerReconnect.Tick
        If Not ReConnect Then
            ReConnect = True
        End If
    End Sub
    Private Sub bgwInitialize_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwInitialize.RunWorkerCompleted
        lblStatus.Text = "Program Initialize Completed"
    End Sub
    Private Sub bgwIntegrate_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwIntegrate.RunWorkerCompleted
        bgwIntegrate.RunWorkerAsync()
    End Sub
    Private Sub timerStarter_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles timerStarter.Tick
        If lblStatus.Text = "Program Initialize Completed" Then
            lblStatus.Text = ""
            bgwIntegrate.RunWorkerAsync()
            timerStarter.Stop()
        End If
    End Sub


    Private Sub timerToday_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles timerToday.Tick
        lblToday.Text = DateTime.Now.ToString("MMMM dd, yyyy hh:mm:ss tt")
    End Sub

    Sub MemoryOptimizeStart()
        'Use of the class
        Dim listener As New Listener()
        Dim thread As New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf listener.StartListener))
        thread.Start()

    End Sub


    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub frmMain_GiveFeedback(sender As Object, e As GiveFeedbackEventArgs) Handles Me.GiveFeedback

    End Sub
End Class

Friend Class Listener

    Public Shared intervalInSeconds As Integer = 10
    Friend Sub StartListener()
        Try
            While True
                If intervalInSeconds > 0 Then
                    System.Threading.Thread.Sleep(intervalInSeconds * 1000)
                    FlushMemory()
                Else
                    System.Threading.Thread.Sleep(1000)
                    GC.Collect()
                End If
            End While
        Catch ex As Exception
            ErrorAppend(ex.Message)
        End Try
    End Sub


    <DllImport("kernel32.dll")>
    Private Shared Function SetProcessWorkingSetSize(ByVal process As IntPtr, ByVal minimumWorkingSetSize As Integer, ByVal maximumWorkingSetSize As Integer) As Integer
    End Function
    Public Shared Sub FlushMemory()
        GC.Collect()
        GC.WaitForPendingFinalizers()
        If Environment.OSVersion.Platform = PlatformID.Win32NT Then
            SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle, -1, -1)
        End If
    End Sub
End Class

