Option Strict On

Imports MySql.Data.MySqlClient
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Net.Mail
Imports System.Xml

Module Catalog_Export
    Const CONNECTIONSTRING As String = "Database=magento;Data Source=localhost;" _
            & "User Id=sd_mage_user;Password=3D4U7p1a0hba82"
    Const KEYASCII As String = "79dfed1eda674abf713e36166e097113"
    Dim productNumber, mailTo, SMTPAddress As String
    Dim successfulOrders As New ArrayList

    Sub Main()
        readConfig()
        Dim fileTime As DateTime = DateTime.Now
        'Dim exportFileName As String = FILEPATH & "SERVU-" & fileTime.ToString("yyyyMMddHHmm") & ".txt"
        Dim exportFileName As String = "CATALOG-EXPORT-SERVU-" & fileTime.ToString("yyyyMMddHHmm") & ".txt"
        Dim catalogConnection As New MySqlConnection(CONNECTIONSTRING)
        Dim billingAndShipping As New MySqlConnection(CONNECTIONSTRING)
        Dim shippingNumbers As New MySqlConnection(CONNECTIONSTRING)
        Dim currentRequest, previousRequest As Integer
        Dim exportOutput As String = ""
        Dim customerType As String

        previousRequest = -1

        Try
            catalogConnection.Open()
            Dim catalogSQLStatement As String = "SELECT * FROM catalogrequest WHERE status = 1 "
            Dim catalogSQLCommand As MySqlCommand = New MySqlCommand(catalogSQLStatement, catalogConnection)
            Dim catalogResultReader As MySqlDataReader = catalogSQLCommand.ExecuteReader()

            While catalogResultReader.Read()
                'Create order header
                currentRequest = catalogResultReader.GetInt16(0)

                'this is a new request so we must create the order header
                previousRequest = currentRequest
                'remember to split out the address

                Dim customerInformation As New Dictionary(Of String, String)
                Dim orderDateString As String = IfNull(catalogResultReader, "time_added", "01/01/2012")
                Dim orderDate As DateTime = Convert.ToDateTime(orderDateString)
                Dim formatedDate As String = orderDate.ToString("MM/dd/yy")

                customerInformation = CType(createCustomerInfo(catalogResultReader), Global.System.Collections.Generic.Dictionary(Of String, String))

                exportOutput &= "A" & CType(createField(customerInformation("telephone"), 28), String)
                exportOutput &= CType(createField(customerInformation("company"), 40), String)
                exportOutput &= CType(createField(customerInformation("firstName"), 12), String)
                exportOutput &= CType(createField(customerInformation("lastName"), 17), String)
                exportOutput &= CType(createField(customerInformation("addressOne"), 35), String)
                exportOutput &= CType(createField(customerInformation("addressTwo"), 35), String)
                'City field has 3 other filler fields with it, see solutions documentation
                exportOutput &= CType(createField(customerInformation("city"), 46), String)
                exportOutput &= CType(createField(customerInformation("state"), 25), String)
                exportOutput &= CType(createField(customerInformation("country"), 25), String)
                exportOutput &= CType(createField(customerInformation("zip"), 10), String)
                '#TODO: This is the media code field that needs to be implemented in Magento
                exportOutput &= CType(createField(customerInformation("heardofus"), 8), String)
                '#TODO: Orderdate also uses ck bank name(40), Bank City(30), Ck Num(6), ck Account Num(53)
                exportOutput &= CType(createField(formatedDate, 137), String)
                'Fax also uses fax_additional(10), srvc bur commnt(65),  dflt_csr_id(4), CUST_TYPE(8)
                'batch(4), ctsy_title(10)
                exportOutput &= CType(createField("", 111), String)
                '#TODO: Commercial field to be implemented later Values are "C" and "R" for commercial
                'and residential. Length is 1 the actual values are 1 for commercial and 2 for residential

                If (customerInformation("res_bus") = "business") Then
                    customerType = "C"
                ElseIf (customerInformation("res_bus") = "residential") Then
                    customerType = "R"
                Else
                    customerType = ""
                End If
                exportOutput &= CType(createField(customerType, 1), String) & vbCrLf

                '00.00 is the Discount_perc field with a length of 4
                exportOutput &= "B" & "00.00"
                'VI, MC, DIS, AX, CHK, COD
                exportOutput &= CType(createField("CHK", 4), String)


                'credit card # & expiration blank for catalog request
                exportOutput &= CType(createField("", 21), String)
                exportOutput &= CType(createField("", 8), String)
                'Expiration also contains approval code(8), customer folloup(1) Y=Customer service followup
                'or catalog request No record type = C or D, and installment flag(1)Y or N
                exportOutput &= CType(createField("", 1), String)
                exportOutput &= CType(createField("", 1), String)
                'Contains Defered Billing(1) and  AUTH_DATE(8)
                exportOutput &= CType(createField(catalogResultReader, "email", "", 44), String)
                exportOutput &= CType(createField(" ", 35), String)
                exportOutput &= CType(createField(customerInformation("addressOne"), 35), String)
                exportOutput &= CType(createField(customerInformation("zip"), 10), String)
                '#TODO: This next field encompasses Invoice Number(10), and CID, CVV Number(3)'
                exportOutput &= CType(createField(" ", 13), String)
                exportOutput &= CType(createField(catalogResultReader, "catalogrequest_id", "NOCAT#", 19), String)

                exportOutput &= CType(createField("IP Address: " & IfNull(catalogResultReader, "ip", "") & " Product Interest: " & IfNull(catalogResultReader, "product_interest", ""), 750), String) & vbCrLf
                exportOutput &= "C" & CType(createField(customerInformation("telephone"), 10), String)
                exportOutput &= CType(createField(customerInformation("company"), 40), String)
                exportOutput &= CType(createField(customerInformation("firstName"), 12), String)
                exportOutput &= CType(createField(customerInformation("lastName"), 17), String)
                exportOutput &= CType(createField(customerInformation("addressOne"), 35), String)
                exportOutput &= CType(createField(customerInformation("addressTwo"), 35), String)
                exportOutput &= CType(createField(customerInformation("city"), 46), String)
                exportOutput &= CType(createField(customerInformation("state"), 25), String)
                exportOutput &= CType(createField(customerInformation("country"), 25), String)
                exportOutput &= CType(createField(customerInformation("zip"), 10), String)
                'This is tax rate
                exportOutput &= CType(createField("00.00", 5), String)
                exportOutput &= CType(createField("00.00", 5), String)
                'This is misc amt
                exportOutput &= CType(createField("00000.00", 8), String)
                'This is handling amt
                exportOutput &= CType(createField("000.00", 6), String)
                'Gift Message
                exportOutput &= CType(createField("", 179), String)
                'PO Number from Web Number
                exportOutput &= CType(createField(catalogResultReader, "catalogrequest_id", "NOCAT#", 15), String)
                'Ctsy Title
                exportOutput &= CType(createField("", 10), String)
                exportOutput &= CType(createField(catalogResultReader, "email", "", 35), String)
                'Print hard copy
                exportOutput &= CType(createField("Y", 1), String) & vbCrLf

                exportOutput &= CType(buildLineItem(), String)

                successfulOrders.Add(catalogResultReader.GetInt32("catalogrequest_id"))

            End While
            catalogResultReader.Close()

        Catch e As MySqlException
            LogError(DateTime.Now & "Error in order connection/query: " & e.ToString())
        Finally
            catalogConnection.Close()
        End Try

        If successfulOrders.Count <> 0 Then
            sendExportResults("Catalog Export: " & exportFileName, exportFileName & " exported successfully.", exportFileName, exportOutput)
            updateRequests()
        Else
            sendExportResults("Catalog Export: No requests", exportFileName & " There were no catalog requests to export.", exportFileName, "")
        End If

        'Threading.Thread.Sleep(100000)

    End Sub

    Sub readConfig()
        productNumber = "BNO"
        SMTPAddress = "192.168.1.50"
        mailTo = "travishill@servu-online.com"

        Try
            Dim doc As New System.Xml.XmlDocument
            doc.Load(My.Application.Info.DirectoryPath & "\config.xml")
            Dim list = doc.GetElementsByTagName("name")

            If (doc.GetElementsByTagName("SMTP")(0).InnerText <> "") Then
                SMTPAddress = doc.GetElementsByTagName("SMTP")(0).InnerText
            End If

            If (doc.GetElementsByTagName("MailTo")(0).InnerText <> "") Then
                mailTo = doc.GetElementsByTagName("MailTo")(0).InnerText
            End If

            If (doc.GetElementsByTagName("CatalogNumber")(0).InnerText <> "") Then
                productNumber = doc.GetElementsByTagName("CatalogNumber")(0).InnerText
            End If
        Catch e As Exception
            LogError("Error in xml: " & e.ToString())
        End Try


    End Sub

    Sub sendExportResults(ByVal subject As String, ByVal body As String, ByVal exportFileName As String, ByVal exportFile As String)
        Dim mail As New MailMessage()
        Dim smtp As New SmtpClient(SMTPAddress)

        mail.From = New MailAddress(mailTo)
        mail.To.Add(mailTo)

        mail.Subject = subject
        mail.Body = body

        If exportFile <> "" Then
            Dim ms As MemoryStream = New MemoryStream()
            Dim sw As StreamWriter = New StreamWriter(ms)
            sw.WriteLine(exportFile)
            sw.Flush()
            ms.Position = 0
            mail.Attachments.Add(New Attachment(ms, exportFileName, "text/plain"))
        End If

        smtp.Send(mail)
    End Sub

    Private Sub updateRequests()
        Dim conn As New MySqlConnection(CONNECTIONSTRING)
        Dim cmd As New MySqlCommand()
        Dim transaction As MySqlTransaction
        Dim whereClause, sqlStatement As String

        whereClause = String.Join(" OR catalogrequest.catalogrequest_id = ", successfulOrders.ToArray())
        sqlStatement = "UPDATE catalogrequest SET catalogrequest.status = 2 WHERE catalogrequest.catalogrequest_id = " & whereClause

        Try
            conn.Open()
            transaction = conn.BeginTransaction()

            cmd.Connection = conn
            cmd.Transaction = transaction

            cmd.CommandText = sqlStatement
            cmd.ExecuteNonQuery()

            transaction.Commit()
            conn.Close()

        Catch e As MySqlException
            transaction.Rollback()
            LogError(DateTime.Now & "Error in commiting update: " & e.ToString())
        End Try

    End Sub

    Private Function buildLineItem() As String
        Dim lineItem, comments As String
        Dim importableAttributes As New ArrayList
        comments = ""
        lineItem = ""

        lineItem = "D" & CType(createField(productNumber, 20), String)
        lineItem &= CType(createField("", 20), String)

        lineItem &= CType(createField("1", 5), String)
        lineItem &= CType(createField("0000.00", 7), String)
        'Gift Wrap
        lineItem &= CType(createField("N", 1), String)
        'Gift wrap fee
        lineItem &= CType(createField("00.00", 5), String)
        'Giftwrap message, contains variable_kit(1), attribute 3(10), attribute 4(10), attribute 5(10)
        lineItem &= CType(createField(" ", 320), String)
        lineItem &= CType(createField(" ", 31), String)
        lineItem &= CType(createField(" ", 74), String) & vbCrLf

        Return lineItem

    End Function

    Private Function createCustomerInfo(ByVal dr As MySqlDataReader) As System.Collections.Generic.Dictionary(Of String, String)
        'Create array for shipping and billing information
        Dim returnValues As New Dictionary(Of String, String)

        returnValues("firstName") = IfNull(dr, "fname", "")
        returnValues("lastName") = IfNull(dr, "lname", "")
        returnValues("city") = IfNull(dr, "city", "")
        returnValues("state") = IfNull(dr, "state", "")
        returnValues("zip") = IfNull(dr, "zip", "")
        returnValues("email") = IfNull(dr, "email", "")
        returnValues("telephone") = IfNull(dr, "phone", "")
        returnValues("company") = IfNull(dr, "company", "")
        returnValues("country") = IfNull(dr, "country", "")
        If (returnValues("country")) = "US" Then
            returnValues("country") = "USA"
        End If

        returnValues("addressOne") = IfNull(dr, "address1", "")
        returnValues("addressTwo") = IfNull(dr, "address2", "")
        returnValues("heardofus") = IfNull(dr, "heardofus", "")
        returnValues("res_bus") = IfNull(dr, "res_bus", "")

        Return returnValues

    End Function

    Private Sub LogError(ByVal e As String)
        'Utility function to log errors
        Dim fileName As String = My.Application.Info.DirectoryPath & "\logs\log.txt"

        If File.Exists(fileName) Then
            Using fileWriter As StreamWriter = New StreamWriter(fileName, True)
                fileWriter.Write(e & vbTab)
            End Using
        Else
            Using fileWriter As StreamWriter = New StreamWriter(fileName)
                fileWriter.Write(e & vbTab)
            End Using
        End If

    End Sub

    Public Function IfNull(Of T)(ByVal dr As MySqlDataReader, ByVal fieldName As String, ByVal _default As T) As T

        If IsDBNull(dr(fieldName)) Then
            Return _default
        Else
            Return CType(dr(fieldName), T)
        End If

    End Function

    Private Function createField(ByVal fieldValue As String, ByVal totalSpaces As Integer) As String

        If fieldValue.Length > totalSpaces Then
            Return fieldValue.Substring(0, totalSpaces)
        Else
            Return fieldValue.PadRight(totalSpaces, CChar(" "))
        End If

    End Function

    Private Function createField(Of T)(ByVal dr As MySqlDataReader, ByVal fieldName As String, ByVal _default As T, _
                                    ByVal totalSpaces As Integer) As String
        Dim fieldValue As String

        fieldValue = IfNull(dr, fieldName, _default).ToString

        If fieldValue.Length > totalSpaces Then
            Return fieldValue.Substring(0, totalSpaces)
        Else
            Return fieldValue.PadRight(totalSpaces, CChar(" "))
        End If

    End Function

End Module
