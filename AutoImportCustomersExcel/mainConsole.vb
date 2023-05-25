Imports Microsoft.Office.Interop
Imports MySql.Data.MySqlClient
Imports System.IO

Module mainConsole

    Dim vHost As String
    Dim vUsername As String
    Dim vPassword As String
    Dim vPort As String
    Dim vDatabase As String

    Sub Main()
        Radisson()
    End Sub

    Sub Knowles()

        Dim cn As clsConnectionDetails

        Try

            If IsDecrypted() Then

                Console.WriteLine("Connecting to database...")
                cn = New clsConnectionDetails
                cn.ConnectToServer(vHost, vUsername, vPassword, vPort, vDatabase)

                If cn.IsConnected Then

                    Console.WriteLine("Connected to " & UCase(vDatabase) & ".")
                    Dim mSystemParameters As New DataTable
                    Dim mCustomerFilePath As String = ""
                    Dim mPreviousFileDate As DateTime
                    Dim mCurrentFileDate As DateTime

                    'Date where import is not forced.
                    Dim mCurrentDate As DateTime
                    Dim mValidDate1 As DateTime
                    Dim mValidDate2 As DateTime

                    NetOpen(mSystemParameters, "select customer_file_path,customer_file_path_date,NOW() curdatetime   " & _
                                               "from system_parameters", cn.Connection)
                    If mSystemParameters.Rows.Count > 0 Then
                        For Each mRow As DataRow In mSystemParameters.Rows
                            mCustomerFilePath = mRow.Item("customer_file_path")
                            mPreviousFileDate = Format(CType(mRow.Item("customer_file_path_date"), DateTime), "MM/dd/yyyy HH:mm:ss")
                            mCurrentDate = Format(CType(mRow.Item("curdatetime"), DateTime), "MM/dd/yyyy HH:mm:ss")
                        Next
                    End If

                    mCurrentFileDate = Format(File.GetLastWriteTime(mCustomerFilePath), "MM/dd/yyyy HH:mm:ss")

                    mValidDate1 = Format(CType("2019-01-01 06:00:00", DateTime), "MM/dd/yyyy HH:mm:ss")
                    mValidDate2 = Format(CType("2019-01-07 06:00:00", DateTime), "MM/dd/yyyy HH:mm:ss")

                    If mCurrentDate <= mValidDate1 Or mCurrentDate >= mValidDate2 Then
                        If mCurrentFileDate = mPreviousFileDate Then
                            Console.WriteLine("No new file found!")
                            Exit Sub
                        End If
                        'Console.ReadLine()
                    End If

                    Dim xlApp As Excel.Application
                    Dim xlWorkbook As Excel.Workbook
                    Dim xlWorkSheet As Excel.Worksheet
                    Dim xlRange As Excel.Range

                    Dim xlRow As Integer

                    xlApp = New Excel.Application
                    xlWorkbook = xlApp.Workbooks.Open(mCustomerFilePath)
                    xlWorkSheet = xlWorkbook.ActiveSheet()
                    xlRange = xlWorkSheet.UsedRange

                    If xlRange.Columns.Count > 0 Then
                        If xlRange.Rows.Count > 0 Then

                            Dim mTrans As MySqlTransaction
                            Dim mCommand As New MySqlCommand
                            Dim rdQuery As MySqlDataReader

                            mTrans = cn.Connection.BeginTransaction
                            mCommand.Transaction = mTrans
                            mCommand.Connection = cn.Connection

                            Console.WriteLine("Importing customer information.")

                            mCommand.CommandText = "delete from customer_allowablecharges where chargetype_id in (1,2,3)"
                            mCommand.ExecuteNonQuery()

                            Dim isNew As Boolean = False

                            mCommand.Parameters.Add("@customer_id", MySqlDbType.Int32)
                            mCommand.Parameters.Add("@customer_code", MySqlDbType.String)
                            mCommand.Parameters.Add("@customer_rfcode", MySqlDbType.String)
                            mCommand.Parameters.Add("@customer_name", MySqlDbType.String)
                            mCommand.Parameters.Add("@subsidy", MySqlDbType.Decimal)
                            mCommand.Parameters.Add("@allowance", MySqlDbType.Decimal)
                            mCommand.Parameters.Add("@customergroup_id", MySqlDbType.Int32)
                            mCommand.Parameters.Add("@customersubgroup_id", MySqlDbType.Int32)

                            Dim mRowCtr As Integer
                            mRowCtr = 0

                            For xlRow = 2 To xlRange.Rows.Count

                                mRowCtr += 1
                                Console.Write("Status : " & Format(mRowCtr / xlRange.Rows.Count * 100, "###") & "%" & vbCr)
                                isNew = True
                                mCommand.Parameters.Item("@customer_code").Value = CType(xlRange.Cells(xlRow, 1).Text, String)
                                mCommand.Parameters.Item("@customer_rfcode").Value = CType(xlRange.Cells(xlRow, 2).Text, String)

                                If Not String.IsNullOrWhiteSpace(CType(xlRange.Cells(xlRow, 5).Text, String)) Then
                                    mCommand.Parameters.Item("@customer_name").Value = CType(xlRange.Cells(xlRow, 5).Text, String) & ", " & _
                                                CType(xlRange.Cells(xlRow, 3).Text, String) & " " & CType(xlRange.Cells(xlRow, 4).Text, String)
                                Else
                                    mCommand.Parameters.Item("@customer_name").Value = CType(xlRange.Cells(xlRow, 4).Text, String) & ", " & _
                                                CType(xlRange.Cells(xlRow, 3).Text, String)
                                End If

                                mCommand.CommandText = "select customer_id from customers where customer_code = @customer_code"
                                rdQuery = mCommand.ExecuteReader

                                If rdQuery.HasRows Then
                                    isNew = False
                                    While rdQuery.Read
                                        mCommand.Parameters.Item("@customer_id").Value = Convert.ToInt32(rdQuery.Item("customer_id"))
                                    End While
                                End If
                                rdQuery.Close()

                                '***Check for duplicate customer code
                                mCommand.CommandText = "select customergroup_id,customersubgroup_id from customersubgroups where customersubgroup_description = mid(@customer_code,1,1)"
                                rdQuery = mCommand.ExecuteReader

                                If rdQuery.HasRows Then
                                    While rdQuery.Read
                                        mCommand.Parameters.Item("@customergroup_id").Value = Convert.ToInt32(rdQuery.Item("customergroup_id"))
                                        mCommand.Parameters.Item("@customersubgroup_id").Value = Convert.ToInt32(rdQuery.Item("customersubgroup_id"))
                                    End While
                                Else
                                    mCommand.Parameters.Item("@customergroup_id").Value = 3
                                    mCommand.Parameters.Item("@customersubgroup_id").Value = 12
                                End If
                                rdQuery.Close()

                                If isNew Then
                                    mCommand.CommandText = "insert into customers (officeaddress,customer_name,customer_code,customer_rfcode,customergroup_id,customersubgroup_id) values " & _
                                                            "('',@customer_name,@customer_code,@customer_rfcode,@customergroup_id,@customersubgroup_id) "
                                    mCommand.ExecuteNonQuery()
                                    mCommand.Parameters.Item("@customer_id").Value = CType(mCommand.LastInsertedId, Integer)
                                Else
                                    mCommand.CommandText = "update customers set officeaddress='', customer_name=@customer_name, " & _
                                                            "customergroup_id=@customergroup_id,customersubgroup_id=@customersubgroup_id " & _
                                                            "where customer_id=@customer_id"
                                    'mCommand.CommandText = "update customers set officeaddress='', customer_name=@customer_name, " & _
                                    '                        "customergroup_id=@customergroup_id,customersubgroup_id=@customersubgroup_id, " & _
                                    '                        "active = 1 " & _
                                    '                        "where customer_id=@customer_id"
                                    mCommand.ExecuteNonQuery()
                                End If

                                '******Check for duplicate RF Codes
                                mCommand.CommandText = "select customer_rfcode from customer_rfcodes where customer_rfcode=@customer_rfcode"
                                rdQuery = mCommand.ExecuteReader
                                If rdQuery.HasRows Then
                                    isNew = False
                                Else
                                    isNew = True
                                End If
                                rdQuery.Close()

                                If isNew Then
                                    mCommand.CommandText = "insert into customer_rfcodes (customer_rfcode,customer_id) values " & _
                                                            "(@customer_rfcode,@customer_id) "
                                    mCommand.ExecuteNonQuery()
                                Else
                                    mCommand.CommandText = "update customer_rfcodes set customer_id=@customer_id where customer_rfcode=@customer_rfcode"
                                    mCommand.ExecuteNonQuery()
                                End If

                                'mCommand.CommandText = "delete from customer_allowablecharges where customer_id=@customer_id"
                                'mCommand.ExecuteNonQuery()

                                mCommand.Parameters.Item("@subsidy").Value = 0
                                mCommand.Parameters.Item("@allowance").Value = 0

                                If UCase(Mid(CType(xlRange.Cells(xlRow, 1).Text, String), 1, 1)) = "A" Then
                                    If mCurrentDate <= mValidDate1 Then
                                        mCommand.CommandText = "insert into customer_allowablecharges (customer_id,chargetype_id,limittype,minimum_amt,maximum_amt) values " & _
                                                                                                "(@customer_id,2,0,35,35)"
                                    Else
                                        mCommand.CommandText = "insert into customer_allowablecharges (customer_id,chargetype_id,limittype,minimum_amt,maximum_amt) values " & _
                                                                                                "(@customer_id,2,0,40,40)"
                                    End If
                                    mCommand.ExecuteNonQuery()

                                    mCommand.CommandText = "insert into customer_allowablecharges (customer_id,chargetype_id,limittype,minimum_amt,maximum_amt) values " & _
                                                            "(@customer_id,3,1,500,500)"
                                    mCommand.ExecuteNonQuery()

                                    'Temp
                                    'mCommand.CommandText = "insert into customer_allowablecharges (customer_id,chargetype_id,limittype,minimum_amt,maximum_amt) values " & _
                                    '                                                            "(@customer_id,5,0,40,40)"
                                    'mCommand.ExecuteNonQuery()

                                ElseIf UCase(Mid(CType(xlRange.Cells(xlRow, 1).Text, String), 1, 1)) = "B" Then
                                    If mCurrentDate <= mValidDate1 Then
                                        mCommand.CommandText = "insert into customer_allowablecharges (customer_id,chargetype_id,limittype,minimum_amt,maximum_amt) values " & _
                                                                                                "(@customer_id,2,0,35,35)"
                                    Else
                                        mCommand.CommandText = "insert into customer_allowablecharges (customer_id,chargetype_id,limittype,minimum_amt,maximum_amt) values " & _
                                                                                                "(@customer_id,2,0,40,40)"
                                    End If
                                    mCommand.ExecuteNonQuery()

                                    'Temp
                                    'mCommand.CommandText = "insert into customer_allowablecharges (customer_id,chargetype_id,limittype,minimum_amt,maximum_amt) values " & _
                                    '                                                            "(@customer_id,5,0,40,40)"
                                    'mCommand.ExecuteNonQuery()

                                ElseIf UCase(Mid(CType(xlRange.Cells(xlRow, 1).Text, String), 1, 1)) = "C" Then
                                    If mCurrentDate <= mValidDate1 Then
                                        mCommand.CommandText = "insert into customer_allowablecharges (customer_id,chargetype_id,limittype,minimum_amt,maximum_amt) values " & _
                                                                                                "(@customer_id,2,0,35,35)"
                                    Else
                                        mCommand.CommandText = "insert into customer_allowablecharges (customer_id,chargetype_id,limittype,minimum_amt,maximum_amt) values " & _
                                                                                                "(@customer_id,2,0,40,40)"
                                    End If
                                    mCommand.ExecuteNonQuery()

                                    'Temp
                                    'mCommand.CommandText = "insert into customer_allowablecharges (customer_id,chargetype_id,limittype,minimum_amt,maximum_amt) values " & _
                                    '                                                            "(@customer_id,5,0,40,40)"
                                    'mCommand.ExecuteNonQuery()

                                ElseIf UCase(Mid(CType(xlRange.Cells(xlRow, 1).Text, String), 1, 1)) = "E" Then
                                    mCommand.CommandText = "insert into customer_allowablecharges (customer_id,chargetype_id,limittype,minimum_amt,maximum_amt) values " & _
                                                        "(@customer_id,2,0,10,10)"
                                    mCommand.ExecuteNonQuery()

                                    'Temp
                                    'mCommand.CommandText = "insert into customer_allowablecharges (customer_id,chargetype_id,limittype,minimum_amt,maximum_amt) values " & _
                                    '                                                            "(@customer_id,5,0,40,40)"
                                    'mCommand.ExecuteNonQuery()

                                ElseIf UCase(Mid(CType(xlRange.Cells(xlRow, 1).Text, String), 1, 1)) = "J" Then

                                Else
                                    mCommand.CommandText = "insert into customer_allowablecharges (customer_id,chargetype_id,limittype,minimum_amt,maximum_amt) values " & _
                                                        "(@customer_id,2,0,35,35)"
                                    mCommand.ExecuteNonQuery()

                                    'Temp
                                    'mCommand.CommandText = "insert into customer_allowablecharges (customer_id,chargetype_id,limittype,minimum_amt,maximum_amt) values " & _
                                    '                                                            "(@customer_id,5,0,40,40)"
                                    'mCommand.ExecuteNonQuery()

                                End If

                                mCommand.CommandText = "update customer_rawdata set ischecked='Y' where customer_code=@customer_code and customer_rfcode=@customer_rfcode"
                                mCommand.ExecuteNonQuery()

                            Next

                            mCommand.CommandText = "update system_parameters set customer_file_path_date ='" & Format(mCurrentFileDate, "yyyy/MM/dd HH:mm:ss") & "'"
                            mCommand.ExecuteNonQuery()

                            mTrans.Commit()
                            Console.WriteLine("Customers imported succesfully.")

                            xlWorkbook.Close()
                            xlApp.Quit()

                            FileSystem.FileCopy(mCustomerFilePath, AppDomain.CurrentDomain.BaseDirectory & Format(mCurrentFileDate, "yyMMddhhmmss"))
                            'Console.ReadLine()
                        End If
                    End If

                Else
                    Console.WriteLine("Import Failed. mysql.exe file not found.")
                    'Console.ReadLine()
                End If
            Else
                Console.WriteLine("Unable to connect to database.")
                'Console.ReadLine()
            End If
        Catch ex As MySqlException
            Console.WriteLine("Import Failed! " & ex.Message)
            Console.ReadLine()
        Catch ex As Exception
            Console.WriteLine("Import Failed! " & ex.Message)
            Console.ReadLine()
        End Try
    End Sub

    Sub Autoliv()

        Dim cn As clsConnectionDetails

        Try

            If IsDecrypted() Then

                Console.WriteLine("Connecting to database...")
                cn = New clsConnectionDetails
                cn.ConnectToServer(vHost, vUsername, vPassword, vPort, vDatabase)

                If cn.IsConnected Then

                    Console.WriteLine("Connected to " & UCase(vDatabase) & ".")
                    Dim mSystemParameters As New DataTable
                    Dim mCustomerFilePath As String = ""
                    Dim mPreviousFileDate As DateTime
                    Dim mCurrentFileDate As DateTime

                    'Date where import is not forced.
                    Dim mCurrentDate As DateTime
                    Dim mValidDate1 As DateTime
                    Dim mValidDate2 As DateTime

                    NetOpen(mSystemParameters, "select customer_file_path,customer_file_path_date,NOW() curdatetime   " & _
                                               "from system_parameters", cn.Connection)
                    If mSystemParameters.Rows.Count > 0 Then
                        For Each mRow As DataRow In mSystemParameters.Rows
                            mCustomerFilePath = mRow.Item("customer_file_path")
                            mPreviousFileDate = Format(CType(mRow.Item("customer_file_path_date"), DateTime), "MM/dd/yyyy HH:mm:ss")
                            mCurrentDate = Format(CType(mRow.Item("curdatetime"), DateTime), "MM/dd/yyyy HH:mm:ss")
                        Next
                    End If

                    mCurrentFileDate = Format(File.GetLastWriteTime(mCustomerFilePath), "MM/dd/yyyy HH:mm:ss")

                    mValidDate1 = Format(CType("2019-01-01 06:00:00", DateTime), "MM/dd/yyyy HH:mm:ss")
                    mValidDate2 = Format(CType("2019-01-07 06:00:00", DateTime), "MM/dd/yyyy HH:mm:ss")

                    If mCurrentDate <= mValidDate1 Or mCurrentDate >= mValidDate2 Then
                        If mCurrentFileDate = mPreviousFileDate Then
                            Console.WriteLine("No new file found!")
                            Exit Sub
                        End If
                        'Console.ReadLine()
                    End If

                    Dim xlApp As Excel.Application
                    Dim xlWorkbook As Excel.Workbook
                    Dim xlWorkSheet As Excel.Worksheet
                    Dim xlRange As Excel.Range

                    Dim xlRow As Integer

                    xlApp = New Excel.Application
                    xlWorkbook = xlApp.Workbooks.Open(mCustomerFilePath)
                    xlWorkSheet = xlWorkbook.ActiveSheet()
                    xlRange = xlWorkSheet.UsedRange

                    If xlRange.Columns.Count > 0 Then
                        If xlRange.Rows.Count > 0 Then

                            Dim mTrans As MySqlTransaction
                            Dim mCommand As New MySqlCommand
                            Dim rdQuery As MySqlDataReader

                            mTrans = cn.Connection.BeginTransaction
                            mCommand.Transaction = mTrans
                            mCommand.Connection = cn.Connection

                            Console.WriteLine("Importing customer information.")

                            mCommand.CommandText = "DELETE FROM customer_allowablecharges WHERE customer_id IN (SELECT customer_id FROM customers WHERE customergroup_id = 2)"
                            mCommand.ExecuteNonQuery()

                            Dim isNew As Boolean = False

                            mCommand.Parameters.Add("@customer_id", MySqlDbType.Int32)
                            mCommand.Parameters.Add("@customer_code", MySqlDbType.String)
                            mCommand.Parameters.Add("@customer_name", MySqlDbType.String)
                            mCommand.Parameters.Add("@customersubgroup_name", MySqlDbType.String)
                            mCommand.Parameters.Add("@has_regular_charges", MySqlDbType.Decimal)
                            mCommand.Parameters.Add("@has_regular_subsidy", MySqlDbType.Decimal)
                            mCommand.Parameters.Add("@has_ot_subsidy", MySqlDbType.Decimal)
                            mCommand.Parameters.Add("@customergroup_id", MySqlDbType.Int32)
                            mCommand.Parameters.Add("@customersubgroup_id", MySqlDbType.Int32)

                            Dim mRowCtr As Integer
                            Dim defaultCustomergroupId As String
                            Dim defaultCustomersubgroupId As String
                            Dim hasRegularCharges As String
                            Dim hasRegularSubsidy As String
                            Dim hasOtSubsidy As String
                            Dim regularCharges As Double
                            Dim regularSubsidy As Double
                            Dim otSubsidy As Double

                            mRowCtr = 0
                            regularCharges = 1000
                            regularSubsidy = 30
                            otSubsidy = 20
                            defaultCustomergroupId = 2
                            defaultCustomersubgroupId = 18

                            For xlRow = 4 To xlRange.Rows.Count

                                mRowCtr += 1
                                Console.Write("Status : " & Format(mRowCtr / xlRange.Rows.Count * 100, "###") & "%" & vbCr)
                                isNew = True

                                mCommand.Parameters.Item("@customer_code").Value = Trim(CType(xlRange.Cells(xlRow, 3).Text, String))
                                mCommand.Parameters.Item("@customersubgroup_name").Value = Trim(CType(xlRange.Cells(xlRow, 8).Text, String))
                                mCommand.Parameters.Item("@customer_name").Value = Trim(CType(xlRange.Cells(xlRow, 4).Text, String)) & ", " & _
                                               Trim(CType(xlRange.Cells(xlRow, 5).Text, String)) & " " & Trim(CType(xlRange.Cells(xlRow, 6).Text, String)) & " " & _
                                               Trim(CType(xlRange.Cells(xlRow, 7).Text, String))

                                hasRegularCharges = Trim(CType(xlRange.Cells(xlRow, 11).Text, String))
                                hasRegularSubsidy = Trim(CType(xlRange.Cells(xlRow, 9).Text, String))
                                hasOtSubsidy = Trim(CType(xlRange.Cells(xlRow, 10).Text, String))

                                mCommand.CommandText = "select customer_id from customers where customer_code = @customer_code"
                                rdQuery = mCommand.ExecuteReader

                                '***Check for duplicate customer code
                                If rdQuery.HasRows Then
                                    isNew = False
                                    While rdQuery.Read
                                        mCommand.Parameters.Item("@customer_id").Value = Convert.ToInt32(rdQuery.Item("customer_id"))
                                    End While
                                End If
                                rdQuery.Close()

                                mCommand.CommandText = "select customergroup_id,customersubgroup_id from customersubgroups where customersubgroup_name = @customersubgroup_name"
                                rdQuery = mCommand.ExecuteReader

                                If rdQuery.HasRows Then
                                    While rdQuery.Read
                                        mCommand.Parameters.Item("@customergroup_id").Value = Convert.ToInt32(rdQuery.Item("customergroup_id"))
                                        mCommand.Parameters.Item("@customersubgroup_id").Value = Convert.ToInt32(rdQuery.Item("customersubgroup_id"))
                                    End While
                                Else
                                    mCommand.Parameters.Item("@customergroup_id").Value = defaultCustomergroupId
                                    mCommand.Parameters.Item("@customersubgroup_id").Value = defaultCustomersubgroupId
                                End If
                                rdQuery.Close()

                                If isNew Then
                                    mCommand.CommandText = "insert into customers (officeaddress,customer_name,customer_code,customergroup_id,customersubgroup_id) values " & _
                                                            "('',@customer_name,@customer_code,@customergroup_id,@customersubgroup_id) "
                                    mCommand.ExecuteNonQuery()
                                    mCommand.Parameters.Item("@customer_id").Value = CType(mCommand.LastInsertedId, Integer)
                                Else
                                    mCommand.CommandText = "update customers set officeaddress='', customer_name=@customer_name, " & _
                                                            "customergroup_id=@customergroup_id,customersubgroup_id=@customersubgroup_id " & _
                                                            "where customer_id=@customer_id"
                                    mCommand.ExecuteNonQuery()
                                End If

                                mCommand.CommandText = "DELETE FROM customer_allowablecharges WHERE customer_id = @customer_id"
                                mCommand.ExecuteNonQuery()


                                If hasRegularCharges = "Y" Then
                                    mCommand.CommandText = "insert into customer_allowablecharges (customer_id,chargetype_id,limittype,minimum_amt,maximum_amt) values " & _
                                                                                                "(@customer_id,1,1," & regularCharges & "," & regularCharges & ")"
                                    mCommand.ExecuteNonQuery()
                                End If

                                If hasRegularSubsidy = "Y" Then
                                    mCommand.CommandText = "insert into customer_allowablecharges (customer_id,chargetype_id,limittype,minimum_amt,maximum_amt) values " & _
                                                                                                "(@customer_id,2,0," & regularSubsidy & "," & regularSubsidy & ")"
                                    mCommand.ExecuteNonQuery()
                                End If

                                If hasOtSubsidy = "Y" Then
                                    mCommand.CommandText = "insert into customer_allowablecharges (customer_id,chargetype_id,limittype,minimum_amt,maximum_amt) values " & _
                                                                                                "(@customer_id,3,0," & otSubsidy & "," & otSubsidy & ")"
                                    mCommand.ExecuteNonQuery()
                                End If

                            Next

                            mCommand.CommandText = "update system_parameters set customer_file_path_date ='" & Format(mCurrentFileDate, "yyyy/MM/dd HH:mm:ss") & "'"
                            mCommand.ExecuteNonQuery()

                            mTrans.Commit()
                            Console.WriteLine("Customers imported succesfully.")

                            xlWorkbook.Close()
                            xlApp.Quit()

                            FileSystem.FileCopy(mCustomerFilePath, AppDomain.CurrentDomain.BaseDirectory & Format(mCurrentFileDate, "yyMMddhhmmss"))
                            'Console.ReadLine()
                        End If
                    End If

                Else
                    Console.WriteLine("Import Failed. mysql.exe file not found.")
                    'Console.ReadLine()
                End If
            Else
                Console.WriteLine("Unable to connect to database.")
                'Console.ReadLine()
            End If
        Catch ex As MySqlException
            Console.WriteLine("Import Failed! " & ex.Message)
            Console.ReadLine()
        Catch ex As Exception
            Console.WriteLine("Import Failed! " & ex.Message)
            Console.ReadLine()
        End Try
    End Sub

    Sub Radisson()

        Dim cn As clsConnectionDetails

        Try

            If IsDecrypted() Then

                Console.WriteLine("Connecting to database...")
                cn = New clsConnectionDetails
                cn.ConnectToServer(vHost, vUsername, vPassword, vPort, vDatabase)

                If cn.IsConnected Then

                    Console.WriteLine("Connected to " & UCase(vDatabase) & ".")
                    Dim mSystemParameters As New DataTable
                    Dim mCustomerFilePath As String = ""
                    Dim mPreviousFileDate As DateTime
                    Dim mCurrentFileDate As DateTime

                    'Date where import is not forced.
                    Dim mCurrentDate As DateTime
                    Dim mValidDate1 As DateTime
                    Dim mValidDate2 As DateTime

                    NetOpen(mSystemParameters, "select customer_file_path,customer_file_path_date,NOW() curdatetime   " & _
                                               "from system_parameters", cn.Connection)
                    If mSystemParameters.Rows.Count > 0 Then
                        For Each mRow As DataRow In mSystemParameters.Rows
                            mCustomerFilePath = mRow.Item("customer_file_path")
                            mPreviousFileDate = Format(CType(mRow.Item("customer_file_path_date"), DateTime), "MM/dd/yyyy HH:mm:ss")
                            mCurrentDate = Format(CType(mRow.Item("curdatetime"), DateTime), "MM/dd/yyyy HH:mm:ss")
                        Next
                    End If

                    mCurrentFileDate = Format(File.GetLastWriteTime(mCustomerFilePath), "MM/dd/yyyy HH:mm:ss")

                    mValidDate1 = Format(CType("2019-01-01 06:00:00", DateTime), "MM/dd/yyyy HH:mm:ss")
                    mValidDate2 = Format(CType("2019-01-07 06:00:00", DateTime), "MM/dd/yyyy HH:mm:ss")

                    If mCurrentDate <= mValidDate1 Or mCurrentDate >= mValidDate2 Then
                        If mCurrentFileDate = mPreviousFileDate Then
                            Console.WriteLine("No new file found!")
                            Exit Sub
                        End If
                        'Console.ReadLine()
                    End If

                    Dim xlApp As Excel.Application
                    Dim xlWorkbook As Excel.Workbook
                    Dim xlWorkSheet As Excel.Worksheet
                    Dim xlRange As Excel.Range

                    Dim xlRow As Integer

                    xlApp = New Excel.Application
                    xlWorkbook = xlApp.Workbooks.Open(mCustomerFilePath)
                    xlWorkSheet = xlWorkbook.ActiveSheet()
                    xlRange = xlWorkSheet.UsedRange

                    If xlRange.Columns.Count > 0 Then
                        If xlRange.Rows.Count > 0 Then

                            Dim mTrans As MySqlTransaction
                            Dim mCommand As New MySqlCommand
                            Dim rdQuery As MySqlDataReader

                            mTrans = cn.Connection.BeginTransaction
                            mCommand.Transaction = mTrans
                            mCommand.Connection = cn.Connection

                            Console.WriteLine("Importing customer information.")

                            mCommand.CommandText = "DELETE FROM customer_allowablecharges WHERE customer_id IN (SELECT customer_id FROM customers WHERE customergroup_id >= 2)"
                            mCommand.ExecuteNonQuery()

                            Dim isNew As Boolean = False

                            mCommand.Parameters.Add("@customer_id", MySqlDbType.Int32)
                            mCommand.Parameters.Add("@customer_code", MySqlDbType.String)
                            mCommand.Parameters.Add("@customer_name", MySqlDbType.String)
                            mCommand.Parameters.Add("@customergroup_id", MySqlDbType.Int32)
                            mCommand.Parameters.Add("@customersubgroup_id", MySqlDbType.Int32)
                            mCommand.Parameters.Add("@customersubgroup_name", MySqlDbType.String)

                            Dim mRowCtr As Integer
                            Dim customerCode As String
                            Dim customerGroupId As String
                            Dim undefinedCustomerSubgroupId As String
                            Dim regularSubsidy As Double

                            mRowCtr = 0

                            regularSubsidy = 75

                            For xlRow = 5 To xlRange.Rows.Count

                                mRowCtr += 1
                                Console.Write("Status : " & Format(mRowCtr / xlRange.Rows.Count * 100, "###") & "%" & vbCr)
                                isNew = True

                                If Trim(CType(xlRange.Cells(xlRow, 1).Text, String)) <> "" Then

                                    customerCode = Trim(CType(xlRange.Cells(xlRow, 1).Text, String))

                                    If customerCode.Substring(0, 4).ToUpper = "MCRI" Then
                                        customerGroupId = 3
                                        undefinedCustomerSubgroupId = 75
                                    ElseIf customerCode.Substring(0, 4).ToUpper = "CGSI" Then
                                        customerGroupId = 4
                                        undefinedCustomerSubgroupId = 76
                                    Else
                                        customerGroupId = 2
                                        undefinedCustomerSubgroupId = 74
                                    End If

                                    mCommand.Parameters.Item("@customer_code").Value = customerCode
                                    mCommand.Parameters.Item("@customergroup_id").Value = customerGroupId
                                    mCommand.Parameters.Item("@customersubgroup_name").Value = Trim(CType(xlRange.Cells(xlRow, 4).Text, String))

                                    If Trim(CType(xlRange.Cells(xlRow, 2).Text, String)) = "" Then
                                        mCommand.Parameters.Item("@customer_name").Value = Trim(CType(xlRange.Cells(xlRow, 1).Text, String))
                                    Else
                                        mCommand.Parameters.Item("@customer_name").Value = Trim(CType(xlRange.Cells(xlRow, 2).Text, String)) & ", " & _
                                                       Trim(CType(xlRange.Cells(xlRow, 3).Text, String))
                                    End If

                                    mCommand.CommandText = "select customer_id from customers where customer_code = @customer_code"
                                    rdQuery = mCommand.ExecuteReader

                                    '***Check for duplicate customer code
                                    If rdQuery.HasRows Then
                                        isNew = False
                                        While rdQuery.Read
                                            mCommand.Parameters.Item("@customer_id").Value = Convert.ToInt32(rdQuery.Item("customer_id"))
                                        End While
                                    End If
                                    rdQuery.Close()

                                    mCommand.CommandText = "select customersubgroup_id from customersubgroups " & _
                                                            "where customersubgroup_name = @customersubgroup_name and customergroup_id=@customergroup_id"
                                    rdQuery = mCommand.ExecuteReader

                                    If rdQuery.HasRows Then
                                        While rdQuery.Read
                                            mCommand.Parameters.Item("@customersubgroup_id").Value = Convert.ToInt32(rdQuery.Item("customersubgroup_id"))
                                        End While
                                    Else
                                        mCommand.Parameters.Item("@customersubgroup_id").Value = undefinedCustomerSubgroupId
                                    End If
                                    rdQuery.Close()

                                    If isNew Then
                                        mCommand.CommandText = "insert into customers (officeaddress,customer_name,customer_code,customergroup_id,customersubgroup_id) values " & _
                                                                "('',@customer_name,@customer_code,@customergroup_id,@customersubgroup_id) "
                                        mCommand.ExecuteNonQuery()
                                        mCommand.Parameters.Item("@customer_id").Value = CType(mCommand.LastInsertedId, Integer)
                                    Else
                                        mCommand.CommandText = "update customers set officeaddress='', customer_name=@customer_name, " & _
                                                                "customergroup_id=@customergroup_id,customersubgroup_id=@customersubgroup_id " & _
                                                                "where customer_id=@customer_id"
                                        mCommand.ExecuteNonQuery()
                                    End If

                                    mCommand.CommandText = "insert into customer_allowablecharges (customer_id,chargetype_id,limittype,minimum_amt,maximum_amt) values " & _
                                                                                                "(@customer_id,2,0," & regularSubsidy & "," & regularSubsidy & ")"
                                    mCommand.ExecuteNonQuery()

                                End If

                            Next

                            mCommand.CommandText = "update system_parameters set customer_file_path_date ='" & Format(mCurrentFileDate, "yyyy-MM-dd HH:mm:ss") & "'"
                            mCommand.ExecuteNonQuery()

                            mTrans.Commit()
                            Console.WriteLine("Customers imported succesfully.")

                            xlWorkbook.Close()
                            xlApp.Quit()

                            FileSystem.FileCopy(mCustomerFilePath, AppDomain.CurrentDomain.BaseDirectory & Format(mCurrentFileDate, "yyMMddhhmmss"))
                            'Console.ReadLine()
                        End If
                    End If

                Else
                    Console.WriteLine("Import Failed. mysql.exe file not found.")
                    'Console.ReadLine()
                End If
            Else
                Console.WriteLine("Unable to connect to database.")
                'Console.ReadLine()
            End If
        Catch ex As MySqlException
            Console.WriteLine("Import Failed! " & ex.Message)
            Console.ReadLine()
        Catch ex As Exception
            Console.WriteLine("Import Failed! " & ex.Message)
            Console.ReadLine()
        End Try
    End Sub

    Private Function IsDecrypted() As Boolean
        Dim wrapper As New clsSimple3Des("un1quep@ssw0rd")
        Try
            Dim cipherText As String = My.Computer.FileSystem.ReadAllText(My.Computer.FileSystem.CurrentDirectory & "\strct.dat")
            Dim plainText As String = wrapper.DecryptData(cipherText)
            Call RetrieveText(plainText) 'rearrange data to match the connection settings
            Return True
        Catch ex As System.Security.Cryptography.CryptographicException
            Console.WriteLine("Unable to load credentials.")
            Return False
        Catch ex As Exception
            Console.WriteLine("An error occured! " & ex.Message)
            Return False
        End Try
    End Function

    Private Sub RetrieveText(ByVal vString As String)
        Dim ctr As Integer
        Dim i As Integer
        Dim charsInFile As Integer
        Dim letter As String
        Dim Decrypt As String = ""
        vHost = ""
        vUsername = ""
        vPassword = ""
        vPort = ""
        vDatabase = ""
        charsInFile = vString.Length 'get length of string
        For i = 0 To charsInFile - 1 'loop once for each char
            letter = vString.Substring(i, 1) 'get character
            If letter = "|" Then
                letter = ""
                ctr = ctr + 1
                If ctr = 1 Then
                    vHost = Decrypt
                    Decrypt = ""
                ElseIf ctr = 2 Then
                    vUsername = Decrypt
                    Decrypt = ""
                ElseIf ctr = 3 Then
                    vPassword = Decrypt
                    Decrypt = ""
                ElseIf ctr = 4 Then
                    vPort = Decrypt
                    Decrypt = ""
                End If
            End If
            Decrypt = Decrypt & letter
        Next i 'and build new string
        vDatabase = Decrypt
        Decrypt = ""
    End Sub

End Module
