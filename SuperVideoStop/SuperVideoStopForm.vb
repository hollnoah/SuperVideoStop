'Noah Holloway
'Spring 2025
'RCET 2265
'Super Video Stop

Option Strict On
Option Explicit On
Option Compare Text
Imports System.IO

Public Class SuperVideoStopForm

    Function Customers(Optional customerData(,) As String = Nothing) As String(,)
        Static _customers(,) As String

        If customerData IsNot Nothing Then
            _customers = customerData

        End If
        Return _customers
    End Function
    Sub ReadFromFile()
        Dim filePath As String = "..\..\..\UserData.txt"
        Dim fileNumber As Integer = FreeFile()
        Dim currentRecord As String = ""
        Dim temp() As String 'use for splitting customer data
        Dim currentID As Integer = 699


        Try
            FileOpen(fileNumber, filePath, OpenMode.Input)

            Do Until EOF(fileNumber)
                Input(fileNumber, currentRecord) 'read exacrly one record
                If currentRecord <> "" Then 'ignore blank records
                    temp = Split(currentRecord, ",")
                    'DisplayListBox.Items.Add(currentRecord) 'add the record to the list box

                    If temp.Length = 4 Then 'ignore malformed records
                        temp(0) = Replace(temp(0), "$", "") 'cleaning the first name
                        DisplayListBox.Items.Add(temp(0))
                        WriteToFile(temp(0)) 'first name
                        WriteToFile(temp(1)) 'last name
                        WriteToFile("") 'place holder for street
                        WriteToFile(temp(2)) 'city
                        WriteToFile("ID") 'place holder for state
                        WriteToFile("") 'place holder for zip
                        WriteToFile("") 'place holder for phone
                        WriteToFile("") 'place holder for email
                        WriteToFile(temp(3), True)
                        WriteToFile("") 'place holder 
                        WriteToFile($"0631{currentID}", True)
                        currentID += 1

                    End If
                End If
            Loop

            FileClose(fileNumber)
        Catch bob As FileNotFoundException
            MsgBox("Bob is sad...")

        Catch ex As Exception
            MsgBox(ex.Message & vbNewLine & ex.StackTrace & vbNewLine)

        End Try

    End Sub

    Sub WriteToFile(newRecord As String, Optional insertNewLine As Boolean = False)
        Dim filePath As String = "..\..\..\CustomerData.txt"
        Dim fileNumber As Integer = FreeFile()

        Try
            FileOpen(fileNumber, filePath, OpenMode.Append)
            Write(fileNumber, newRecord)
            If insertNewLine Then
                WriteLine(fileNumber)
            End If
            FileClose(fileNumber)

        Catch ex As Exception
            MsgBox($"Error writing to {filePath}")
        End Try


    End Sub

    Sub LoadCustomerData()
        Dim filePath As String = "..\..\..\CustomerData.dat"
        Dim fileNumber As Integer = FreeFile()
        Dim currentRecord As String
        Dim invalidFileName As Boolean = True
        Dim customers(NumberOfCustomers(filePath) - 1, 8) As String 'array for customer data
        Dim currentCustomer As Integer = 0
        Try
            FileOpen(fileNumber, filePath, OpenMode.Input)

            Do Until EOF(fileNumber)
                Input(fileNumber, currentRecord)
                customers(0, 0) = currentRecord 'first name
                Input(fileNumber, currentRecord)
                customers(0, 1) = currentRecord 'last name
                Input(fileNumber, currentRecord)
                customers(0, 2) = currentRecord 'address
                Input(fileNumber, currentRecord)
                customers(0, 3) = currentRecord 'city
                Input(fileNumber, currentRecord)
                customers(0, 4) = currentRecord 'state
                Input(fileNumber, currentRecord)
                customers(0, 5) = currentRecord 'zip
                Input(fileNumber, currentRecord)
                customers(0, 6) = currentRecord 'phone number
                Input(fileNumber, currentRecord)
                customers(0, 7) = currentRecord 'email
                Input(fileNumber, currentRecord)
                customers(0, 8) = currentRecord 'customer ID

            Loop
            ' MsgBox($"There are {NumberOfCustomers(filePath)} customers")
            FileClose(fileNumber)

        Catch noFile As FileNotFoundException
            OpenCustomerFileDialog.FileName = ""
            OpenCustomerFileDialog.InitialDirectory = "C:\Users\noahh\Visual HW\SuperVideoStop"
            OpenCustomerFileDialog.Filter = "txt files (*txt)|*.txt|All files (*.*)|*.*"
            OpenCustomerFileDialog.ShowDialog()
            filePath = OpenCustomerFileDialog.FileName
            MsgBox($"The current file is {filePath}")

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Function NumberOfCustomers(fileName As String) As Integer
        Dim count As Integer = 0
        Dim fileNumber As Integer = FreeFile()

        Try
            FileOpen(fileNumber, fileName, OpenMode.Input)
            Do Until EOF(fileNumber)
                count += 1
            Loop
            LineInput(fileNumber)

            FileClose(fileNumber)
        Catch ex As Exception
            'pass
            'maybe set count to -1 to indicate error
        End Try

        Return count
    End Function



    '*************************************EVENT HANDLERS**********************************************
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        Me.Close()
    End Sub

    Private Sub UpdateButton_Click(sender As Object, e As EventArgs) Handles UpdateButton.Click
        ReadFromFile()
    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click

        DisplayListBox.Items.Clear()
    End Sub

    Private Sub OpenTopMenuItem_Click(sender As Object, e As EventArgs) Handles OpenTopMenuItem.Click
        LoadCustomerData()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim temp() As String

        temp = Split(ComboBox1.SelectedIndex.ToString, ",")

        temp(1) = temp(1).Trim() 'remove whitespace from both ends of string

        MsgBox($"the first name is: {temp(1)} and the last name is: {temp(0)}")
    End Sub
End Class
