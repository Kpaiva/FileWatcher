'   Copyright 2013-2016 Kevin Paiva 

'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.

'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.

'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.

Imports System
Imports System.Security
Imports System.Security.Cryptography
Imports System.Text
Imports System.IO
Imports System.Timers
Imports outlook = Microsoft.Office.Interop.Outlook

Public Class frmMain

    Private Running As Boolean
    Private fileToCheck As String = ""
    Private TimeToCompare As Boolean
    Private hashBefore As Byte()
    Private hashAfter As Byte()
    Private emails As List(Of String) = New List(Of String)

    'Modify the following 2 lines to customize the Subject and Body
    Private subjLine As String = "Enter Subject Line Here"
    Private bodyMsg As String = "Enter Body Msg here"


    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Add emails like this
        emails.Add("email@example.com")

        Running = False
        lblRunning.Text = "Stopped"

        hashBefore = ComputeFileHash(fileToCheck)

        'Change these intervals for the timers if you want to change wait times.
        'Note: probably a good idea to have Timer1 use a bit more time than Timer2

        Timer1.Interval() = 310000
        AddHandler Timer1.Tick, AddressOf OnTimerEvent
        Timer1.Enabled = True
        Timer1.Stop()

        Timer2.Interval() = 300000
        AddHandler Timer2.Tick, AddressOf OnTimerEvent2
        Timer2.Enabled = True
        Timer2.Stop()

    End Sub

    Private Sub btnStartStop_Click(sender As Object, e As EventArgs) Handles btnStartStop.Click

        'Just a basic running/not running toggle.

        If Running = False Then
            fileToCheck = TextBox1.Text
            If fileToCheck <> "" Then
                Running = True
                lblRunning.Text = "Running"
                TextBox1.Enabled = False
                Timer2.Start()
                Timer1.Start()
            Else
                MessageBox.Show("Please Enter a Full File Path.")
            End If
        Else
            Running = False
            lblRunning.Text = "Stopped"
            TextBox1.Enabled = True
            Timer2.Stop()
            Timer1.Stop()
        End If

    End Sub

    Private Sub CheckFile(ByVal fileName As String)

        If (Running = True) Then

            hashAfter = ComputeFileHash(fileToCheck)

            If CompareByteHashes(hashAfter, hashBefore) Then
                SendEmailNotification()
            End If

        End If

    End Sub

    ' Calculates a file's hash value and returns it as a byte array.
    Private Function ComputeFileHash(ByVal fileName As String) As Byte()
        Dim ourHash(0) As Byte

        ' If file exists, create a HashAlgorithm instance based off of MD5 encryption
        ' You could use a variant of SHA or RIPEMD160 if you like with larger hash bit sizes.
        If File.Exists(fileName) Then
            'Try
            Dim ourHashAlg As HashAlgorithm = HashAlgorithm.Create("MD5")
            Dim fileToHash As FileStream = New FileStream(fileName, FileMode.Open, FileAccess.Read)

            'Compute the hash to return using the Stream we created.
            ourHash = ourHashAlg.ComputeHash(fileToHash)

            fileToHash.Close()
            'Catch ex As IOException
            '    MessageBox.Show("There was an error opening the file: " & ex.Message)
            'End Try
        End If

        Return ourHash
    End Function


    ' Return true/false if the two hashes are the same.
    Private Function CompareByteHashes(ByVal newHash As Byte(), ByVal oldHash As Byte()) As Boolean

        ' If any of these conditions are true, the hashes are definitely not the same.
        If newHash Is Nothing Or oldHash Is Nothing Or newHash.Length <> oldHash.Length Then
            Return False
        End If


        ' Compare each byte of the two hashes. Any time they are not the same, we know there was a change.
        For i As Integer = 0 To newHash.Length - 1
            If newHash(i) <> oldHash(i) Then
                Return False
            End If
        Next i

        Return True
    End Function

    Private Sub OnTimerEvent(ByVal [source] As Object, ByVal e As EventArgs)
        If fileToCheck <> "" Then
            CheckFile(fileToCheck)
        End If
    End Sub 'OnTimerEvent

    Private Sub OnTimerEvent2(ByVal [source] As Object, ByVal e As EventArgs)
        If fileToCheck <> "" Then
            hashBefore = ComputeFileHash(fileToCheck)
        End If
    End Sub 'OnTimerEvent

    Private Sub SendEmailNotification()

        'This function can be used however you like. Currently it uses outlook to send an email with the provided
        'body and subjects as well as the provided email recipients

        Dim OutlookMessage As outlook.MailItem
        Dim AppOutlook As New outlook.Application
        Try
            OutlookMessage = AppOutlook.CreateItem(outlook.OlItemType.olMailItem)
            Dim Recipents As outlook.Recipients = OutlookMessage.Recipients

            For Each t In emails
                Recipents.Add(t)
            Next
            OutlookMessage.Subject = subjLine
            OutlookMessage.Body = bodyMsg
            OutlookMessage.BodyFormat = outlook.OlBodyFormat.olFormatHTML
            OutlookMessage.Send()
        Catch ex As Exception
            MessageBox.Show("Mail could not be sent") 'if you dont want this message, simply delete this line 
        Finally
            OutlookMessage = Nothing
            AppOutlook = Nothing
        End Try

    End Sub

End Class
