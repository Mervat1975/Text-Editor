Option Strict On
Imports System.IO
Public Class FrmTextEditor
    Private Const APP_NAME As String = "Text Editor"
    Private dataDirty As Boolean
    Private fileName As String




    'Set the Dialog's initial directory.
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim init_dir As String = Application.StartupPath
        If init_dir.EndsWith("\bin") Then init_dir = init_dir.Substring(0, init_dir.Length - 4)
        dlgOpenFile.InitialDirectory = init_dir
        dlgSaveFile.InitialDirectory = init_dir
        Me.Text = APP_NAME + ": Select file to open"

    End Sub


    'Return true if it is safe to discard the current data.
    Private Function DataSafe() As Boolean
        If Not dataDirty Then Return True

        Select Case MessageBox.Show("The data has been modified. Do you want to save the changes?",
        "Save Changes?", MessageBoxButtons.YesNoCancel)
            Case Windows.Forms.DialogResult.Cancel
                'the user is canceling the operation.
                'dont' discard the changes.
                Return False
            Case Windows.Forms.DialogResult.No
                'the user wants to discard the changes
                Return True
            Case Windows.Forms.DialogResult.Yes
                'try to save the data
                SaveData(fileName)
                'see if the data was saved
                Return (Not dataDirty)
            Case Else
                Return False
        End Select
    End Function

    'load a data file
    Private Sub LoadData(ByVal file_name As String)
        Dim fileStream As IO.FileStream = Nothing
        Dim streamReader As IO.StreamReader = Nothing
        Try
            'load the file
            fileStream = New IO.FileStream(file_name, IO.FileMode.Open, IO.FileAccess.Read)
            streamReader = New IO.StreamReader(fileStream)
            Dim txt As String = streamReader.ReadToEnd
            rchFile.Text = txt

            'save the file name and title
            fileName = file_name
            Me.Text = APP_NAME & " [" & fileName & "]"
            dataDirty = False


            mnuFileSave.Enabled = False
            mnuFileSaveAs.Enabled = False
        Catch ex As Exception
            MessageBox.Show("Error loading file" & file_name &
            vbCrLf & ex.Message, "Load Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If Not (streamReader Is Nothing) Then streamReader.Close()
        End Try
    End Sub

    'save the file.
    Private Sub SaveData(ByVal file_name As String)
        Dim filestream As IO.FileStream = Nothing
        Dim streamWriter As IO.StreamWriter = Nothing
        Try
            'save the file
            filestream = New IO.FileStream(file_name, IO.FileMode.Create, IO.FileAccess.Write)
            streamWriter = New IO.StreamWriter(filestream)
            streamWriter.Write(Me.rchFile.Text)
            streamWriter.Close()
            'save the file name and title.
            fileName = file_name
            Me.Text = APP_NAME & " [" & file_name & "]"
            dataDirty = False

            Me.mnuFileSave.Enabled = False
            Me.mnuFileSaveAs.Enabled = False
        Catch ex As Exception
            MessageBox.Show("Error saving file " & file_name & vbCrLf & ex.Message,
            "Save Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'mark the data as modified.
    Private Sub rchFile_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rchFile.TextChanged
        If Not dataDirty Then
            Me.Text = APP_NAME & "*[" & fileName & "]"
            dataDirty = True
            mnuFileSave.Enabled = True
            mnuFileSaveAs.Enabled = True
        End If
    End Sub

    ' save the file.
    Private Sub mnuFileSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileSave.Click
        If fileName Is Nothing Then
            mnuFileSaveAs_Click(sender, e)
        Else
            SaveData(fileName)
        End If
    End Sub
    'save the file with a new name.
    Private Sub mnuFileSaveAs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileSaveAs.Click
        dlgSaveFile.Filter = "Word (*.doc) |*.doc;*.rtf|(*.txt) |*.txt|(*.*) |*.*"
        If dlgSaveFile.ShowDialog = Windows.Forms.DialogResult.OK Then
            SaveData(dlgSaveFile.FileName)
        End If
    End Sub
    'close the application
    Private Sub mnuFileExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileExit.Click
        Me.Close()
    End Sub
    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        e.Cancel = Not DataSafe()
    End Sub

    'Open a file
    Private Sub mnuFileOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileOpen.Click
        'make sure the current data is safe
        If Not DataSafe() Then Exit Sub
        dlgOpenFile.Filter = "Word Documents|*.doc|Excel Worksheets|*.xls|PowerPoint Presentations|*.ppt|Office Files|*.doc;*.xls;*.ppt|All Files|*.*"
        If dlgOpenFile.ShowDialog = Windows.Forms.DialogResult.OK Then
            LoadData(dlgOpenFile.FileName)
        End If
    End Sub
    'start a new document
    Private Sub mnuFileNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileNew.Click
        rchFile.Enabled = True

        'Make sure the current data is safe.
        If Not DataSafe() Then Exit Sub

        Me.rchFile.Text = ""
        fileName = Nothing
        Me.Text = APP_NAME & " []"
        dataDirty = False

        'no point in saving a blank file
        Me.mnuFileSave.Enabled = False
        Me.mnuFileSaveAs.Enabled = False
    End Sub

    Private Sub mnuHelpAbout_Click(sender As Object, e As EventArgs) Handles mnuHelpAbout.Click
        MessageBox.Show("NETD 2202" & vbNewLine & "Lab#5" & vbNewLine & "MervatMustafa", "About", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub mnuEditCut_Click(sender As Object, e As EventArgs) Handles mnuEditCut.Click
        'Checks to see if the user selected anything
        If rchFile.SelectedText <> "" Then
            'Good, the user selected something
            'Copy the information to the clipbaord
            My.Computer.Clipboard.SetText(rchFile.SelectedText)
            'Since this is a cut command, we want to clear whatever 
            'text they had selected when they clicked cut
            rchFile.SelectedText = ""
        Else
            'If there was no text selected, print out an error message box
            MessageBox.Show("No text is selected to cut", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If
    End Sub

    Private Sub mnuEditCopy_Click(sender As Object, e As EventArgs) Handles mnuEditCopy.Click
        'Checks to see if the user selected anything
        If rchFile.SelectedText <> "" Then
            'Copy the information to the clipboard
            My.Computer.Clipboard.SetText(rchFile.SelectedText)
        Else
            'If no text was selected, print out an error message box
            MessageBox.Show("No text is selected to copy", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If
    End Sub

    Private Sub mnuEditPaste_Click(sender As Object, e As EventArgs) Handles mnuEditPaste.Click
        'Get the data stored in the clipboard
        Dim iData As IDataObject = My.Computer.Clipboard.GetDataObject()
        'Check to see if the data is in a text format
        If iData.GetDataPresent(DataFormats.Text) Then
            'If it's text, then paste it into the textbox
            rchFile.SelectedText = CType(iData.GetData(DataFormats.Text), String)
        Else
            'If it's not text, print a warning message
            MessageBox.Show("Data in the clipboard is not availble for entry into a textbox", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If
    End Sub


End Class
