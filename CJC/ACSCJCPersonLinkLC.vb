'Option Explicit On
'Option Strict On

Imports System.Drawing
Imports System.Windows.Forms
Imports Aptify.Framework.BusinessLogic.GenericEntity
Imports Aptify.Framework.DataServices
Imports Aptify.Framework.WindowsControls

Public Class ACSCJCPersonLinkLC
    Inherits FormTemplateLayout
    Private m_oAppObj As New Aptify.Framework.Application.AptifyApplication
    Private m_oDA As New DataAction
    Private bAdded As Boolean = False
    Private lGridID As Long = -1

    Private WithEvents grdPeopleSearch As DataGridView
    Protected WithEvents lblSearch As CultureLabel
    Public WithEvents btnCreate As AptifyActiveButton
    Public WithEvents btnUpdate As AptifyActiveButton

    Private WithEvents lbPerson As AptifyLinkBox
    Private WithEvents txtFirstName As AptifyTextBox
    Private WithEvents txtLastName As AptifyTextBox
    Private WithEvents txtEmail As AptifyTextBox

    Private _sCheckPersonName As String = Nothing
    Private _sCheckPersonType As String = Nothing

    Private _sCheckCountryCode As String = Nothing
    Private _sCheckCity As String = Nothing
    Private _sCheckState As String = Nothing

    Dim dtSearch As DataTable
    Dim Sql As String
    Dim FirstName As String
    Dim LastName As String
    Dim currentDate As DateTime = Now
    Dim acsCJCMain As AptifyGenericEntity

    Protected Overrides Sub OnFormTemplateLoaded(ByVal e As FormTemplateLoadedEventArgs)
        Try
            Me.AutoScroll = True

            Dim acsCJCMainId As Long = FormTemplateContext.GE.RecordID
            acsCJCMain = CType(m_oAppObj.GetEntityObject("acsCJCMain", acsCJCMainId), AptifyGenericEntity)
            FindControls()
            If grdPeopleSearch Is Nothing Then
                grdPeopleSearch = CreateGrid()
            End If
            'AssignBaseFieldsForChecking() 
            'lblSearch.Text = ""

            Dim lTypeID As Long = -1



            FirstName = Me.FormTemplateContext.GE.GetValue("FirstName")
            LastName = Me.FormTemplateContext.GE.GetValue("LastName")

            PersonLookup(FirstName, LastName, "")
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
        'MyBase.OnFormTemplateLoaded(e)
    End Sub
    Protected Overridable Sub FindControls()
        Try
            If lblSearch Is Nothing OrElse lblSearch.IsDisposed = True Then
                lblSearch = TryCast(GetFormComponent(Me, "ACS.ACSCJCMain.Tabs.General.Culture Label.1"), CultureLabel)
            End If
            If btnUpdate Is Nothing OrElse btnUpdate.IsDisposed = True Then
                btnUpdate = TryCast(GetFormComponent(Me, "ACS.ACSCJCMain.Tabs.General.Active Button.2"), AptifyActiveButton)
            End If

            If btnCreate Is Nothing OrElse btnCreate.IsDisposed = True Then
                btnCreate = TryCast(GetFormComponent(Me, "ACS.ACSCJCMain.Tabs.General.Active Button.1"), AptifyActiveButton)
            End If
            If lbPerson Is Nothing OrElse lbPerson.IsDisposed = True Then
                lbPerson = TryCast(GetFormComponent(Me, "ACS.ACSCJCMain.PersonId"), AptifyLinkBox)
            End If
            If txtFirstName Is Nothing OrElse txtFirstName.IsDisposed = True Then
                txtFirstName = TryCast(GetFormComponent(Me, "ACS.ACSCJCMain.FirstName"), AptifyTextBox)
            End If
            If txtLastName Is Nothing OrElse txtLastName.IsDisposed = True Then
                txtLastName = TryCast(GetFormComponent(Me, "ACS.ACSCJCMain.LastName"), AptifyTextBox)
            End If


            If txtEmail Is Nothing OrElse txtEmail.IsDisposed = True Then
                txtEmail = TryCast(GetFormComponent(Me, "ACS.ACSCJCMain.Email"), AptifyTextBox)
            End If


            CheckComLB()
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub
    Public Sub CheckComLB()
        If Not lbPerson Is Nothing AndAlso CInt(lbPerson.Value) > 0 Then
            grdPeopleSearch.Hide()
            lblSearch.Hide()
            btnUpdate.Visible = False
            btnCreate.Visible = False
        ElseIf dtSearch.Rows.Count > 0 Then
            grdPeopleSearch.Show()
            btnUpdate.Visible = True
            btnCreate.Visible = True
        Else
            grdPeopleSearch.Show()
            btnUpdate.Visible = False
            btnCreate.Visible = True
        End If
    End Sub

    Private Function CreateGrid() As DataGridView
        Try
            Dim grdReturn As DataGridView
            'Dim gridtop = lCompanyLinkbox.Top + lCompanyLinkbox.Height + 10 
            grdReturn = New DataGridView
            grdReturn.Name = "grdPeopleSearch"
            grdReturn.Size = New Drawing.Size(500, 150)
            grdReturn.Location = New Drawing.Point(50, 55)
            Controls.Add(grdReturn)
            grdReturn.Visible = True
            Return grdReturn
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
            Return Nothing
        End Try
    End Function

    Private Sub txtFirstName_ValueChanged(ByVal sender As Object, ByVal OldValue As Object, ByVal NewValue As Object) Handles txtFirstName.ValueChanged
        Dim lTypeID As Long = -1

        FirstName = NewValue.ToString()
        LastName = txtLastName.Value
        PersonLookup(FirstName, LastName, "")
    End Sub


    Private Sub txtLastName_ValueChanged(ByVal sender As Object, ByVal OldValue As Object, ByVal NewValue As Object) Handles txtLastName.ValueChanged
        Dim lTypeID As Long = -1

        LastName = NewValue.ToString()
        FirstName = txtFirstName.Value
        'FirstName = Me.FormTemplateContext.GE.GetValue("FirstName")
        PersonLookup(FirstName, LastName, "")
    End Sub

    Private Sub PersonLookup(ByVal FirstName As String, ByVal LastName As String, ByVal Message As String)
        Try
            'Dim dtSearch As DataTable
            'lblSearch.Text = ""

            If Not FirstName Is Nothing AndAlso Not FirstName.ToString = "" Or Not LastName Is Nothing AndAlso Not LastName.ToString = "" Then
                Sql = "select ID, FirstName'First',LastName'Last', format(Birthday,'MM/dd/yyyy')'DOB',email1'Email',ACSMemberClassID_Name'Member Class' from aptify..vwpersons where Firstname like " & "'%" & FirstName.ToString & "%'" & "and Lastname like " & "'%" & LastName.ToString & "%' "
                dtSearch = m_oDA.GetDataTable(Sql)

                If dtSearch.Rows.Count > 0 Then
                    'btnUpdate.Visible = True
                    'lblSearch.Text = Message
                    grdPeopleSearch.DataSource = dtSearch
                    grdPeopleSearch.Columns(0).ReadOnly = False
                    'Check to see if the CheckBox Column Exists
                    If grdPeopleSearch.Columns(0).CellType.ToString = "System.Windows.Forms.DataGridViewCheckBoxCell" Then
                        bAdded = True
                    Else
                        bAdded = False
                    End If
                    If bAdded = False Then
                        Dim AddColumn As New DataGridViewCheckBoxColumn

                        With AddColumn
                            .HeaderText = ""
                            .Name = "grdChecked"
                            .Width = 21
                        End With

                        grdPeopleSearch.Columns.Insert(0, AddColumn)

                        For i As Integer = 1 To grdPeopleSearch.ColumnCount - 1
                            grdPeopleSearch.Columns(i).ReadOnly = True
                        Next

                        bAdded = True
                    End If
                    lblSearch.Text = "Please choose one of the following records and click update person."
                Else
                    If LTrim(RTrim(FirstName)) = "" OrElse LTrim(RTrim(FirstName.ToString())) = "" Then

                        lblSearch.Text = "No match found as Person Name is blank."


                    Else
                        lblSearch.Text = "No match found. Please try a different Person Name or click Create New Person."



                    End If

                    lblSearch.Refresh()
                    grdPeopleSearch.DataSource = dtSearch
                    grdPeopleSearch.Columns(0).ReadOnly = True
                End If
            Else

                If LTrim(RTrim(FirstName)) = "" OrElse LTrim(RTrim(FirstName.ToString())) = "" Then

                    lblSearch.Text = "No match found as Person Name is blank."


                Else
                    lblSearch.Text = "No match found. Please try a different Person Name or click Create New Person."


                End If

                lblSearch.Refresh()
                dtSearch = m_oDA.GetDataTable("select ID = '', FirstName = '', LastName = '', Birthday = '',Email1 = '', ACSMemberClassID_Name", IAptifyDataAction.DSLCacheSetting.BypassCache)
                dtSearch.Clear()
                grdPeopleSearch.DataSource = dtSearch
                grdPeopleSearch.Columns(0).ReadOnly = True

            End If
            CheckComLB()
            grdPeopleSearch.AllowUserToAddRows = False
            grdPeopleSearch.Refresh()
            lGridID = -1
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub

    Private Sub grdPeopleSearch_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grdPeopleSearch.CellValueChanged
        Try
            Dim bFound As Boolean = False
            If e.ColumnIndex = 0 Then
                For i As Integer = 0 To grdPeopleSearch.RowCount - 1
                    If Not grdPeopleSearch.Item(0, i).Value Is Nothing _
                           AndAlso grdPeopleSearch.Item(0, i).Value.ToString.ToUpper = "TRUE" Then
                        lGridID = CLng(grdPeopleSearch.Item("ID", i).Value)
                        bFound = True
                    End If
                Next
            End If
            If bFound = False Then
                lGridID = -1
            End If
            grdPeopleSearch.Invalidate()

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub

    Private Sub grdPersonSearch_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grdPeopleSearch.CellClick
        Try
            If e.ColumnIndex = 0 Then
                For i As Integer = 0 To grdPeopleSearch.RowCount - 1
                    If Not grdPeopleSearch.Item(0, i).Value Is Nothing _
                        AndAlso Not grdPeopleSearch.Item(0, i).Value.ToString.ToUpper = "FALSE" Then
                        grdPeopleSearch.Item(0, i).Value = False

                    End If
                Next
            End If
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub

    Private Sub grdPeopleSearch_CurrentCellDirtyStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdPeopleSearch.CurrentCellDirtyStateChanged
        If grdPeopleSearch.IsCurrentCellDirty Then
            grdPeopleSearch.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub
    Private Sub btnUpdate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpdate.Click

        Try

            Dim lID As Long = -1
            If lGridID > 0 Then
                lID = lGridID
                Dim oPerson As AptifyGenericEntity
                oPerson = CType(m_oAppObj.GetEntityObject("Persons", lID), AptifyGenericEntity)

                Dim sPersonFirstName As String = ""
                Dim sPersonLastName As String = ""
                Dim sPersonEmail As String = ""
                sPersonFirstName = oPerson.GetValue("FirstName").ToString
                sPersonLastName = oPerson.GetValue("LastName").ToString
                'sPersonEmail = FormTemplateContext.GE.GetValue("Email")
                sPersonEmail = oPerson.GetValue("LastName").ToString


                Select Case MsgBox("Are you sure you want to update " & oPerson.GetValue("FirstName").ToString & " " & oPerson.GetValue("LastName").ToString & " ?", MsgBoxStyle.YesNo, "Person Update")
                    Case MsgBoxResult.Yes

                        With oPerson
                            '    'Need to make sure that the values on the Form are updating the GE
                            '    'UpdateGE()

                            .SetValue("FirstName", sPersonFirstName)
                            .SetValue("LastName", sPersonLastName)
                            .SetValue("CountryCodeId", 222)
                            .SetValue("City", FormTemplateContext.GE.GetValue("City"))
                            .SetValue("State", FormTemplateContext.GE.GetValue("State"))
                            If .GetValue("Email1") Is "" Then
                                .SetValue("Email1", sPersonEmail)
                            End If
                        End With
                        Dim sErr As String = ""
                        oPerson.Save(sErr)

                        If sErr.Length > 0 Then
                            ShowMessage("There was a problem updating " & sPersonFirstName & " " & sPersonLastName & ". Please refer to the Aptify exception log.", True)
                        Else
                            lblSearch.Text = "Person " & sPersonFirstName & " " & sPersonLastName & " updated successfully!"
                        End If
                        ''''''''''''''''''''''

                        Dim lTypeID As Long = -1

                        If oPerson.RecordID > 0 Then
                            With acsCJCMain
                                lbPerson.Value = oPerson.RecordID

                            End With
                        End If

                        PersonLookup(txtFirstName.Value.ToString(), LastName, "")

                End Select
            Else
                MsgBox("Please select one matching People from the list above.")
            End If
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try

    End Sub
    Private Sub btnCreate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCreate.Click
        Try

            Dim sPersonFirstName As String = ""
            Dim sPersonLastName As String = ""
            sPersonFirstName = txtFirstName.Value.ToString
            sPersonLastName = txtLastName.Value.ToString

            If Not lbPerson Is Nothing Then
                If IsNumeric(lbPerson.RecordID) AndAlso lbPerson.RecordID > 0 Then
                    If lbPerson.RecordName = lbPerson.Text Then
                        MsgBox("You cannot create a duplicate which currently is linked to " & lbPerson.RecordName & ".")
                        Exit Sub
                    End If
                End If
            End If

            Select Case MsgBox("Are you sure you want to create " & sPersonFirstName & " " & sPersonLastName & " ?", MsgBoxStyle.YesNo, "Create Person")
                Case MsgBoxResult.Yes

                    Dim oPerson As AptifyGenericEntity
                    oPerson = CType(m_oAppObj.GetEntityObject("Persons", -1), AptifyGenericEntity)
                    With oPerson
                        .SetValue("FirstName", FormTemplateContext.GE.GetValue("FirstName"))
                        .SetValue("LastName", FormTemplateContext.GE.GetValue("LastName"))
                        .SetValue("CountryCodeId", 222)
                        .SetValue("Phone", FormTemplateContext.GE.GetValue("Phone"))
                        .SetValue("Email1", FormTemplateContext.GE.GetValue("Email"))

                    End With

                    Dim sErr As String = ""
                    oPerson.Save(sErr)

                    If sErr.Length > 0 Then
                        ShowMessage("There was a problem updating " & sPersonFirstName & " " & sPersonLastName & ". Please refer to the Aptify exception log.", True)
                    Else
                        lblSearch.Text = "Person " & sPersonFirstName & " " & sPersonLastName & " updated successfully!"
                    End If

                    If oPerson.RecordID > 0 Then
                        lbPerson.Value = oPerson.RecordID

                    End If

                    lbPerson.Value = oPerson.RecordID
                    oPerson.Save()

                    Dim lTypeID As Long = -1

                    PersonLookup(txtFirstName.Value.ToString, txtLastName.Value.ToString(), "")

            End Select

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try

    End Sub

    Private Sub InitializeComponent()
        SuspendLayout()
        '
        'ACSMBSPreAppCompanyLookupLC
        '
        Name = "ACSCJCPersonLinkLC"
        Size = New Drawing.Size(1237, 707)
        ResumeLayout(False)

    End Sub

    Private Sub ShowMessage(ByVal Message As String, Optional ByVal IsError As Boolean = False)
        lblSearch.Text = Message
        lblSearch.ForeColor = Color.Black
        If IsError Then
            lblSearch.ForeColor = Color.Red
        End If
    End Sub

End Class