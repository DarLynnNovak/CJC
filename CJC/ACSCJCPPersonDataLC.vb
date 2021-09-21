
Imports System.Drawing
Imports System.Windows.Forms
Imports Aptify.Framework.BusinessLogic.GenericEntity
Imports Aptify.Framework.DataServices
Imports Aptify.Framework.WindowsControls

Public Class ACSCJCPPersonDataLC
    Inherits FormTemplateLayout
    Dim recordId As Long
    Private m_oAppObj As New Aptify.Framework.Application.AptifyApplication
    Private m_oDA As New DataAction
    Private WithEvents lbPerson As AptifyLinkBox
    Private WithEvents tbFirstName As AptifyTextBox
    Private WithEvents tbMiddleName As AptifyTextBox
    Private WithEvents tbLastName As AptifyTextBox
    Private WithEvents tbDOB As AptifyTextBox
    Private WithEvents tbAge As AptifyTextBox
    Private WithEvents tbCity As AptifyTextBox
    Private WithEvents tbState As AptifyTextBox
    Private WithEvents dcbSpecialty As AptifyDataComboBox
    Dim Sql As String
    Dim PersonId As Long
    Dim dt As DataTable

    Protected Overrides Sub OnFormTemplateLoaded(ByVal e As FormTemplateLoadedEventArgs)
        Try
            Me.AutoScroll = True
            recordId = FormTemplateContext.GE.RecordID
            FindControls()
            If recordId > 0 Then
                PersonLookup()
            End If
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
        'MyBase.OnFormTemplateLoaded(e)
    End Sub

    Protected Overridable Sub FindControls()
        Try
            If lbPerson Is Nothing OrElse lbPerson.IsDisposed = True Then
                lbPerson = TryCast(GetFormComponent(Me, "ACS.ACSCJCMain.PersonID"), AptifyLinkBox)
            End If
            If tbFirstName Is Nothing OrElse tbFirstName.IsDisposed = True Then
                tbFirstName = TryCast(GetFormComponent(Me, "ACS.ACSCJCMain.FirstName"), AptifyTextBox)
            End If
            If tbMiddleName Is Nothing OrElse tbMiddleName.IsDisposed = True Then
                tbMiddleName = TryCast(GetFormComponent(Me, "ACS.ACSCJCMain.MiddleName"), AptifyTextBox)
            End If
            If tbLastName Is Nothing OrElse tbLastName.IsDisposed = True Then
                tbLastName = TryCast(GetFormComponent(Me, "ACS.ACSCJCMain.LastName"), AptifyTextBox)
            End If
            If tbDOB Is Nothing OrElse tbDOB.IsDisposed = True Then
                tbDOB = TryCast(GetFormComponent(Me, "ACS.ACSCJCMain.DOB"), AptifyTextBox)
            End If
            If tbAge Is Nothing OrElse tbAge.IsDisposed = True Then
                tbAge = TryCast(GetFormComponent(Me, "ACS.ACSCJCMain.Age"), AptifyTextBox)
            End If
            If tbCity Is Nothing OrElse tbCity.IsDisposed = True Then
                tbCity = TryCast(GetFormComponent(Me, "ACS.ACSCJCMain.City"), AptifyTextBox)
            End If
            If tbState Is Nothing OrElse tbState.IsDisposed = True Then
                tbState = TryCast(GetFormComponent(Me, "ACS.ACSCJCMain.State"), AptifyTextBox)
            End If
            If dcbSpecialty Is Nothing OrElse dcbSpecialty.IsDisposed = True Then
                dcbSpecialty = TryCast(GetFormComponent(Me, "ACS.ACSCJCMain.Specialty"), AptifyDataComboBox)
            End If
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub

    Private Sub PersonLookup()
        Try
            If lbPerson.Value > 0 Then
                Sql = "select ID,FirstName,LastName,MiddleName, format(Birthday,'MM/dd/yyyy')'DOB',(CASE WHEN DATEADD(yy, DATEDIFF(yy, Birthday, GETDATE()), Birthday) < GETDATE() THEN DATEDIFF(yy, Birthday, GETDATE()) ELSE DATEDIFF(yy, Birthday, GETDATE()) - 1 END) Age ,City,State,(select Name from vwACSSpecialty where id = ACSSpecID) Specialty from aptify..vwpersons (nolock) where id = " & PersonId
                dt = m_oDA.GetDataTable(Sql)
                If dt.Rows.Count > 0 Then
                    For Each dr As DataRow In dt.Rows
                        tbFirstName.Value = dr.Item("FirstName")
                        tbMiddleName.Value = dr.Item("MiddleName")
                        tbLastName.Value = dr.Item("LastName")
                        tbDOB.Value = dr.Item("DOB")
                        tbAge.Value = dr.Item("Age")
                        tbCity.Value = dr.Item("City")
                        tbState.Value = dr.Item("State")
                        dcbSpecialty.Value = dr.Item("Specialty")
                    Next
                End If
            End If
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try

    End Sub

    Private Sub lbPerson_ValueChanged(ByVal sender As Object, ByVal OldValue As Object, ByVal NewValue As Object) Handles lbPerson.ValueChanged
        Dim lTypeID As Long = -1

        PersonId = NewValue

        PersonLookup()
    End Sub
End Class
