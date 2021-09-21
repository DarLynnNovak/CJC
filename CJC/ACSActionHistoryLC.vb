
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Configuration
Imports Aptify.Framework.BusinessLogic.GenericEntity
Imports Aptify.Framework.DataServices
Imports Aptify.Framework.WindowsControls
Imports System.Data
Imports System
Imports System.IO
Imports IronPdf


Public Class ACSActionHistoryLC
    Inherits FormTemplateLayout
    Private m_oAppObj As New Aptify.Framework.Application.AptifyApplication
    Private m_oDA As New DataAction
    Private WithEvents tbAHName As AptifyTextBox
    Private WithEvents tbAHArticle As AptifyComboBox
    Private WithEvents tbAHSection As AptifyComboBox
    Private WithEvents tbAHPart As AptifyComboBox
    Private WithEvents tbAHPart2 As AptifyComboBox
    Private WithEvents tbAHPart3 As AptifyComboBox
    Private WithEvents tbAHPart4 As AptifyComboBox
    Dim Article As String
    Dim Section As String
    Dim Part As String
    Dim Part2 As String
    Dim Part3 As String
    Dim Part4 As String
    Dim ActionHistoryName As String


    'Dim is_licensed As Boolean = IronPdf.License.IsLicensed
    Protected Overrides Sub OnFormTemplateLoaded(ByVal e As FormTemplateLoadedEventArgs)
        Try

            Me.AutoScroll = True

            FindControls()


            'MsgBox(is_licensed)
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
        'MyBase.OnFormTemplateLoaded(e)
    End Sub

    Protected Overridable Sub FindControls()
        Try
            If tbAHName Is Nothing OrElse tbAHName.IsDisposed = True Then
                tbAHName = TryCast(GetFormComponent(Me, "ACSCJCActionHistoryCharge.Name"), AptifyTextBox)
            End If
            If tbAHArticle Is Nothing OrElse tbAHArticle.IsDisposed = True Then
                tbAHArticle = TryCast(GetFormComponent(Me, "ACSCJCActionHistoryCharge.Article"), AptifyComboBox)
            End If
            If tbAHSection Is Nothing OrElse tbAHSection.IsDisposed = True Then
                tbAHSection = TryCast(GetFormComponent(Me, "ACSCJCActionHistoryCharge.Section"), AptifyComboBox)
            End If
            If tbAHPart Is Nothing OrElse tbAHPart.IsDisposed = True Then
                tbAHPart = TryCast(GetFormComponent(Me, "ACSCJCActionHistoryCharge.Part"), AptifyComboBox)
            End If
            If tbAHPart2 Is Nothing OrElse tbAHPart2.IsDisposed = True Then
                tbAHPart2 = TryCast(GetFormComponent(Me, "ACSCJCActionHistoryCharge.Part2"), AptifyComboBox)
            End If
            If tbAHPart3 Is Nothing OrElse tbAHPart3.IsDisposed = True Then
                tbAHPart3 = TryCast(GetFormComponent(Me, "ACSCJCActionHistoryCharge.Part3"), AptifyComboBox)
            End If
            If tbAHPart4 Is Nothing OrElse tbAHPart4.IsDisposed = True Then
                tbAHPart4 = TryCast(GetFormComponent(Me, "ACSCJCActionHistoryCharge.Part4"), AptifyComboBox)
            End If
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub


    Private Sub tbAHArticle_ValueChanged(ByVal sender As Object, ByVal OldValue As Object, ByVal NewValue As Object) Handles tbAHArticle.ValueChanged
        If Not NewValue Is Nothing Then
            Article = "Article " & NewValue
            UpdateName()
        End If

    End Sub

    Private Sub tbAHSection_ValueChanged(ByVal sender As Object, ByVal OldValue As Object, ByVal NewValue As Object) Handles tbAHSection.ValueChanged
        If Not NewValue Is Nothing Then
            Section = "Section " & NewValue
            UpdateName()

        End If

    End Sub
    Private Sub tbAHPart_ValueChanged(ByVal sender As Object, ByVal OldValue As Object, ByVal NewValue As Object) Handles tbAHPart.ValueChanged
        If Not NewValue Is Nothing Then
            Part = "(" & NewValue & ")"
            UpdateName()

        End If

    End Sub
    Private Sub tbAHPart2_ValueChanged(ByVal sender As Object, ByVal OldValue As Object, ByVal NewValue As Object) Handles tbAHPart2.ValueChanged
        If Not NewValue Is Nothing Then
            Part2 = ", (" & NewValue & ")"
            UpdateName()

        End If

    End Sub
    Private Sub tbAHPart3_ValueChanged(ByVal sender As Object, ByVal OldValue As Object, ByVal NewValue As Object) Handles tbAHPart3.ValueChanged
        If Not NewValue Is Nothing Then
            Part3 = ", (" & NewValue & ")"
            UpdateName()

        End If

    End Sub
    Private Sub tbAHPart4_ValueChanged(ByVal sender As Object, ByVal OldValue As Object, ByVal NewValue As Object) Handles tbAHPart4.ValueChanged
        If Not NewValue Is Nothing Then
            Part4 = ", (" & NewValue & ")"
            UpdateName()

        End If

    End Sub
    Private Sub UpdateName()
        'tbAHName.Value = ""
        ActionHistoryName = Article & ", " & Section & " " & Part & Part2 & Part3 & Part4
        tbAHName.Value = ActionHistoryName

    End Sub
End Class
