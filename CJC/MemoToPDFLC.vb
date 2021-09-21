Imports System.Configuration
Imports System.IO
Imports Aptify.Framework.BusinessLogic.GenericEntity
Imports Aptify.Framework.DataServices
Imports Aptify.Framework.WindowsControls
Imports IronPdf




Public Class MemoToPDFLC
    Inherits FormTemplateLayout

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
    Private WithEvents btnCreateEWNCM As AptifyActiveButton
    Private WithEvents btnCreateEWDM As AptifyActiveButton
    Private WithEvents btnCreateEWAM As AptifyActiveButton
    Private WithEvents btnCreateEWCS As AptifyActiveButton
    Private WithEvents btnCreateCJCCS As AptifyActiveButton
    Private WithEvents btnCreateCJCDM As AptifyActiveButton
    Private WithEvents dcbSpecialty As AptifyDataComboBox
    Dim AttachmentsGE As AptifyGenericEntity
    Dim CJCGE As AptifyGenericEntityBase
    Dim entityId As Long
    Dim entityIdSql As String
    Dim attachmentCatId As Long
    Dim attachmentCatIdSql As String
    Dim actionHistoryIdSql As String
    Dim checkAttachmentIdSql As String
    Dim checkAttachmentId As DataTable
    Dim ahdt As DataTable
    Dim attachId As Long
    Dim data As Byte()
    Dim filename As String
    Dim result As String = "Failed"
    Dim recordId As Long
    Dim Sql As String
    Dim dt As DataTable
    Dim saveLocalPrefix As String = "C:\Users\Public\Documents\"
    Dim saveLocation As String
    Dim FirstName As String
    Dim MiddleName As String
    Dim LastName As String
    Dim City As String
    Dim State As String
    Dim DOB As String
    Dim Age As Long
    Dim FACSStatus As String
    Dim Specialty As String
    Dim StateActionDesc As String
    Dim StateActionDate As String
    Dim CurrentStateLicensure As String
    Dim Charge As String
    Dim Comments As String
    Dim CaseName As String
    Dim AHBORFinalActionDate As String
    Dim ACSCJCBORFinalID_Name As String
    Dim PDFHeader As String
    Dim PDFFooter As String
    Dim PDFText As String
    Dim PDFText2 As String
    Dim theDate As Date = Now()
    Dim outPdfBuffer As Byte()
    Dim Licresult As Boolean = License.IsValidLicense("IRONPDF-49375960F4-720786-6BFD0A-2F308E1D40-7ECEEDDF-UEx60905F54E21F8D8-AMERICANCOLLEGEOFSURGEONS.IRO200326.7419.20147.ORG.5DEV.1YR.SUPPORTED.UNTIL.27.MAR.2021")
    Dim is_licensed As Boolean = License.IsLicensed
    Dim AHRow As String
    Dim AHText As String
    Dim AHCharge As String
    Dim SRRow As String
    Dim SRText As String
    Dim staffRecSql As String
    Dim srdt As DataTable
    Dim SRName As String



    'Dim is_licensed As Boolean = IronPdf.License.IsLicensed
    Protected Overrides Sub OnFormTemplateLoaded(ByVal e As FormTemplateLoadedEventArgs)
        Try

            Me.AutoScroll = True
            recordId = FormTemplateContext.GE.RecordID
            FindControls()

            CheckAttachment()
            'MsgBox(is_licensed)
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
            If btnCreateEWNCM Is Nothing OrElse btnCreateEWNCM.IsDisposed = True Then
                btnCreateEWNCM = TryCast(GetFormComponent(Me, "ACSCJCMain Memos.Active Button.1"), AptifyActiveButton)
            End If
            If btnCreateEWDM Is Nothing OrElse btnCreateEWDM.IsDisposed = True Then
                btnCreateEWDM = TryCast(GetFormComponent(Me, "ACSCJCMain Memos.Active Button.2"), AptifyActiveButton)
            End If
            If btnCreateEWAM Is Nothing OrElse btnCreateEWAM.IsDisposed = True Then
                btnCreateEWAM = TryCast(GetFormComponent(Me, "ACSCJCMain Memos.Active Button.3"), AptifyActiveButton)
            End If
            If btnCreateEWCS Is Nothing OrElse btnCreateEWCS.IsDisposed = True Then
                btnCreateEWCS = TryCast(GetFormComponent(Me, "ACSCJCMain Memos.Active Button.4"), AptifyActiveButton)
            End If
            If btnCreateCJCCS Is Nothing OrElse btnCreateCJCCS.IsDisposed = True Then
                btnCreateCJCCS = TryCast(GetFormComponent(Me, "ACSCJCMain Memos.Active Button.5"), AptifyActiveButton)
            End If
            If btnCreateCJCDM Is Nothing OrElse btnCreateCJCDM.IsDisposed = True Then
                btnCreateCJCDM = TryCast(GetFormComponent(Me, "ACSCJCMain Memos.Active Button.6"), AptifyActiveButton)
            End If
            CheckFormFields()
            CheckKey()
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub
    Private Sub CheckFormFields()
        If recordId > 0 Then
            btnCreateEWNCM.Visible = True
            btnCreateEWDM.Visible = True
            btnCreateEWAM.Visible = True
            btnCreateEWCS.Visible = True
            btnCreateCJCCS.Visible = True
            btnCreateCJCDM.Visible = True
        Else
            btnCreateEWNCM.Visible = False
            btnCreateEWDM.Visible = False
            btnCreateEWAM.Visible = False
            btnCreateEWCS.Visible = False
            btnCreateCJCCS.Visible = False
            btnCreateCJCDM.Visible = False
        End If

    End Sub
    Private Sub CheckKey()

        Try
            ' Dim appSettingsReader = New AppSettingsReader()
            'Dim key As String = appSettingsReader.GetValue("IronPdfKey", GetType(String))
            Dim key As String = "IRONPDF-49375960F4-720786-6BFD0A-2F308E1D40-7ECEEDDF-UEx60905F54E21F8D8-AMERICANCOLLEGEOFSURGEONS.IRO200326.7419.20147.ORG.5DEV.1YR.SUPPORTED.UNTIL.27.MAR.2021"
            ' Example reading key from appsettings
            License.LicenseKey = key
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try


    End Sub

    Private Sub btnCreateEWNCM_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCreateEWNCM.Click
        Try
            If Me.FormTemplateContext.GE.RecordID > 0 Then
                CreateEWNewCaseMemo()
                CreatePDF()
            Else
                MsgBox("This record has not been created yet.  Please save the form to create the record")
            End If
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub
    Private Sub btnCreateEWDM_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCreateEWDM.Click
        Try
            If Me.FormTemplateContext.GE.RecordID > 0 Then
                CreateEWDirectorsMemo()
                CreatePDF()
            Else
                MsgBox("This record has not been created yet.  Please save the form to create the record")
            End If
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub
    Private Sub btnCreateEWAM_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCreateEWAM.Click
        Try
            If Me.FormTemplateContext.GE.RecordID > 0 Then
                CreateEWActionMemo()
                CreatePDF()
            Else
                MsgBox("This record has not been created yet.  Please save the form to create the record")
            End If
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub
    Private Sub btnCreateEWCS_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCreateEWCS.Click
        Try
            If Me.FormTemplateContext.GE.RecordID > 0 Then
                CreateEWCaseSummaryMemo()
                CreatePDF()
            Else
                MsgBox("This record has not been created yet.  Please save the form to create the record")
            End If
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub
    Private Sub btnCreateCJCCS_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCreateCJCCS.Click
        Try
            If Me.FormTemplateContext.GE.RecordID > 0 Then
                CreateCJCCaseSummaryMemo()
                CreatePDF()
            Else
                MsgBox("This record has not been created yet.  Please save the form to create the record")
            End If
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub
    Private Sub btnCreateCJCDM_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCreateCJCDM.Click
        Try
            If Me.FormTemplateContext.GE.RecordID > 0 Then
                CreateCJCDirectorsMemo()
                CreatePDF()
            Else
                MsgBox("This record has not been created yet.  Please save the form to create the record")
            End If
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub
    Protected Overridable Sub GetGEData()
        Try
            CJCGE = m_oAppObj.GetEntityObject("ACSCJCMain", recordId)
            FirstName = CJCGE.GetValue("FirstName")
            MiddleName = CJCGE.GetValue("MiddleName")
            LastName = CJCGE.GetValue("LastName")
            City = CJCGE.GetValue("City")
            State = CJCGE.GetValue("State")
            DOB = CJCGE.GetValue("DOB")
            Age = CJCGE.GetValue("Age")
            FACSStatus = CJCGE.GetValue("FACSStatus")
            Specialty = CJCGE.GetValue("Specialty")
            StateActionDesc = CJCGE.GetValue("StateActionDesc")
            StateActionDate = CJCGE.GetValue("StateActionDate")
            CurrentStateLicensure = CJCGE.GetValue("CurrentStateLicensure")
            Charge = CJCGE.GetValue("Charge")
            Comments = CJCGE.GetValue("Comments")
            CaseName = CJCGE.GetValue("CaseName")
            AHBORFinalActionDate = CJCGE.GetValue("AHBORFinalActionDate")
            ACSCJCBORFinalID_Name = CJCGE.GetValue("ACSCJCBORFinalID_Name")

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub

    Private Sub CreateEWNewCaseMemo()
        Try
            saveLocation = saveLocalPrefix & "EWNewCaseMemo.pdf"
            GetGEData()


            PDFHeader = "<div style='page-break-after: always;'\><table width='100%'><tbody><TR><td><STRONG>RE:</STRONG></td><td colspan='2'> " & FirstName & " " & MiddleName & " " & LastName & " </td></tr>
                        <tr><td></td><td colspan='2'> " & City & " " & State & "</td></tr><tr><td></td><td>Date of Birth: " & DOB & "</td><td>Age: " & Age & " </td></tr><tr><td></td><td> 
                         Status: " & FACSStatus & "</td><td>Specialty: " & Specialty & "</td></tr></tbody></table> 
                        <CENTER><hr><table width='100%'><tbody><tr><td><P align=center><b><BR>This is the first consideration of this matter. <BR>    
                        After review and deliberation the CJC may recommend that the Executive Director issue a formal charge of Bylaws violation or recommend no further action.
                        <BR></b></P></td></tr></tbody></table><hr></CENTER>"

            PDFText = "<table width='100%'><tbody><tr><td>State of License:</td><td>State Action:</td><td>Date of State Action:</td><td>Current State Licensure:</td>  
                        </tr><tr><td>&nbsp;</td><td> " & StateActionDesc & "</td><td>" & StateActionDate & "</td><td>" & CurrentStateLicensure & "</td></tr></tbody></table> 
                        <table width='100%'><tbody><tr><td><br /><strong>Summary:</strong></td></tr><tr><td> " & Charge & "</td></tr><tbody></table></div>"


            PDFText2 = "<table width='100%'><tbody><tr><td><strong>Comments/Notes</strong></td></tr><tr><td>" & Comments & "</td></tr>
                        <tr><td><strong>CJC Action Recommendation</strong></td></tr><tr><td>&nbsp;</td></tr></tbody></table>
                        <table width='100%'><tr><td><input type='checkbox' id='cb1'><label for='cb1'>Bylaws violation (see below)</label></td></tr>
                        <tr><td><input type='checkbox' id='cb2'><label for='cb2'>No further action</label></td></tr>
                        <tr><td><input type='checkbox' id='cb3'><label for='cb3'>Other (specify)</td></tr></table>"

            PDFFooter = "<table><p><span style='font-family: Times;'>VII.&nbsp; Maintenance of Fellowship and Membership, Section 1.&nbsp; Discipline</span></p><ul>
                        <li style='box-sizing: border-box; font-size: 13px; vertical-align: baseline; background: none transparent scroll repeat 0% 0%; outline-width: 0px; outline-style: none; margin: 0px; outline-color: invert; border: 0px; padding: 0px;'><span style='font-family: Times; font-size: medium;'>Conviction of a felony or of any crime relating to or arising out of the practice of medicine, or involving moral turpitude. </span></li>
                        <li style='box-sizing: border-box; font-size: 13px; vertical-align: baseline; background: none transparent scroll repeat 0% 0%; outline-width: 0px; outline-style: none; margin: 0px; outline-color: invert; border: 0px; padding: 5px 0px 0px 0px;'><span style='font-family: Times; font-size: medium;'>Limitation or termination of any right associated with the practice of medicine in any state, province, or country, including the imposition of any requirement for surveillance, supervision, or review, by reason of violation of a medical practice act or other statute or governmental regulation, disciplinary action by any medical licensing authority, entry into a consent order, or voluntary surrender of license. </span></li>
                        <li style='box-sizing: border-box; font-size: 13px; vertical-align: baseline; background: none transparent scroll repeat 0% 0%; outline-width: 0px; outline-style: none; margin: 0px; outline-color: invert; border: 0px; padding: 5px 0px 0px 0px;'><span style='font-family: Times; font-size: medium;'>Improper financial dealings, including the payment or acceptance of rebates of fees for services or appliances, and the charging of exorbitant fees, or engaging in any activities which put personal financial consideration above the welfare of patients. </span></li>
                        <li style='box-sizing: border-box; font-size: 13px; vertical-align: baseline; background: none transparent scroll repeat 0% 0%; outline-width: 0px; outline-style: none; margin: 0px; outline-color: invert; border: 0px; padding: 5px 0px 0px 0px;'><span style='font-family: Times; font-size: medium;'>Participating in the deception of a patient. </span></li>
                        <li style='box-sizing: border-box; font-size: 13px; vertical-align: baseline; background: none transparent scroll repeat 0% 0%; outline-width: 0px; outline-style: none; margin: 0px; outline-color: invert; border: 0px; padding: 5px 0px 0px 0px;'><span style='font-family: Times; font-size: medium;'>Performance of unjustified surgery. </span></li>
                        <li style='box-sizing: border-box; font-size: 13px; vertical-align: baseline; background: none transparent scroll repeat 0% 0%; outline-width: 0px; outline-style: none; margin: 0px; outline-color: invert; border: 0px; padding: 5px 0px 0px 0px;'><span style='font-family: Times; font-size: medium;'>Unprofessional conduct. </span></li>
                        <li style='box-sizing: border-box; font-size: 13px; vertical-align: baseline; background: none transparent scroll repeat 0% 0%; outline-width: 0px; outline-style: none; margin: 0px; outline-color: invert; border: 0px; padding: 5px 0px 0px 0px;'><span style='font-family: Times; font-size: medium;'>The performance of surgical operations (except on patients whose chances of recovery would be prejudiced by removal to another hospital) under circumstances in which the responsibility for diagnosis or nonoperative care of the patient is delegated to another who is not fully qualified to undertake it. </span></li>
                        <li style='box-sizing: border-box; font-size: 13px; vertical-align: baseline; background: none transparent scroll repeat 0% 0%; outline-width: 0px; outline-style: none; margin: 0px; outline-color: invert; border: 0px; padding: 5px 0px 0px 0px;'><span style='font-family: Times; font-size: medium;'>Failure or refusal to cooperate reasonably with an investigation by the College of a disciplinary matter. </span></li>
                        <li style='box-sizing: border-box; font-size: 13px; vertical-align: baseline; background: none transparent scroll repeat 0% 0%; outline-width: 0px; outline-style: none; margin: 0px; outline-color: invert; border: 0px; padding: 5px 0px 0px 0px;'><span style='font-family: Times; font-size: medium;'>Participating in communications to the College, to the public, or as part of a judicial process which convey false, untrue, deceptive, or misleading information through statements, testimonials, photographs, graphics, or other means, or which omit material information without which the communication is deceptive.</span></li>
                        </ul></table>"
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub
    Private Sub CreateEWDirectorsMemo()
        Try
            saveLocation = saveLocalPrefix & "EWDirectorsMemo.pdf"
            GetGEData()


            PDFHeader = "<table width='100%'><tbody><TR><td><STRONG>RE:</STRONG></td><td colspan='2'> " & FirstName & " " & MiddleName & " " & LastName & " </td></tr>
                           <tr><td></td><td colspan='2'>" & City & " " & State & "</td></tr><tr><td></td><td>Date of Birth: " & DOB & "</td><td>Age: " & Age & " </td></tr>
                           <tr><td></td><td> Status: " & FACSStatus & "</td><td>Specialty: " & Specialty & "</td></tr></tbody></table><CENTER><hr></CENTER>"

            PDFText = "<table width='100%'><tbody><tr><td><STRONG>CaseName:</STRONG></td><td>" & CaseName & "</td></tr>
                        <tr><td><STRONG>Case Jurisdiction:</STRONG></td><td>??</td><td>Current State of Licensure:</td><td>" & CurrentStateLicensure & "
                        </td></tr></tbody></table><hr></CENTER><table width='100%'><tbody><tr><td><BR><STRONG>Summary:</STRONG></td></tr><tr><td>" & Charge & "</td></tr></tbody></table>"

            PDFText2 = ""

            PDFFooter = ""
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub
    Private Sub CreateEWCaseSummaryMemo()
        Try
            saveLocation = saveLocalPrefix & "EWCaseSummaryMemo.pdf"
            GetGEData()


            PDFHeader = "<table width='100%'><tbody><tr><td colspan='2'>Name: " & FirstName & " " & MiddleName & " " & LastName & "</td>
                        </tr><tr><td colspan='2'>City: " & City & " State: " & State & "</td></tr></tbody></table><center><hr /></center>
                        <table width='100%'><tbody><tr><td><br /><strong>Summary:</strong></td></tr><tr><td>" & Charge & " </td></tr></tbody></table>"

            PDFText = "<center><table style='background-color: #d3d3d3;' width='25%'>'<tbody><tr><td><strong>Fellowship Status:</strong></td>
                    <td>" & FACSStatus & "</td></tr><tr><td><strong>Date of Birth:</strong></td><td>" & DOB & "</td></tr>
                    <tr><td><strong>Age:</strong></td><td>" & Age & "</td></tr><tr><td><strong>Surgical Specialty:</strong></td>
                    <td>" & Specialty & "</td></tr></tbody></table></center>"

            PDFText2 = ""

            PDFFooter = ""
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub
    Private Sub CreateEWActionMemo()
        Try
            saveLocation = saveLocalPrefix & "EWActionMemo.pdf"
            GetGEData()
            getActionHistory()

            PDFHeader = "<table width='100%'><tbody><TR><td><STRONG>RE:</STRONG></td><td colspan='2'>" & FirstName & " " & MiddleName & " " & LastName & "</td></tr>
                        <tr><td></td><td colspan='2'>" & City & "," & State & "</td></tr><tr><td></td><td>
                        Date of Birth: " & DOB & "</td><td>Age: " & Age & "</td></tr><tr><td></td><td>
                        Status: " & FACSStatus & "</td><td>Specialty: " & Specialty & " </td></tr></tbody></table><CENTER><hr>
                        <table width='100%'><tbody><TR><td align='center'>At its last meeting, the CJC charged the surgeon named above with the following violation. </td</tr></table>"

            'PDFText = AHText
            '-> need to add dt of action history here

            PDFText2 = "<table width ='100%'><tbody><tr><td align='center'>After review of information provided by the Fellow (attached), the 
                       Committee may recommend specific&nbsp;disciplinary action.</td></tr><tr><td align='center'> <BR>The Case summary below Is provided For your review </td></tr></tbody></table><hr></CENTER>"

            PDFFooter = "<table width='100%'><tbody><tr><td align='center'><BR><STRONG>Case Summary</STRONG></td></tr>
                        <tr><td><BR><STRONG>Summary:</STRONG></td></tr><tr><td>" & Charge & "<BR></td></tr>
                        <tr><td><STRONG>Comments/Notes</STRONG></td></tr><tr><td>" & Comments & " <BR></td></tr>
                        <tr><td><STRONG>CJC Action Recommendation</STRONG></td></tr><tr><td></td></tr></tbody></table>"

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub
    Private Sub CreateCJCCaseSummaryMemo()
        Try
            saveLocation = saveLocalPrefix & "CJCCaseSummaryMemo.pdf"
            GetGEData()


            PDFHeader = "<table width='100%'><tbody><TR><td><STRONG>RE:</STRONG></td><td colspan='2'>" & FirstName & " " & MiddleName & " " & LastName & "</td></tr>
                        <tr><td></td><td colspan='2'>" & City & "," & State & "</td></tr><tr><td></td><td>
                        Date of Birth: " & DOB & "</td><td>Age: &lt;&lt;Age&gt;&gt;</td></tr><tr><td></td><td>
                        Status: " & FACSStatus & "</td><td>Specialty: " & Specialty & " </td></tr></tbody></table><CENTER><hr>"

            PDFText = "<table width='100%'><tbody><tr><td><STRONG>Final Action Date:</STRONG>  " & AHBORFinalActionDate & "  </td><td><STRONG>Final Action: 
                        </STRONG>" & ACSCJCBORFinalID_Name & "<STRONG></STRONG></td></tr></tbody></table><hr></CENTER>
	                    <table width='100%'><tbody><tr><td>" & Charge & "</td></tr></tbody></table>"

            PDFText2 = ""

            PDFFooter = ""
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub
    Private Sub CreateCJCDirectorsMemo()
        Try
            saveLocation = saveLocalPrefix & "CJCDirectorsMemo.pdf"
            GetGEData()
            getStaffRec()

            PDFHeader = "<table width='100%'><tbody><TR><td><STRONG>RE:</STRONG></td><td colspan='2'>" & FirstName & " " & MiddleName & " " & LastName & "</td></tr>
                        <tr><td></td><td colspan='2'>" & City & "," & State & "</td></tr><tr><td></td><td>
                        Date of Birth: " & DOB & "</td><td>Age: &lt;&lt;Age&gt;&gt;</td></tr><tr><td></td><td>
                        Status: " & FACSStatus & "</td><td>Specialty: " & Specialty & " </td></tr></tbody></table><CENTER><hr>"

            PDFText = "<table width='100%'><tbody><tr><td>State Action:</td><td>" & StateActionDesc & "</td></tr>
                        <tr><td>Date of State Action:</td><td>" & StateActionDate & "</td><td>Current State Licensure:</td><td>" & CurrentStateLicensure & "</td></tr></tbody></table>
                        </CENTER><table width='100%'><tbody><tr><td><STRONG><BR>Summary:</STRONG></td></tr><tr><td>" & Charge & "</td></tr></tbody></table>"

            'PDFText2 = ""

            PDFFooter = ""
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub
    Private Sub CheckAttachment()
        Try
            entityIdSql = "select ID from Entities where name like 'ACSCJCMain'"
            entityId = Convert.ToInt32(m_oDA.ExecuteScalar(entityIdSql))
            checkAttachmentIdSql = "select * from attachment where entityid = " & entityId & " And recordid = " & recordId
            checkAttachmentId = m_oDA.GetDataTable(checkAttachmentIdSql)
            If checkAttachmentId.Rows.Count > 0 Then
                For Each dr As DataRow In checkAttachmentId.Rows
                    If dr.Item("LocalFileName") = saveLocalPrefix & "EWNewCaseMemo.pdf" Then
                        btnCreateEWNCM.Visible = False
                    End If
                    If dr.Item("LocalFileName") = saveLocalPrefix & "EWDirectorsMemo.pdf" Then
                        btnCreateEWDM.Visible = False
                    End If
                    If dr.Item("LocalFileName") = saveLocalPrefix & "EWCaseSummaryMemo.pdf" Then
                        btnCreateEWCS.Visible = False
                    End If
                    If dr.Item("LocalFileName") = saveLocalPrefix & "EWActionMemo.pdf" Then
                        btnCreateEWAM.Visible = False
                    End If
                    If dr.Item("LocalFileName") = saveLocalPrefix & "CJCCaseSummaryMemo.pdf" Then
                        btnCreateCJCCS.Visible = False
                    End If
                    If dr.Item("LocalFileName") = saveLocalPrefix & "CJCDirectorsMemo.pdf" Then
                        btnCreateCJCDM.Visible = False
                    End If
                Next
            End If

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try

    End Sub

    Private Sub getActionHistory()
        Try
            actionHistoryIdSql = "Select AHCharge from ACSCJCActionHistory where ACSCJCMainID = " & recordId
            ahdt = m_oDA.GetDataTable(actionHistoryIdSql)
            AHRow = ""
            If ahdt.Rows.Count > 0 Then
                'AHHeader = "<table width='100%'><tbody>"
                For Each dr As DataRow In ahdt.Rows
                    If dr.Item("AHCharge") IsNot Nothing Then
                        AHCharge = dr.Item("AHCharge")
                        AHRow = AHRow & "<table width='100%'><tbody><tr><td align='center'>" & AHCharge & " </td></tr></tbody></table>"
                        'AHRow = "<table width='100%'><tbody><tr><td>" & AHCharge & " </td></tr></tbody></table>"
                    End If

                    PDFText = AHRow
                Next
                'AHFooter = "</tbody></table>"

            End If

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub

    Private Sub getStaffRec()
        Try
            staffRecSql = "Select * from ACSCJCStaffRec where Id <> 10"
            srdt = m_oDA.GetDataTable(staffRecSql)
            SRRow = "<Hr><table width='100%' height='40px'><tbody><tr><td align='Left'>Director's Recommendation: </td></tr><tr><td align='left'>"
            If srdt.Rows.Count > 0 Then
                'AHHeader = "<table width='100%'><tbody>"
                For Each dr As DataRow In srdt.Rows
                    If dr.Item("ID") IsNot Nothing Then
                        SRName = dr.Item("Name")
                        SRRow = SRRow & SRName & "<br>"

                    End If
                    Dim SRFooter As String = "</td></tr></table>"
                    PDFText2 = SRRow & SRFooter
                Next
                'AHFooter = "</tbody></table>"

            End If

        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub
    Protected Sub CreatePDF()
        CheckKey()
        Dim renderer = New IronPdf.HtmlToPdf()
        Dim document = renderer.RenderHtmlAsPdf(PDFHeader & PDFText & PDFText2 & PDFFooter)

        document.SaveAs(saveLocation)
        CreateAttachment()
    End Sub
    Private Sub CreateAttachment()
        Try
            entityIdSql = "select ID from Entities where name like 'ACSCJCMain'"
            entityId = Convert.ToInt32(m_oDA.ExecuteScalar(entityIdSql))
            attachmentCatIdSql = "select ID from vwAttachmentCategories where name like 'CJC'"
            attachmentCatId = Convert.ToInt32(m_oDA.ExecuteScalar(attachmentCatIdSql))
            filename = Path.GetFileName(saveLocation)
            data = File.ReadAllBytes(saveLocation)

            AttachmentsGE = m_oAppObj.GetEntityObject("Attachments", -1)
            AttachmentsGE.SetValue("Name", filename)
            AttachmentsGE.SetValue("Description", "")
            AttachmentsGE.SetValue("EntityID", entityId)
            AttachmentsGE.SetValue("RecordID", recordId)
            AttachmentsGE.SetValue("CategoryID", attachmentCatId)
            AttachmentsGE.SetValue("LocalFileName", saveLocation)
            AttachmentsGE.SetValue("BlobData", data)

            If AttachmentsGE.IsDirty Then
                If Not AttachmentsGE.Save(False) Then
                    Throw New Exception("Problem Saving Record:" & AttachmentsGE.RecordID)
                    result = "Error"
                Else
                    AttachmentsGE.Save(True)
                    result = "Success"
                    attachId = AttachmentsGE.RecordID
                End If

            End If

            If result = "Success" Then
                SaveAttachmentBlob()
            End If
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub
    Protected Overridable Sub SaveAttachmentBlob()
        Try
            Dim dp = New IDataParameter(1) {}
            dp(0) = m_oDA.GetDataParameter("@ID", SqlDbType.BigInt, attachId)
            dp(1) = m_oDA.GetDataParameter("@BLOBData", SqlDbType.Image, data.Length, data)
            m_oDA.ExecuteNonQueryParametrized("Aptify.dbo.spInsertAttachmentBlob", CommandType.StoredProcedure, dp)
            RemoveLocalFile()
        Catch ex As Exception
            Aptify.Framework.ExceptionManagement.ExceptionManager.Publish(ex)
        End Try
    End Sub

    Private Sub RemoveLocalFile()
        Try
            Dim FileToDelete As String

            FileToDelete = saveLocation

            If System.IO.File.Exists(FileToDelete) = True Then

                System.IO.File.Delete(FileToDelete)
                'MsgBox("File Deleted")

            End If
            FormTemplateContext.GE.Save()
            CheckAttachment()
        Catch ex As Exception

        End Try
    End Sub

End Class
