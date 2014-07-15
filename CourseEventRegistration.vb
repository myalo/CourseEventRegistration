Imports System
Imports System.Configuration
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Xml
Imports System.Runtime.InteropServices
Namespace CourseEventRegistration
#Region "Registration"
	Public Class Registration
		''' <summary>
		''' Changes made between old computer to 6520. Before adding to git
		''' add new property ProcessorRemarks
		''' in studentDC.Insert add dim errorMessage as string and in try con.open() add to the catch errorMessage = ex.Message.
		''' This was only done for debugging so I can break there and see the error. For now I am leaving it there.
		''' </summary>
		''' <remarks></remarks>
		Private _registrationId As Long
		Private _studentId As Long
		Private _registrationFee As Decimal
		Private _submittedOn As Date
		Private _errorMessage As String
		Private _connectionString As String
		Private _processedOn As Date
		Private _processedBy As String
		Private _processorsRemarks As String
		Private _coursesTaken As Collection
		Private _status As String 'Submitted, Paid, Reviewed
		Private _statusid As Integer
		Private _student As Student
		Private _statuses As Collection
		Public Sub New(ByVal ConnectionString As String)
			_coursesTaken = New Collection
			_statuses = New Collection
			Status = "Submitted"
			StatusId = 1
			SubmittedOn = Now()
			RegistrationId = 0
			_student = New Student
			ProcessedBy = ""
			ProcessedOn = Nothing
			Me.ConnectionString = ConnectionString
			If String.IsNullOrEmpty(ConnectionString) = False Then
				Dim cs As RegistrationStatus
				Dim dt As DataTable
				Dim dc As New RegistrationDC
				dt = dc.LoadCourseRegistrationStatus(ConnectionString)
				For i As Integer = 0 To dt.Rows.Count - 1
					cs = New RegistrationStatus
					cs.Description = dt.Rows(i)("description")
					cs.StatusId = dt.Rows(i)("Statusid")
					cs.SystemName = dt.Rows(i)("SystemName")
					Statuses.Add(cs)
				Next
				Status = GetStatusById(StatusId)
			End If
		End Sub
		Public Property CoursesTaken() As Collection
			Get
				Return _coursesTaken
			End Get
			Set(ByVal value As Collection)
				_coursesTaken = value
			End Set
		End Property
		Public Property Statuses() As Collection
			Get
				Return _statuses
			End Get
			Set(ByVal value As Collection)
				_statuses = value
			End Set
		End Property
		Public Property Student() As Student
			Get
				Return _student
			End Get
			Set(ByVal value As Student)
				_student = value
			End Set
		End Property
		Public Property RegistrationId() As Integer
			Get
				RegistrationId = _registrationId
			End Get
			Set(ByVal Value As Integer)
				_registrationId = Value
			End Set
		End Property
		Public Property StudentId() As Integer
			Get
				StudentId = _studentId
			End Get
			Set(ByVal Value As Integer)
				_studentId = Value
			End Set
		End Property
		Public Property StatusId() As Integer
			Get
				StatusId = _statusid
			End Get
			Set(ByVal Value As Integer)
				_statusid = Value
			End Set
		End Property
		Public Property Status() As String
			Get
				Status = _status
			End Get
			Set(ByVal Value As String)
				_status = Value
				_statusid = GetStatusIdByStatusDescription(Value)
			End Set
		End Property
		Public Property RegistrationFees() As Decimal
			Get
				Return _registrationFee
			End Get
			Set(ByVal value As Decimal)
				_registrationFee = value
			End Set
		End Property
		Public Property SubmittedOn() As Date
			Get
				SubmittedOn = _submittedOn
			End Get
			Set(ByVal Value As Date)
				_submittedOn = Value
			End Set
		End Property
		Public Property ProcessedOn() As Date
			Get
				ProcessedOn = _processedOn
			End Get
			Set(ByVal Value As Date)
				_processedOn = Value
			End Set
		End Property
		Public Property ProcessedBy() As String
			Get
				ProcessedBy = _processedBy
			End Get
			Set(ByVal Value As String)
				_processedBy = Value
			End Set
		End Property

		Public Property ProcessorRemarks() As String
			Get
				Return _processorsRemarks
			End Get
			Set(ByVal value As String)
				_processorsRemarks = value
			End Set
		End Property

		Public ReadOnly Property TotalAmount() As Decimal
			Get
				Dim t As Decimal
				Dim c As CourseEventData
				For i As Integer = 1 To CoursesTaken.Count
					c = CoursesTaken.Item(i)
					t += c.Tuition
				Next
				TotalAmount = t
			End Get
		End Property
		Public Property ErrorMessage() As String
			Get
				ErrorMessage = _errorMessage
			End Get
			Set(ByVal Value As String)
				_errorMessage = Value
			End Set
		End Property
		Public Property ConnectionString() As String
			Get
				ConnectionString = _connectionString
			End Get
			Set(ByVal Value As String)
				_connectionString = Value
			End Set
		End Property
		Private Sub ClearRegistration()
			_coursesTaken = New Collection
			Status = "Submitted"
			StatusId = 1
			SubmittedOn = Now()
			RegistrationId = 0
			_student = New Student
			ProcessedBy = ""
			ProcessedOn = Nothing
		End Sub
		Private Function GetStatusById(ByVal StatusId As Integer) As String
			Dim cs As New RegistrationStatus
			Dim Status As String = ""
			If Not Statuses Is Nothing Then
				For i As Integer = 1 To Statuses.Count
					cs = Statuses(i)
					If cs.StatusId = StatusId Then
						Status = cs.Description
						Exit For
					End If
				Next
			End If
			GetStatusById = Status
		End Function
		Private Function GetStatusIdByStatusDescription(ByVal StatusDescription As String) As Integer
			Dim cs As New RegistrationStatus
			Dim StatusId As Integer = 0
			If Not Statuses Is Nothing Then
				For i As Integer = 1 To Statuses.Count
					cs = Statuses(i)
					If cs.Description.ToLower = StatusDescription.ToLower Then
						StatusId = cs.StatusId
						Exit For
					End If
				Next
			End If
			GetStatusIdByStatusDescription = StatusId
		End Function
		Public Sub LoadRegistration(ByVal RegistrationId As Integer)
			Dim oDC As New RegistrationDC
			Dim dsRegistration As New DataSet
			Dim dt As DataTable
			dsRegistration = oDC.GetRegistrationSetById(RegistrationId, Me.ConnectionString)
			If dsRegistration.Tables.Count > 0 Then
				'fill the properties with 
				If dsRegistration.Tables(0).Rows.Count > 0 Then
					'assume 1 row
					dt = dsRegistration.Tables(0)
					RegistrationId = dt.Rows(0)("CORegistrationID")
					StudentId = dt.Rows(0)("StudentID")
					SubmittedOn = dt.Rows(0)("SubmittedOn")
					StatusId = dt.Rows(0)("StatusId")
					Status = GetStatusById(StatusId)
					Dim ce As CourseEventData
					Dim Student As Student
					dt = dsRegistration.Tables(1)
					For i As Integer = 0 To dt.Rows.Count - 1
						ce = New CourseEventData
						With ce
							.CourseId = dt.Rows(i)("CourseId")
							.CourseNumber = dt.Rows(i)("CourseNumber")
							.NumberOfSeats = dt.Rows(i)("NumberOfSeats")
							.RegistrationFee = dt.Rows(i)("RegistrationFee")
							.RegistrationId = RegistrationId
							.Tuition = dt.Rows(i)("Tuition")
							CoursesTaken.Add(ce)
						End With
					Next
					dt = dsRegistration.Tables(2)
					If dt.Rows.Count > 0 Then
						Student = New Student
						With Student
							.StudentId = dt.Rows(0)("StudentId")
							.FirstName = dt.Rows(0)("StudentFirstName")
							.LastName = dt.Rows(0)("StudentLastName")
							.Email = dt.Rows(0)("Email")
							.Address = dt.Rows(0)("Address")
							.City = dt.Rows(0)("City")
							.State = dt.Rows(0)("State")
							.ZipCode = dt.Rows(0)("ZipCode")
							.WorkDayPhone = dt.Rows(0)("WorkDayPhone")
							.HomeEveningPhone = dt.Rows(0)("HomeEveningPhone")
							.WorkDayPhone = dt.Rows(0)("WorkDayPhone")
							.CellPhone = dt.Rows(0)("CellPhone")
							.DateFirstEnrolled = dt.Rows(0)("DateFirstEnrolled")
							.isSubscribed = False
							If IsDBNull(dt.Rows(0)("Subscribed")) = False AndAlso dt.Rows(0)("Subscribed") = 1 Then
								.isSubscribed = True
							End If
						End With
						Me.Student = Student
					End If

				End If
			Else
				Me.ClearRegistration()
			End If
		End Sub
		''' <summary>
		''' Current logic: New status need to be the next status. FOr example from Subimetted to Paid or from Paid to In Studentmanager
		''' </summary>
		''' <param name="CourseRegistrationId"></param>
		''' <param name="NewStatus"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function ChangeStatus(ByVal CourseRegistrationId As Integer, ByVal NewStatus As String) As Boolean
			Dim result As Boolean = False
			Dim dc As New RegistrationDC
			'get current status for Business rules
			Dim dt As DataTable
			Dim currentStatusId As Integer
			Dim newStatusId As Integer = 0
			Dim results As Boolean = False
			dt = dc.GetRegistrationById(CourseRegistrationId, ConnectionString)
			If dt.Rows.Count > 0 Then
				currentStatusId = dt.Rows(0)("StatusId")
				'get the status Id for new status
				newStatusId = GetStatusIdByStatusDescription(NewStatus)
				If currentStatusId + 1 <> newStatusId Then
				Else
					'all ok
					results = dc.ChangeStatus(CourseRegistrationId, newStatusId, Me.ConnectionString)
				End If
			End If
			ChangeStatus = results
		End Function
		''' <summary>
		''' Current logic: New status need to be the next status. FOr example from Subimetted to Paid or from Paid to In Studentmanager
		''' this is overloading with additional parameter ReferenceNumber (normally the PNRER)
		''' </summary>
		''' <param name="CourseRegistrationId"></param>
		''' <param name="NewStatus"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function ChangeStatus(ByVal CourseRegistrationId As Integer, ByVal NewStatus As String, ByVal PaymentReferenceNumber As String) As Boolean
			Dim result As Boolean = False
			Dim dc As New RegistrationDC
			'get current status for Business rules
			Dim dt As DataTable
			Dim currentStatusId As Integer
			Dim newStatusId As Integer = 0
			Dim results As Boolean = False
			dt = dc.GetRegistrationById(CourseRegistrationId, ConnectionString)
			If dt.Rows.Count > 0 Then
				currentStatusId = dt.Rows(0)("StatusId")
				'get the status Id for new status
				newStatusId = GetStatusIdByStatusDescription(NewStatus)
				If currentStatusId + 1 <> newStatusId Then
				Else
					'all ok
					results = dc.ChangeStatus(CourseRegistrationId, newStatusId, PaymentReferenceNumber, Me.ConnectionString)
				End If
			End If
			ChangeStatus = results
		End Function
		''' <summary>
		''' 
		''' </summary>
		''' <returns>newly create CourseRegistrationId</returns>
		''' <remarks>No transaction at this point
		''' Although the object has all the information and the caller can call any of the properties after the save it will return the new reigstrationId for simplicity
		''' </remarks>
		Public Function save() As Integer
			Dim ErrMsg As String = ""
			Dim bReturn As Boolean = True	'Returning results
			Dim dc As New RegistrationDC
			Dim oStudent As New Student
			Dim iNewStudentId As Integer = 0
			Dim iNewRegistrationId As Integer = 0
			Dim ReturnValue As Integer = 0
			'oStudent.connectionString = Me.ConnectionString

			If _connectionString = "" Then
				bReturn = False
			Else
				If Not Me.Student Is Nothing AndAlso Not Me.CoursesTaken Is Nothing AndAlso Me.CoursesTaken.Count > 0 Then
					'save the student
					iNewStudentId = _student.CreateStudent(Me.Student)
					If iNewStudentId = 0 Then
						bReturn = False
					End If
				End If
				If iNewStudentId > 0 Then
					Me.StudentId = iNewStudentId
					'save registration
					iNewRegistrationId = dc.insert(Me, Me.ConnectionString)
					If iNewRegistrationId > 0 Then
						Me.RegistrationId = iNewRegistrationId
						'Save courseTaken
						Dim iNewCourseTakenId As Integer
						Dim cedata As New CourseEventData
						For i As Integer = 1 To Me.CoursesTaken.Count
							cedata = CoursesTaken(i)
							cedata.RegistrationId = Me.RegistrationId
							iNewCourseTakenId = dc.insertCourseTaken(cedata, Me.ConnectionString)
							If iNewCourseTakenId = 0 Then
								bReturn = False
								Exit For
							End If
						Next
					End If
				End If
			End If
			If bReturn = True Then
				ReturnValue = iNewRegistrationId
			End If
			save = ReturnValue
		End Function
		Public Function AddCourseEvent(ByVal oCourseEvent As CourseEventData) As Boolean
			'Return 1
			Dim bReturn As Boolean
			Dim oBR As BrokenRules
			oBR = New BrokenRules
			bReturn = True
			ErrorMessage = ""
			'Verify mandatory fields and set default
			'CourseId
			If oCourseEvent.CourseId = 0 Then
				oBR.Add("CourseId is mandatory")
			End If
			If oCourseEvent.CourseNumber = "" Then
				oBR.Add("Course# is mandatory")
			End If
			'Number of seats 0 then 1
			If oCourseEvent.NumberOfSeats = 0 Then
				oCourseEvent.NumberOfSeats = 1
			End If
			'Ignore - will be set by the save
			If oCourseEvent.RegistrationId = 0 Then
			End If
			If oCourseEvent.Tuition = 0 Then
				oBR.Add("Tuition cannot be 0")
			End If
			If oBR.IsValid = True Then
				'Make sure uniqu CourseId
				Dim oCT As CourseEventData

				If _coursesTaken.Count > 0 Then
					For Each oCT In _coursesTaken
						If Not oCT Is Nothing Then
							If oCT.CourseId = oCourseEvent.CourseId Then
								oBR.Add("Duplicate courses not allowed")
								Exit For
							End If
						End If
					Next
				End If
				_coursesTaken.Add(oCourseEvent)
			End If
			ErrorMessage = oBR.ToString
			AddCourseEvent = oBR.IsValid
		End Function
		''' <summary>
		''' Add a course object to the CourseRegistration object. 
		''' can Ignore rule Tuition > 0
		''' </summary>
		''' <param name="oCourseEvent"></param>
		''' <param name="IgnoreTuitionZero"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function AddCourseEvent(ByVal oCourseEvent As CourseEventData, ByVal IgnoreTuitionZero As Boolean) As Boolean
			'Return 1
			Dim bReturn As Boolean
			Dim oBR As BrokenRules
			oBR = New BrokenRules
			bReturn = True
			ErrorMessage = ""
			'Verify mandatory fields and set default
			'CourseId
			If oCourseEvent.CourseId = 0 Then
				oBR.Add("CourseId is mandatory")
			End If
			If oCourseEvent.CourseNumber = "" Then
				oBR.Add("Course# is mandatory")
			End If
			'Number of seats 0 then 1
			If oCourseEvent.NumberOfSeats = 0 Then
				oCourseEvent.NumberOfSeats = 1
			End If
			'Ignore - will be set by the save
			If oCourseEvent.RegistrationId = 0 Then
			End If
			If oCourseEvent.Tuition = 0 And IgnoreTuitionZero = False Then
				oBR.Add("Tuition cannot be 0")
			End If
			If oBR.IsValid = True Then
				'Make sure uniqu CourseId
				Dim oCT As CourseEventData

				If _coursesTaken.Count > 0 Then
					For Each oCT In _coursesTaken
						If Not oCT Is Nothing Then
							If oCT.CourseId = oCourseEvent.CourseId Then
								oBR.Add("Duplicate courses not allowed")
								Exit For
							End If
						End If
					Next
				End If
				_coursesTaken.Add(oCourseEvent)
			End If
			ErrorMessage = oBR.ToString
			AddCourseEvent = oBR.IsValid
		End Function

	End Class
#End Region
#Region "CourseEventData"

   Public Class CourseEventData
      Private m_CourseNumber As String
      Private _registrationId As Long
      Private m_Tuition As Decimal
      Private m_NumberOfSeats As Decimal
      Private _registrationFee As Decimal
      Private m_CourseId As Long
      Public Property CourseId() As Long
         Get
            Return m_CourseId
         End Get
         Set(ByVal value As Long)
            m_CourseId = value
         End Set
      End Property
      Public Property RegistrationId() As Long
         Get
            Return _registrationId
         End Get
         Set(ByVal value As Long)
            _registrationId = value
         End Set
      End Property
      Public Property Tuition() As Decimal
         Get
            Return m_Tuition
         End Get
         Set(ByVal value As Decimal)
            m_Tuition = value
         End Set
      End Property
      Public Property NumberOfSeats() As Integer
         Get
            Return m_NumberOfSeats
         End Get
         Set(ByVal value As Integer)
            m_NumberOfSeats = value
         End Set
      End Property
      Public Property RegistrationFee() As Decimal
         Get
            Return _registrationFee
         End Get
         Set(ByVal value As Decimal)
            _registrationFee = value
         End Set
      End Property
      Public Property CourseNumber() As String
         Get
            Return m_CourseNumber
         End Get
         Set(ByVal value As String)
            m_CourseNumber = value
         End Set
      End Property

   End Class
#End Region

   Public Class Student
#Region "Student properties"
      Private _studentId As Integer
      Private _firstName As String
      Private _lastName As String
      Private _middleName As String
      Private _address As String
      Private _city As String
      Private _state As String
      Private _zipCode As String
      Private _workDayPhone As String
      Private _homeEveningPhone As String
      Private _cellPhone As String
      Private _ssn As String
      Private _email As String
      Private _gender As String
      Private _dateFirstEnrolled As Date
      Private _studentTypeId As Integer
      Private _password As String
      Private _issubscribed As Boolean
      Private _studentManagerId As String
      Private _connectionString As String
      Public Property StudentId() As Integer
         Get
            Return _StudentId
         End Get
         Set(ByVal value As Integer)
            _studentId = value
         End Set
      End Property
      Public Property FirstName() As String
         Get
            Return _firstName
         End Get
         Set(ByVal value As String)
            _firstName = value
         End Set
      End Property
      Public Property LastName() As String
         Get
            Return _lastName
         End Get
         Set(ByVal value As String)
            _lastName = value
         End Set
      End Property
      Public Property MiddleName() As String
         Get
            Return _middleName
         End Get
         Set(ByVal value As String)
            _middleName = value
         End Set
      End Property
      Public Property Address() As String
         Get
            Return _address
         End Get
         Set(ByVal value As String)
            _address = value
         End Set
      End Property
      Public Property City() As String
         Get
            Return _city
         End Get
         Set(ByVal value As String)
            _city = value
         End Set
      End Property
      Public Property State() As String
         Get
            Return _state
         End Get
         Set(ByVal value As String)
            _state = value
         End Set
      End Property
      Public Property ZipCode() As String
         Get
            Return _zipCode
         End Get
         Set(ByVal value As String)
            _zipCode = value
         End Set
      End Property
      Public Property WorkDayPhone() As String
         Get
            Return _workDayPhone
         End Get
         Set(ByVal value As String)
            _workDayPhone = value
         End Set
      End Property
      Public Property HomeEveningPhone() As String
         Get
            Return _homeEveningPhone
         End Get
         Set(ByVal value As String)
            _homeEveningPhone = value
         End Set
      End Property
      Public Property CellPhone() As String
         Get
            Return _cellPhone
         End Get
         Set(ByVal value As String)
            _cellPhone = value
         End Set
      End Property
      Public Property ssn() As String
         Get
            Return _ssn
         End Get
         Set(ByVal value As String)
            _ssn = value
         End Set
      End Property
      Public Property Email() As String
         Get
            Return _email
         End Get
         Set(ByVal value As String)
            _email = value
         End Set
      End Property
      Public Property Gender() As String
         Get
            Return _gender
         End Get
         Set(ByVal value As String)
            _gender = value
         End Set
      End Property
      Public Property DateFirstEnrolled() As Date
         Get
            Return _dateFirstEnrolled
         End Get
         Set(ByVal value As Date)
            _dateFirstEnrolled = value
         End Set
      End Property
      Public Property StudentTypeId() As Integer
         Get
            Return _studentTypeId
         End Get
         Set(ByVal value As Integer)
            _studentTypeId = value
         End Set
      End Property
      Public Property Password() As String
         Get
            Return _password
         End Get
         Set(ByVal value As String)
            _password = value
         End Set
      End Property
      Public Property isSubscribed() As Boolean
         Get
            Return _issubscribed
         End Get
         Set(ByVal value As Boolean)
            _issubscribed = value
         End Set
      End Property
      Public Property StudentManagerId() As String
         Get
            Return _studentManagerId
         End Get
         Set(ByVal value As String)
            _studentManagerId = value
         End Set
      End Property
      Public Property connectionString() As String
         Get
            Return _connectionString
         End Get
         Set(ByVal value As String)
            _connectionString = value
         End Set
      End Property
#End Region
#Region "Student methods"
      ''' <summary>
      ''' Check if one exists - logic in the stored procedure
      ''' </summary>
      ''' <param name="student"></param>
      ''' <returns></returns>
      ''' <remarks></remarks>
      Public Function CreateStudent(ByRef student As Student) As Integer
         Dim iNewId As Integer = 0
         Dim dc As New StudentDC
         Dim existingStudent As New Student
         Dim dt As DataTable
         Dim iRowsAffected As Integer = 0
         Dim bUpdate As Boolean = False
         iNewId = dc.IsStudentExists(student.LastName, student.FirstName, student.Address, student.City, student.State, student.ZipCode, student.Email, student.WorkDayPhone, student.HomeEveningPhone, student.CellPhone, Me.connectionString)
         If iNewId = 0 Then
            iNewId = dc.insert(student, connectionString)
            CreateStudent = iNewId
         Else
            'update Student with newer information
            'load the Student
            dt = dc.Load(iNewId, Me.connectionString)
            If Not dt Is Nothing AndAlso dt.Rows.Count() > 0 Then
               student.StudentId = iNewId
               With existingStudent
                  .FirstName = dt.Rows(0)("StudentFirstName")
                  .LastName = dt.Rows(0)("StudentLastName")
                  .Address = dt.Rows(0)("Address")
                  .City = dt.Rows(0)("City")
                  .State = dt.Rows(0)("State")
                  .ZipCode = dt.Rows(0)("ZipCode")
                  .WorkDayPhone = dt.Rows(0)("WorkDayPhone")
                  .HomeEveningPhone = dt.Rows(0)("HomeEveningPhone")
                  .CellPhone = dt.Rows(0)("CellPhone")
                  .Email = dt.Rows(0)("Email")
                  .DateFirstEnrolled = dt.Rows(0)("DateFirstEnrolled")
                  .StudentId = iNewId
               End With
               With existingStudent
                  If CompareFields(.Address, student.Address) = False Then
                     bUpdate = True
                  End If
                  If CompareFields(.City, student.City) = False Then
                     bUpdate = True
                  End If
                  If CompareFields(.State, student.State) = False Then
                     bUpdate = True
                  End If
                  If CompareFields(.ZipCode, student.ZipCode) = False Then
                     bUpdate = True
                  End If
                  If CompareFields(.Email, student.Email) = False Then
                     bUpdate = True
                  End If
                  If CompareFields(.WorkDayPhone, student.WorkDayPhone) = False Then
                     bUpdate = True
                  End If
                  If CompareFields(.HomeEveningPhone, student.HomeEveningPhone) = False Then
                     bUpdate = True
                  End If
                  If CompareFields(.CellPhone, student.CellPhone) = False Then
                     bUpdate = True
                  End If

               End With
            End If
            If bUpdate = True Then
               iRowsAffected = dc.Update(student, iNewId, Me.connectionString)
            End If
         End If
         CreateStudent = iNewId
      End Function
      Private Function CompareFields(ByVal FieldA As Object, ByVal FieldB As Object) As Boolean
         Dim returnValue As Boolean = False
         'if both nothing return true
         If FieldA Is Nothing And FieldB Is Nothing Then
            returnValue = True
         Else
            'one is nothing and the other not - not equal
            If (FieldA Is Nothing And Not FieldB Is Nothing) Or (FieldB Is Nothing And Not FieldA Is Nothing) Then
               returnValue = False
            Else
               If FieldA = FieldB Then
                  returnValue = True
               Else
                  returnValue = False
               End If
            End If
         End If
         CompareFields = returnValue
      End Function
#End Region

      Public Sub New()
         Me.DateFirstEnrolled = Now()
      End Sub
   End Class
#Region "StudentDC"
   Friend Class StudentDC
      Function insert(ByRef Student As Student, ByRef connectionString As String) As Integer
         Dim con As SqlConnection = New SqlConnection
         Dim iReturnStudentId As Integer
			Dim iIsSubscribed As Integer
			Dim errorMessage As String = ""
         If Student.isSubscribed = True Then
            iIsSubscribed = 1
         End If
         con.ConnectionString = connectionString
         Dim cmd As SqlCommand = New SqlCommand("USP_Student_Insert1", con)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.Parameters.Add("ReturnValue", SqlDbType.Int)
         cmd.Parameters("ReturnValue").Direction = ParameterDirection.ReturnValue
         cmd.Parameters.Add("@StudentId", SqlDbType.Int)
         cmd.Parameters("@StudentId").Direction = ParameterDirection.Output

         'cmd.Parameters.AddWithValue("@StudentId", Student.StudentId)
         'cmd.Parameters.AddWithValue("@StudentNumber", student.sStudentNumber)
         cmd.Parameters.AddWithValue("@StudentFirstName", Student.FirstName)
         cmd.Parameters.AddWithValue("@StudentLastName", Student.LastName)
         cmd.Parameters.AddWithValue("@StudentNumber", "")
         cmd.Parameters.AddWithValue("@StudentMiddleName", Student.MiddleName)
         cmd.Parameters.AddWithValue("@Address", Student.Address)
         cmd.Parameters.AddWithValue("@City", Student.City)
         cmd.Parameters.AddWithValue("@State", Student.State)
         cmd.Parameters.AddWithValue("@ZipCode", Student.ZipCode)
         cmd.Parameters.AddWithValue("@WorkDayPhone", Student.WorkDayPhone)
         cmd.Parameters.AddWithValue("@HomeEveningPhone", Student.HomeEveningPhone)
         cmd.Parameters.AddWithValue("@SSN", Student.ssn)
         cmd.Parameters.AddWithValue("@email", Student.Email)
         cmd.Parameters.AddWithValue("@Gender", Student.Gender)
         cmd.Parameters.AddWithValue("@DateFirstEnrolled", Student.DateFirstEnrolled)
         cmd.Parameters.AddWithValue("@FriendOfDCE", System.DBNull.Value)
         cmd.Parameters.AddWithValue("@FriendOfDCEDate", System.DBNull.Value)
         cmd.Parameters.AddWithValue("@StudentTypeId", Student.StudentTypeId)
         cmd.Parameters.AddWithValue("@Password", Student.Password)
         cmd.Parameters.AddWithValue("@Subscribed", iIsSubscribed)
         cmd.Parameters.AddWithValue("@NMID", Student.StudentManagerId)
         cmd.Parameters.AddWithValue("@CellPhone", Student.CellPhone)
         Try
            con.Open()
            cmd.ExecuteNonQuery()
            If cmd.Parameters.Count > 0 Then
               iReturnStudentId = cmd.Parameters(1).Value
            End If
         Catch ex As Exception
				errorMessage = ex.Message
         End Try

         Return iReturnStudentId
      End Function
      Function IsStudentExists(ByVal StudentLastName As String, ByVal StudentFirstName As String, ByVal Address As String, ByVal City As String, ByVal State As String, ByVal Zipcode As String, ByVal Email As String, ByVal WorkDayPhone As String, ByVal HomeEveningPhone As String, ByVal CellPhone As String, ByVal ConnectionString As String) As Integer
         Dim con As SqlConnection = New SqlConnection
         Dim iStudentId As Integer
         con.ConnectionString = ConnectionString
         Dim cmd As SqlCommand = New SqlCommand("USP_Student_IsStudentExists", con)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.Parameters.Add("ReturnValue", SqlDbType.Int)
         cmd.Parameters("ReturnValue").Direction = ParameterDirection.ReturnValue
         cmd.Parameters.AddWithValue("@StudentLastName", StudentLastName)
         cmd.Parameters.AddWithValue("@StudentFirstName", StudentFirstName)
         cmd.Parameters.AddWithValue("@Address", Address)
         cmd.Parameters.AddWithValue("@City", City)
         cmd.Parameters.AddWithValue("@State", State)
         cmd.Parameters.AddWithValue("@Zipcode", Zipcode)
         cmd.Parameters.AddWithValue("@Email", Email)
         cmd.Parameters.AddWithValue("@WorkDayPhone", WorkDayPhone)
         cmd.Parameters.AddWithValue("@HomeEveningPhone", HomeEveningPhone)
         cmd.Parameters.AddWithValue("@CellPhone", CellPhone)
         Try
            con.Open()
            cmd.ExecuteNonQuery()
            If cmd.Parameters.Count > 0 Then
               iStudentId = cmd.Parameters(0).Value
            End If
         Catch ex As Exception

         End Try

         Return iStudentId
      End Function
      Function Update(ByRef newerStudent As Student, ByVal StudentId As Integer, ByVal connectionString As String) As Integer
         Dim con As SqlConnection = New SqlConnection
         Dim iRowsAffected As Integer
         con.ConnectionString = connectionString
         Dim cmd As SqlCommand = New SqlCommand("USP_Student_Update2", con)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.Parameters.Add("ReturnValue", SqlDbType.Int)
         cmd.Parameters("ReturnValue").Direction = ParameterDirection.ReturnValue
         cmd.Parameters.AddWithValue("@StudentId", newerStudent.StudentId)
         cmd.Parameters.AddWithValue("@StudentLastName", newerStudent.LastName)
         cmd.Parameters.AddWithValue("@StudentFirstName", newerStudent.FirstName)
         cmd.Parameters.AddWithValue("@Address", newerStudent.address)
         cmd.Parameters.AddWithValue("@City", newerStudent.city)
         cmd.Parameters.AddWithValue("@State", newerStudent.state)
         cmd.Parameters.AddWithValue("@Zipcode", newerStudent.zipcode)
         cmd.Parameters.AddWithValue("@Email", newerStudent.email)
         cmd.Parameters.AddWithValue("@WorkDayPhone", newerStudent.WorkDayPhone)
         cmd.Parameters.AddWithValue("@HomeEveningPhone", newerStudent.HomeEveningPhone)
         cmd.Parameters.AddWithValue("@CellPhone", newerStudent.CellPhone)
         Try
            con.Open()
            cmd.ExecuteNonQuery()
            If cmd.Parameters.Count > 0 Then
               iRowsAffected = cmd.Parameters(0).Value
            End If
         Catch ex As Exception

         End Try

         Return iRowsAffected
      End Function
      Function Load(ByVal studentId As Integer, ByVal ConnectionString As String) As DataTable
         Dim con As SqlConnection = New SqlConnection
         Dim da As SqlDataAdapter
         Dim ds As New DataSet
         Dim dt As New DataTable
         Dim iRecords As Integer
         con.ConnectionString = ConnectionString
         da = New SqlDataAdapter("USP_Student_Load1", con)
         da.SelectCommand.CommandType = CommandType.StoredProcedure
         da.SelectCommand.Parameters.Add("ReturnValue", SqlDbType.Int)
         da.SelectCommand.Parameters("ReturnValue").Direction = ParameterDirection.ReturnValue
         da.SelectCommand.Parameters.AddWithValue("@StudentId", studentId)
         Try
            con.Open()
            iRecords = da.Fill(ds, "Student")
            dt = ds.Tables(0)
            da.Dispose()
            con.Close()
            con.Dispose()
         Catch ex As Exception

         End Try
         Load = dt
      End Function
   End Class
#End Region
#Region "RegistrationDC"

   Friend Class RegistrationDC
      Function insert(ByRef registration As Registration, ByRef connectionString As String) As Integer
         Dim con As SqlConnection = New SqlConnection
         Dim iNewRegistrationId As Integer
         con.ConnectionString = connectionString
         Dim cmd As SqlCommand = New SqlCommand("USP_CourseRegistration_Insert1", con)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.Parameters.Add("ReturnValue", SqlDbType.Int)
         cmd.Parameters("ReturnValue").Direction = ParameterDirection.ReturnValue
         cmd.Parameters.Add("@CORegistrationId", SqlDbType.Int)
         cmd.Parameters("@CORegistrationId").Direction = ParameterDirection.Output
         cmd.Parameters.AddWithValue("@StudentID", registration.StudentId)
         cmd.Parameters.AddWithValue("@SubmittedOn", registration.SubmittedOn)
         cmd.Parameters.AddWithValue("@RegistrationFee", registration.RegistrationFees)
         cmd.Parameters.AddWithValue("@StatusId", registration.StatusId)
         If registration.ProcessedOn = Nothing Then
            cmd.Parameters.AddWithValue("@ProcessedOn", System.DBNull.Value)
         Else
            cmd.Parameters.AddWithValue("@ProcessedOn", registration.ProcessedOn)
         End If
         cmd.Parameters.AddWithValue("@ProcessedBy", registration.ProcessedBy)
			cmd.Parameters.AddWithValue("@ProcessorsRemarks", registration.ProcessorRemarks)
         Try
            con.Open()
            cmd.ExecuteNonQuery()
            If cmd.Parameters.Count > 0 Then
               iNewRegistrationId = cmd.Parameters(1).Value
            End If
         Catch ex As Exception

         End Try

         Return iNewRegistrationId
      End Function
      Function insertCourseTaken(ByRef course As CourseEventData, ByRef connectionString As String) As Integer
         Dim con As SqlConnection = New SqlConnection
         Dim iCourseTakenId As Integer
         con.ConnectionString = connectionString
         Dim cmd As SqlCommand = New SqlCommand("USP_CourseTaken_Insert1", con)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.Parameters.Add("ReturnValue", SqlDbType.Int)
         cmd.Parameters("ReturnValue").Direction = ParameterDirection.ReturnValue
         cmd.Parameters.Add("@CourseTakenID", SqlDbType.Int)
         cmd.Parameters("@CourseTakenID").Direction = ParameterDirection.Output
         cmd.Parameters.AddWithValue("@CourseID", course.CourseId)
         cmd.Parameters.AddWithValue("@CourseRegistrationID", course.RegistrationId)
         cmd.Parameters.AddWithValue("@CourseNumber", course.CourseNumber)
         cmd.Parameters.AddWithValue("@RegistrationFee", course.RegistrationFee)
         cmd.Parameters.AddWithValue("@Tuition", course.Tuition)
         cmd.Parameters.AddWithValue("@NumberOfSeats", course.NumberOfSeats)
         Try
            con.Open()
            cmd.ExecuteNonQuery()
            If cmd.Parameters.Count > 0 Then
               iCourseTakenId = cmd.Parameters(1).Value
            End If
         Catch ex As Exception

         End Try

         Return iCourseTakenId
      End Function
      Function GetRegistrationSetById(ByVal Registrationid As Integer, ByVal ConnectionString As String) As DataSet
         Dim con As SqlConnection = New SqlConnection
         Dim da As SqlDataAdapter
         Dim ds As New DataSet
         Dim dt As New DataTable
         Dim iRecords As Integer
         con.ConnectionString = ConnectionString
         da = New SqlDataAdapter("USP_CourseRegistration_LoadSet", con)
         da.SelectCommand.CommandType = CommandType.StoredProcedure
         da.SelectCommand.Parameters.Add("ReturnValue", SqlDbType.Int)
         da.SelectCommand.Parameters("ReturnValue").Direction = ParameterDirection.ReturnValue
         da.SelectCommand.Parameters.AddWithValue("@CORegistrationId", Registrationid)
         Try
            con.Open()
            iRecords = da.Fill(ds, "Registration")
            da.Dispose()
            con.Close()
            con.Dispose()
         Catch ex As Exception

         End Try
         GetRegistrationSetById = ds
      End Function
      Function GetRegistrationById(ByVal Registrationid As Integer, ByVal ConnectionString As String) As DataTable
         Dim con As SqlConnection = New SqlConnection
         Dim da As SqlDataAdapter
         Dim ds As New DataSet
         Dim dt As New DataTable
         Dim iRecords As Integer
         con.ConnectionString = ConnectionString
         da = New SqlDataAdapter("USP_CourseRegistration_Load", con)
         da.SelectCommand.CommandType = CommandType.StoredProcedure
         da.SelectCommand.Parameters.Add("ReturnValue", SqlDbType.Int)
         da.SelectCommand.Parameters("ReturnValue").Direction = ParameterDirection.ReturnValue
         da.SelectCommand.Parameters.AddWithValue("@CORegistrationId", Registrationid)
         Try
            con.Open()
            iRecords = da.Fill(ds, "Registration")
            dt = ds.Tables(0)
            da.Dispose()
            con.Close()
            con.Dispose()
         Catch ex As Exception

         End Try
         GetRegistrationById = dt
      End Function
      Function GetCoursesTakenByRegistrationid(ByVal Registrationid As Integer, ByVal ConnectionString As String) As DataTable
         Dim con As SqlConnection = New SqlConnection
         Dim da As SqlDataAdapter
         Dim ds As New DataSet
         Dim dt As New DataTable
         Dim iRecords As Integer
         con.ConnectionString = ConnectionString
         da = New SqlDataAdapter("USP_CourseTaken_LoadListByCourseRegistrationID", con)
         da.SelectCommand.CommandType = CommandType.StoredProcedure
         da.SelectCommand.Parameters.Add("ReturnValue", SqlDbType.Int)
         da.SelectCommand.Parameters("ReturnValue").Direction = ParameterDirection.ReturnValue
         da.SelectCommand.Parameters.AddWithValue("@CORegistrationId", Registrationid)
         Try
            con.Open()
            iRecords = da.Fill(ds, "CourseTaken")
            da.Dispose()
            dt = ds.Tables("CourseTaken")
         Catch ex As Exception

         End Try
         GetCoursesTakenByRegistrationid = dt
      End Function
      Function LoadCourseRegistrationStatus(ByVal connectionString) As DataTable
         Dim con As SqlConnection = New SqlConnection
         Dim da As SqlDataAdapter
         Dim ds As New DataSet
         Dim dt As New DataTable
         Dim iRecords As Integer
         con.ConnectionString = connectionString
         da = New SqlDataAdapter("USP_CourseRegistrationStatus_LoadList", con)
         da.SelectCommand.CommandType = CommandType.StoredProcedure
         da.SelectCommand.Parameters.Add("ReturnValue", SqlDbType.Int)
         da.SelectCommand.Parameters("ReturnValue").Direction = ParameterDirection.ReturnValue
         Try
            con.Open()
            iRecords = da.Fill(ds, "CourseRegistrationStatus")
            dt = ds.Tables("CourseRegistrationStatus")
            da.Dispose()
            con.Close()
            con.Dispose()
         Catch ex As Exception

         End Try
         LoadCourseRegistrationStatus = dt
      End Function
      Function ChangeStatus(ByVal CourseRegistrationId As Integer, ByVal newstatusId As Integer, ByVal connectionString As String) As Boolean
         Dim con As SqlConnection = New SqlConnection
         Dim result As Boolean = False
         con.ConnectionString = connectionString
         Dim cmd As SqlCommand = New SqlCommand("USP_CourseRegistration_ChangeStatus", con)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.Parameters.Add("ReturnValue", SqlDbType.Int)
         cmd.Parameters("ReturnValue").Direction = ParameterDirection.ReturnValue
         cmd.Parameters.AddWithValue("@CORegistrationID", CourseRegistrationId)
         cmd.Parameters.AddWithValue("@StatusId", newstatusId)
         Try
            con.Open()
            cmd.ExecuteNonQuery()
            If cmd.Parameters.Count > 0 Then
               If cmd.Parameters(1).Value > 1 Then
                  result = True
               End If
            End If
         Catch ex As Exception
            Throw New Exception("Chaging status failed" & ex.Message)
         End Try

         Return result
      End Function
      ''' <summary>
      ''' This overload with PaymentReference #
      ''' Needed to keep the other one not changed in case it still used somewhere
      ''' </summary>
      ''' <param name="CourseRegistrationId"></param>
      ''' <param name="newstatusId"></param>
      ''' <param name="PaymentReferenceNumber"></param>
      ''' <param name="connectionString"></param>
      ''' <returns></returns>
      ''' <remarks></remarks>
      Function ChangeStatus(ByVal CourseRegistrationId As Integer, ByVal newstatusId As Integer, ByVal PaymentReferenceNumber As String, ByVal connectionString As String) As Boolean
         Dim con As SqlConnection = New SqlConnection
         Dim result As Boolean = False
         con.ConnectionString = connectionString
         Dim cmd As SqlCommand = New SqlCommand("USP_CourseRegistration_ChangeStatusAndReference", con)
         cmd.CommandType = CommandType.StoredProcedure
         cmd.Parameters.Add("ReturnValue", SqlDbType.Int)
         cmd.Parameters("ReturnValue").Direction = ParameterDirection.ReturnValue
         cmd.Parameters.AddWithValue("@CORegistrationID", CourseRegistrationId)
         cmd.Parameters.AddWithValue("@StatusId", newstatusId)
         cmd.Parameters.AddWithValue("@paymentReferenceNumber", PaymentReferenceNumber)
         Try
            con.Open()
            cmd.ExecuteNonQuery()
            If cmd.Parameters.Count > 0 Then
               If cmd.Parameters(1).Value > 1 Then
                  result = True
               End If
            End If
         Catch ex As Exception
            Throw New Exception("Chaging status failed" & ex.Message)
         End Try

         Return result
      End Function
   End Class
#End Region
#Region "RegistrationStatus"

   Public Class RegistrationStatus
      Private _StatusID
      Private _Description
      Private _SystemName
      Public Property StatusId() As Integer
         Get
            StatusId = _StatusID
         End Get
         Set(ByVal Value As Integer)
            _StatusID = Value
         End Set
      End Property
      Public Property Description() As String
         Get
            Description = _Description
         End Get
         Set(ByVal Value As String)
            _Description = Value
         End Set
      End Property
      Public Property SystemName() As String
         Get
            SystemName = _SystemName
         End Get
         Set(ByVal Value As String)
            _SystemName = Value
         End Set
      End Property
   End Class
#End Region
End Namespace