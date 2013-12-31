Imports System.Data.OleDb
Imports Word = Microsoft.Office.Interop.Word
'( - - - - - - - - - - - )'
'( Harry Mumford -Turner )'
'( Student Tracking Sys  )'
'(  Last Edited : 29/03  )'
'( - - - - - - - - - - - )'
PublicClassForm1
Dim SSStudentID, SSCourseID, SSTMAID AsInteger
Dim DTStudents, DTCourses, DTStudentCourses, DTContacts, DTTMAS, DTStudentTMAS AsDataTable
Dim DTofTMA AsDataTable = NewDataTable
Dim DTofContact AsDataTable = NewDataTable
Dim DTAllStud AsDataTable = NewDataTable
Dim DA AsOleDbDataAdapter
Dim DS AsDataSet
Dim oleCom AsOleDbCommand
Dim oleCon AsOleDbConnection
DimTables(6) AsString

PrivateSub MainForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) HandlesMyBase.Load
EndSub

'Database load
PrivateSub Btn_CreateDB_Click(sender As System.Object, e As System.EventArgs) Handles Btn_CreateDB.Click
'Check to see where the Checkdatabase function has been called from
Dim CreateDB AsBoolean = True
CheckDatabaseExists(CreateDB)
EndSub

FunctionCheckDatabaseExists(ByRef CreateDB AsBoolean)
'Gives the DBPath name based on user entered values
Dim DatabasePathName AsString = TBox_DriveLetter.Text &":\"& TBox_DatabaseName.Text &".accdb"
'Applies the connection string with the pathname
Dim ConnectionString AsString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="& DatabasePathName &";"
'If the directory exists and the user has clicked the create DB button then a message box is displayed informing the user and 'new' is added to the filename
IfDir(DatabasePathName) <>""And CreateDB = TrueThen
MsgBox("Database allready exists.")
            TBox_DatabaseName.Text = TBox_DatabaseName.Text &"new"
ElseIfDir(DatabasePathName) <>""And CreateDB = FalseThen
'If the database does exist and the user clicked on the Display Students button then simply get the connection string
Else
'Or if the database really does not exist then create it using the above strings
CreateDatabase(ConnectionString)
EndIf
'Returns the connection string to the caller of this function
        CheckDatabaseExists = ConnectionString
EndFunction



PrivateSubCreateDatabase(ByVal ConnectionString AsString)
'Create New ADOX Database Object, commands which enable the creation of the database
Dim adoxDB AsNew ADOX.Catalog
Dim adodbCon AsNew ADODB.Connection
Dim adodbCom AsNew ADODB.Command

'Validation for the creation of tables 
 'A counter which increments when a table is created
Dim NoTables AsInteger = 0        
	'Creates the database in the entered pathname
adoxDB.Create(ConnectionString)

'The PathName is set to the connection object so we connected to the database
        adodbCon.ConnectionString = ConnectionString

'Opens the connection between the database and the Data set
adodbCon.Open()
        adoxDB.ActiveConnection = adodbCon

'Creates each table for the database
Dim SQLString AsString

' ------------------------
'Create tblStudent
' ------------------------

        SQLString = "CREATE TABLE tblStudent " _
&"(" _
&"StudID LONG, " _
&"CONSTRAINT pk_StudID PRIMARY KEY (StudID), " _
&"FName TEXT(60), " _
&"SName TEXT(60), " _
&"IDNumber TEXT(6), " _
&"StudComment TEXT(200), " _
&"PhoneNo TEXT(15), " _
&"EmailAddress TEXT(40), " _
&"DateCreated DATE" _
&")"

'The SQL is inserted into the command object which is then executed
        adodbCom.CommandText = SQLString
        adodbCom.ActiveConnection = adodbCon
adodbCom.Execute()
'The counter incrementing
        NoTables += 1

' ------------------------
' Create tblCourse
' ------------------------

        SQLString = "CREATE TABLE tblCourse " _
&"(" _
&"CourseID LONG, " _
&"CONSTRAINT pk_CourseID PRIMARY KEY (CourseID), " _
&"CourseName TEXT(20)" _
&")"
        adodbCom.CommandText = SQLString
        adodbCom.ActiveConnection = adodbCon
adodbCom.Execute()
        NoTables += 1


' ------------------------
' Create tblStudentCourse
' ------------------------
'Note defining a compound primary key below
        SQLString = "CREATE TABLE tblStudentCourse " _
&"(" _
&"StudID LONG, " _
&"CourseID LONG, " _
&"CONSTRAINT fk_Coursetbl FOREIGN KEY (CourseID) REFERENCES tblCourse(CourseID), " _
&"CONSTRAINT fk_Studenttbl FOREIGN KEY (StudID) REFERENCES tblStudent(StudID), " _
&"PRIMARY KEY (StudID, CourseID), " _
&"StartDate DATE, " _
&"EndDate DATE, " _
&"Dormant TEXT(6)" _
&")"
        adodbCom.CommandText = SQLString
        adodbCom.ActiveConnection = adodbCon
adodbCom.Execute()
        NoTables += 1

' ------------------------
' Create tblTMA
' ------------------------

        SQLString = "CREATE TABLE tblTMA " _
&"(" _
&"TMAID LONG, " _
&"CONSTRAINT pk_TMAID PRIMARY KEY (TMAID), " _
&"CourseID LONG, " _
&"CONSTRAINT fk_TMACoursetbl FOREIGN KEY (CourseID) REFERENCES tblCourse(CourseID), " _
&"TMALetter TEXT(1)" _
&")"
        adodbCom.CommandText = SQLString
        adodbCom.ActiveConnection = adodbCon
adodbCom.Execute()
        NoTables += 1

' ------------------------
' Create tblStudent-TMA
' ------------------------
        SQLString = "CREATE TABLE tblStudentTMA " _
&"(" _
&"StudID LONG, " _
&"CourseID LONG, " _
&"TMAID LONG, " _
&"CONSTRAINT fk_TMAtbl FOREIGN KEY (TMAID) REFERENCES tblTMA(TMAID), " _
&"CONSTRAINT fk_StudentCourseStudTMAtbl FOREIGN KEY (StudID, CourseID) REFERENCES tblStudentCourse(StudID, CourseID), " _
&"Grade INTEGER, " _
&"DateIn DATE, " _
&"DateOut DATE, " _
&"TMAComment TEXT(200)" _
&")"
        adodbCom.CommandText = SQLString
        adodbCom.ActiveConnection = adodbCon
adodbCom.Execute()
        NoTables += 1




' ------------------------
' Create tblContact Table
' ------------------------
        SQLString = "CREATE TABLE tblContact " _
&"(" _
&"StudID LONG, " _
&"ContactID LONG, " _
&"DateOfContact DATE, " _
&"Type TEXT(1), " _
&"Duration INTEGER, " _
&"ContactComment TEXT(200), " _
&"DateForNextContact DATE, " _
&"CONSTRAINT fk_StudentContacttbl FOREIGN KEY (StudID) REFERENCES tblStudent(StudID)" _
&")"
        adodbCom.CommandText = SQLString
        adodbCom.ActiveConnection = adodbCon
adodbCom.Execute()
        NoTables += 1
'Dialog box to notify the user of completion
MessageBox.Show("Database with "& NoTables &" table(s) has been successfully created!")
EndSub

'Display Students
PrivateSub Btn_DisplayStudents_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_DisplayStudents.Click
'Enables the searching for dormant students
Dim ShowDormantOnly AsBoolean = False
Dim CreateDB AsBoolean = False
If CBox_ShowDormant.Checked = TrueThen
            ShowDormantOnly = True
ElseIf CBox_ShowDormant.Checked = FalseThen
            ShowDormantOnly = False
EndIf
'Subroutine to populate the dataset with datatables extracted from the database
PopulateDSandDT(CheckDatabaseExists(CreateDB)) ' "CheckDatabaseExists" is a function which does just that and creates the DB if it doesn't
'the function returns the DB connection string, which is used when populating the dataset
DisplayStudents(ShowDormantOnly)
EndSub


PrivateSubPopulateDSandDT(ByVal ConStr)
'Form connection string and open the database connection
oleCon = NewOleDbConnection(ConStr)
oleCon.Open()

'Define DA and DS
        DA = NewOleDbDataAdapter
        DS = NewDataSet
oleCom = NewOleDbCommand
        oleCom.Connection = oleCon

Tables(0) = "tblStudent"
Tables(1) = "tblCourse"
Tables(2) = "tblStudentCourse"
Tables(3) = "tblTMA"
Tables(4) = "tblStudentTMA"
Tables(5) = "tblContact"

'Fill a dataset with multiple Data tables
        DA.SelectCommand = oleCom
For i = 0 ToTables.GetUpperBound(0) - 1
            DA.SelectCommand.CommandText = "SELECT * FROM "&Tables(i)
DA.Fill(DS, Tables(i))
Next

'Close the connection
oleCon.Close()

'Map the 'default' names to actual names
DA.TableMappings.Add("Table", "tblStudent")
DA.TableMappings.Add("Table1", "tblCourse")
DA.TableMappings.Add("Table2", "tblStudentCourse")
DA.TableMappings.Add("Table3", "tblTMA")
DA.TableMappings.Add("Table4", "tblStudentTMA")
DA.TableMappings.Add("Table5", "tblContact")

'Create explicit table names to clarify code
        DTStudents = DS.Tables(0)
        DTCourses = DS.Tables(1)
        DTStudentCourses = DS.Tables(2)
        DTTMAS = DS.Tables(3)
        DTStudentTMAS = DS.Tables(4)
        DTContacts = DS.Tables(5)
'Create database relationships for the dataset
'Student - StudentCourse
DS.Relations.Add(NewDataRelation("StudStudCorRelation", DTStudents.Columns("StudID"), DTStudentCourses.Columns("StudID"), True))
'Course - StudentCourse
DS.Relations.Add(NewDataRelation("CourStudCorRelation", DTCourses.Columns("CourseID"), DTStudentCourses.Columns("CourseID"), True))
'Course - TMA
DS.Relations.Add(NewDataRelation("CourTMARelation", DTCourses.Columns("CourseID"), DTTMAS.Columns("CourseID"), True))
'TMA - StudentTMA
DS.Relations.Add(NewDataRelation("TMAStudTMA", DTTMAS.Columns("TMAID"), DTStudentTMAS.Columns("TMAID"), True))
'Student - Contact
DS.Relations.Add(NewDataRelation("StudContactRelation", DTStudents.Columns("StudID"), DTContacts.Columns("StudID"), True))

'Compound Primary Key For 
' StudentCourse - StudentTMA
'Create an array of datacolumns where the 2 primary keys are held
DimSCcpk(1), SCSTcpk(1) AsDataColumn
SCcpk(0) = DTStudentCourses.Columns("StudID")
SCcpk(1) = DTStudentCourses.Columns("CourseID")
SCSTcpk(0) = DTStudentTMAS.Columns("StudID")
SCSTcpk(1) = DTStudentTMAS.Columns("CourseID")

'Enforces the relations between these 2 keys
DS.Relations.Add(NewDataRelation("CPKStudentTMA", SCcpk, SCSTcpk, True))
        DS.EnforceConstraints = True

'Gets the number of rows in the students table and displays it for the user
        TBox_allstudNoRows.Text = DTStudents.Rows.Count

'Enables the user commands for the ability to print
        GBox_SS.Enabled = True
EndSub

SubDisplayStudents(ShowDormantOnly AsBoolean)
        DTAllStud = CreateTempDTStudents()

Dim StudCourseRow AsDataRow
Dim CourseRow AsDataRow
Dim StudRow AsDataRow

'Goes through each row in DTStudents to see if the StudID matches in DTStudentCourses
If ShowDormantOnly = TrueThen
ForEach StudRow InDTStudents.Rows()
ForEach StudCourseRow InDTStudentCourses.Select("StudID = "& StudRow.Item("StudID") &" AND Dormant = 'True'")
'Goes through each row in DTCourses to see if the courseID of DTStudentCourses match
ForEach CourseRow InDTCourses.Select("CourseID = "& StudCourseRow.Item("CourseID"))
'If they do then add a new row to the 'AllStud' datatable based on the other records in the matching course row
Dim Addingrow AsDataRow = DTAllStud.NewRow()
Dim tempDateCreated AsDate = StudRow.Item("DateCreated")
With Addingrow
                            .Item(0) = StudRow.Item("StudID")
                            .Item(1) = StudRow.Item("FName")
                            .Item(2) = StudRow.Item("SName")
                            .Item(3) = tempDateCreated.ToString("d")
                            .Item(4) = StudRow.Item("StudComment")
                            .Item(5) = CourseRow.Item("CourseID")
                            .Item(6) = CourseRow.Item("CourseName")
                            .Item(7) = StudRow.Item("IDNumber")
EndWith
DTAllStud.Rows.Add(Addingrow)
Next
Next
Next
Else
ForEach StudRow InDTStudents.Rows()
ForEach StudCourseRow InDTStudentCourses.Select("StudID = "& StudRow.Item("StudID") &" AND Dormant = 'False'")
'Goes through each row in DTCourses to see if the courseID of DTStudentCourses match
ForEach CourseRow InDTCourses.Select("CourseID = "& StudCourseRow.Item("CourseID"))
'If they do then add a new row to the 'AllStud' datatable based on the other records in the matching course row
Dim Addingrow AsDataRow = DTAllStud.NewRow()
Dim tempDateCreated AsDate = StudRow.Item("DateCreated")
With Addingrow
     .Item(0) = StudRow.Item("StudID")
     .Item(1) = StudRow.Item("FName")
     .Item(2) = StudRow.Item("SName")
     .Item(3) = tempDateCreated.ToString("d")
     .Item(4) = StudRow.Item("StudComment")
     .Item(5) = CourseRow.Item("CourseID")
     .Item(6) = CourseRow.Item("CourseName")
     .Item(7) = StudRow.Item("IDNumber")
EndWith
DTAllStud.Rows.Add(Addingrow)
Next
Next
Next
EndIf




'Creates a new view on the data which uses the temporary data table to be displayed in the Data grid view
Dim view AsDataView = NewDataView
        view.Table = DTAllStud
        DGrid_Students.DataSource = view

'Renames the headers - making them look more pleasurable to the user
With DGrid_Students
            .Columns.Item("FName").HeaderText() = "First Name"
            .Columns.Item("SName").HeaderText() = "Second Name"
            .Columns.Item("DateCreated").HeaderText() = "Date Created"
            .Columns.Item("StudComment").Visible = False
            .Columns.Item("StudID").Visible = False
            .Columns.Item("CourseID").Visible = False
            .Columns.Item("StudentNumber").Visible = False
EndWith
EndSub

FunctionCreateTempDTStudents()
'Declare a new datatable, name it and give it new datacolumns
Dim tempDT AsDataTable = NewDataTable
        tempDT.TableName = "AllStudents"

'StudentID Column
Dim StudentIDColumn AsDataColumn = NewDataColumn
        StudentIDColumn.DataType = System.Type.GetType("System.Single")
        StudentIDColumn.ColumnName = "StudID"

' FirstName Column
Dim FirstNameColumn AsDataColumn = NewDataColumn
        FirstNameColumn.DataType = System.Type.GetType("System.String")
        FirstNameColumn.ColumnName = "FName"

'SurName Column
Dim SurNameColumn AsDataColumn = NewDataColumn
        SurNameColumn.DataType = System.Type.GetType("System.String")
        SurNameColumn.ColumnName = "SName"

'DateCreatedC olumn
Dim DateCreatedColumn AsDataColumn = NewDataColumn
        DateCreatedColumn.DataType = System.Type.GetType("System.String")
        DateCreatedColumn.ColumnName = "DateCreated"

'StudComment Column
Dim StudCommentColumn AsDataColumn = NewDataColumn
        StudCommentColumn.DataType = System.Type.GetType("System.String")
        StudCommentColumn.ColumnName = "StudComment"

'CourseID Column
Dim CourseIDColumn AsDataColumn = NewDataColumn
        CourseIDColumn.DataType = System.Type.GetType("System.Single")
        CourseIDColumn.ColumnName = "CourseID"

'CourseName Column
Dim CourseNameColumn AsDataColumn = NewDataColumn
        CourseNameColumn.DataType = System.Type.GetType("System.String")
        CourseNameColumn.ColumnName = "CourseName"

'Student Number Column
Dim StudentNumberColumn AsDataColumn = NewDataColumn
        StudentNumberColumn.DataType = System.Type.GetType("System.String")
        StudentNumberColumn.ColumnName = "StudentNumber"
'Add the columns to the newly created Datatable
With tempDT.Columns
            .Add(StudentIDColumn)
            .Add(FirstNameColumn)
            .Add(SurNameColumn)
            .Add(DateCreatedColumn)
            .Add(StudCommentColumn)
            .Add(CourseIDColumn)
            .Add(CourseNameColumn)
            .Add(StudentNumberColumn)
EndWith

Return tempDT
EndFunction


'Select a student
PrivateSub DataGridView_Students_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGrid_Students.CellClick
        Btn_MakeDormant.Enabled = True
        Btn_MakeUNDormant.Enabled = True
'If the user selects a column to sort the data by that column then an error wont occur
If e.RowIndex < 0 Then
Else
            Btn_ViewContact.Enabled = True
            Btn_ViewTMA.Enabled = True
'When the user clicks on a cell it selects the row which contains the information below this information is placed into variables which are displayed in textboxes

'Join First and Second Names together for the main box
            TBox_allstudStudName.Text = DGrid_Students.Rows(e.RowIndex).Cells("FName").Value &" "& DGrid_Students.Rows(e.RowIndex).Cells("SName").Value

'Student Comment
            TBox_StudComment.Text = DGrid_Students.Rows(e.RowIndex).Cells("StudComment").Value

'CourseName is placed in a textbox as well as in the data grid for the user to find it more easily
            TBox_allstudCourse.Text = DGrid_Students.Rows(e.RowIndex).Cells("CourseName").Value

'Gets the student ID number
            TBox_allstudSIDNO.Text = DGrid_Students.Rows(e.RowIndex).Cells("StudentNumber").Value

'Count the number of rows in the datagrid
            TBox_allstudNoRows.Text = DGrid_Students.Rows.Count.ToString

'Fetch the Selected Student StudentID and Course ID
            SSStudentID = DGrid_Students.Rows(e.RowIndex).Cells("StudID").Value
            SSCourseID = DGrid_Students.Rows(e.RowIndex).Cells("CourseID").Value
EndIf
EndSub







'TMA details
PrivateSub Btn_ViewTMA_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_ViewTMA.Click
'CourseName
        TBox_tmaCourse.Text = TBox_allstudCourse.Text

'Find Exam Date
Dim tempStartDate AsDateTime
ForEach row InDTStudentCourses.Select("StudID = "& SSStudentID &" AND CourseID = "& SSCourseID)
tempStartDate = row.Item("StartDate")
If TBox_tmaCourse.Text = "Geography 2"Then
                TBox_tmaExamDate.Text = tempStartDate.AddYears(2).Year.ToString("d")
Else
                TBox_tmaExamDate.Text = tempStartDate.AddYears(1).Year.ToString("d")
EndIf
Next

'DTPicker
        DTP_DateIn.Value = Date.Today
        DTP_DateOut.Value = Date.Today.AddDays(7)

'Border
        TBox_BorderTMA.Text = TBox_allstudSIDNO.Text &" -- "& TBox_allstudStudName.Text &" T.M.A. Details"


'Moves the display over to the TMA tab
        TCtrl_Student.SelectTab(1)
'Populates the TMA datagrid with the TMA relating to the selected student
PopulateTMADGrid()
EndSub



SubPopulateTMADGrid()
'Resets the datatable so errors won’t occur
DTofTMA.Reset()

'Creates a new table with the below columns
        DTofTMA.TableName = "SelectedStudent"

Dim TMALetterColumn AsDataColumn = NewDataColumn
        TMALetterColumn.DataType = System.Type.GetType("System.String")
        TMALetterColumn.ColumnName = "TMALetter"

Dim TMAGradeColumn AsDataColumn = NewDataColumn
        TMAGradeColumn.DataType = System.Type.GetType("System.Decimal")
        TMAGradeColumn.ColumnName = "TMAGrade"

Dim DateInColumn AsDataColumn = NewDataColumn
        DateInColumn.DataType = System.Type.GetType("System.String")
        DateInColumn.ColumnName = "DateIn"

Dim DateOutColumn AsDataColumn = NewDataColumn
        DateOutColumn.DataType = System.Type.GetType("System.String")
        DateOutColumn.ColumnName = "DateOut"

Dim TMAIDColumn AsDataColumn = NewDataColumn
        TMAIDColumn.DataType = System.Type.GetType("System.Single")
        TMAIDColumn.ColumnName = "TMAID"

Dim StudIDColumn AsDataColumn = NewDataColumn
        StudIDColumn.DataType = System.Type.GetType("System.Single")
        StudIDColumn.ColumnName = "StudID"

Dim CourseIDColumn AsDataColumn = NewDataColumn
        CourseIDColumn.DataType = System.Type.GetType("System.Single")
        CourseIDColumn.ColumnName = "CourseID"

Dim TMACommentColumn AsDataColumn = NewDataColumn
        TMACommentColumn.DataType = System.Type.GetType("System.String")
        TMACommentColumn.ColumnName = "TMAComment"

With DTofTMA
            .Columns.Add(TMALetterColumn)
            .Columns.Add(TMAGradeColumn)
            .Columns.Add(DateInColumn)
            .Columns.Add(DateOutColumn)
            .Columns.Add(StudIDColumn)
            .Columns.Add(CourseIDColumn)
            .Columns.Add(TMACommentColumn)
            .Columns.Add(TMAIDColumn)
EndWith


'Creates a new view and fill the above columns with data from the pure data tables
Dim view AsDataView = NewDataView
Dim StudTMArow AsDataRow
Dim TMArow AsDataRow

ForEach StudTMArow InDTStudentTMAS.Select("StudID = "& SSStudentID &" AND CourseID = "& SSCourseID)
ForEach TMArow InDTTMAS.Select("TMAID = "& StudTMArow.Item("TMAID") &" AND CourseID = "& SSCourseID)
Dim addingrow AsDataRow
addingrow = DTofTMA.NewRow()
Dim tempDateIn AsDate = StudTMArow.Item("DateIn")
Dim tempDateOut AsDate = StudTMArow.Item("DateOut")
With addingrow
                    .Item(0) = TMArow.Item("TMALetter")
                    .Item(1) = StudTMArow.Item("Grade")
                    .Item(2) = tempDateIn.ToString("d")
                    .Item(3) = tempDateOut.ToString("d")
                    .Item(4) = StudTMArow.Item("StudID")
                    .Item(5) = StudTMArow.Item("CourseID")
                    .Item(6) = StudTMArow.Item("TMAComment")
                    .Item(7) = TMArow.Item("TMAID")
EndWith
DTofTMA.Rows.Add(addingrow)
Next
Next
'Point the newly created datatable to the view which points to the datagrid
        view.Table = DTofTMA
        DGrid_TMA.DataSource = view

'Hide the index columns and others that wont be visible to the user
With DGrid_TMA
            .Columns.Item("StudID").Visible = False
            .Columns.Item("CourseID").Visible = False
            .Columns.Item("TMAID").Visible = False
            .Columns.Item("TMAComment").Visible = False
EndWith
'Get the maximum TMA Letter
Dim RowOfTMALetter AsInteger = DGrid_TMA.Rows.Count - 1
Dim TMALetter AsString
Try
            TMALetter = DGrid_TMA.Rows(RowOfTMALetter).Cells("TMALetter").Value
Catch ex AsException
            TMALetter = "Null"
EndTry
'Clears the Old TMA Letters that were previously in the Combo Box
        CBox_TMALetterAdd.Items.Clear()

'For the adding of new student
'Fills the combo box with the correct TMALetter Names
Dim CourseNameRow AsDataRow
Dim AddTMALetters AsBoolean
ForEach CourseNameRow InDTCourses.Select("CourseID = "& SSCourseID)
SelectCaseCourseNameRow.Item("CourseName")
Case"Mathematics Year 7"To"Mathematics Year 8"
ForEach c In"ABCDEF".ToCharArray()
If c = TMALetter Then
                            AddTMALetters = True
ElseIf TMALetter = "Null"Then
                            CBox_TMALetterAdd.Items.Add(c)
Else
If AddTMALetters = TrueThen
                                CBox_TMALetterAdd.Items.Add(c)
EndIf
EndIf
Next c
Case"Geography 1"To"Geography 2"
ForEach c In"ABCDEFGHIJ".ToCharArray()
If c = TMALetter Then
                            AddTMALetters = True
ElseIf TMALetter = "Null"Then
                            CBox_TMALetterAdd.Items.Add(c)
Else
If AddTMALetters = TrueThen
                                CBox_TMALetterAdd.Items.Add(c)
EndIf
EndIf
Next c
EndSelect
Next
EndSub


'TMA select
PrivateSub DataGridView_TMAStudent_CellClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGrid_TMA.CellClick
'If the user selects a column to sort the data by that column then an error wont occur
If e.RowIndex < 0 Then
Else
'Fill the textboxes with the information about the selected student
            lb_TMALetter.Text = DGrid_TMA.Rows(e.RowIndex).Cells("TMALetter").Value
            Tbox_Grade.Text = DGrid_TMA.Rows(e.RowIndex).Cells("TMAGrade").Value &"%"
            TBox_TMAComment.Text = DGrid_TMA.Rows(e.RowIndex).Cells("TMAComment").Value
            SSTMAID = DGrid_TMA.Rows(e.RowIndex).Cells("TMAID").Value
EndIf
EndSub



'TMA Add
PrivateSub Button_AddnewTMA_Click(sender As System.Object, e As System.EventArgs) Handles Btn_AddTMA.Click
'Get the max TMAID
Dim SQL AsString = "SELECT MAX(TMAID) FROM tblStudentTMA"
Dim MaxTMAID AsInteger
Dim oleComScalar AsNewOleDbCommand(SQL, oleCon)
oleComScalar.Connection.Open()
Try
            MaxTMAID = oleComScalar.ExecuteScalar()
Catch ex AsException
'If there isn’t any values let TMAID = 1
            MaxTMAID = 0
EndTry
        MaxTMAID += 1
        oleCom.Connection = oleCon

'Add to Dataset
Dim TMAtblrow AsDataRow
        TMAtblrow = DS.Tables("tblTMA").NewRow()
TMAtblrow.Item("CourseID") = SSCourseID
TMAtblrow.Item("TMAID") = MaxTMAID
TMAtblrow.Item("TMALetter") = CBox_TMALetterAdd.Text
DTTMAS.Rows.Add(TMAtblrow)

Dim StudTMAtblrow AsDataRow
        StudTMAtblrow = DS.Tables("tblStudentTMA").NewRow()
StudTMAtblrow.Item("StudID") = SSStudentID
StudTMAtblrow.Item("CourseID") = SSCourseID
StudTMAtblrow.Item("TMAID") = MaxTMAID
StudTMAtblrow.Item("Grade") = CInt(TBox_GradeAdd.Text)
StudTMAtblrow.Item("DateIn") = DTP_DateIn.Value.ToString("d")
StudTMAtblrow.Item("DateOut") = DTP_DateOut.Value.ToString("d")
StudTMAtblrow.Item("TMAComment") = TBox_TMAComment.Text
DTStudentTMAS.Rows.Add(StudTMAtblrow)

'Add to Database
        DA.InsertCommand = NewOleDbCommand("INSERT INTO tblTMA (CourseID, TMAID, TMALetter) " _
&"VALUES ("& SSCourseID &", "& MaxTMAID &", '"& CBox_TMALetterAdd.Text &"')", oleCon)
DA.Update(DTTMAS)
        DA.InsertCommand = NewOleDbCommand("INSERT INTO tblStudentTMA (StudID, CourseID, TMAID, Grade, DateIn, DateOut, TMAComment) " _
&"VALUES ("& SSStudentID &", "& SSCourseID &", "& MaxTMAID &", '"&CInt(TBox_GradeAdd.Text) &"', #"& DTP_DateIn.Value.ToString("d") &"#, #"& DTP_DateOut.Value.ToString("d") &"#, '"& TBox_TMAComment.Text &"')", oleCon)
DA.Update(DTStudentTMAS)
DS.AcceptChanges()
oleCon.Close()
'Repopulates the datagrid with the new information
PopulateTMADGrid()

EndSub









'TMA Print
PrivateSub Btn_tmaPrintTMA_Click(sender As System.Object, e As System.EventArgs) Handles Btn_tmaPrintTMA.Click
Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oTable As Word.Table
Dim oPara1, oPara2, oPara3, oPara4 As Word.Paragraph
Dim missing AsObject = System.Reflection.Missing.Value

'Start Word and open the document template.
oWord = CreateObject("Word.Application")
        oWord.Visible = True
oDoc = oWord.Documents.Add

'Insert a paragraph at the beginning of the document.
        oPara1 = oDoc.Content.Paragraphs.Add
        oPara1.Range.Text = "T.M.A. Details for Student: "& TBox_allstudStudName.Text        &" -- "& TBox_allstudSIDNO.Text
        oPara1.Range.Font.Bold = True
        oPara1.Format.SpaceAfter = 12    '12 pt spacing after paragraph.
oPara1.Range.InsertParagraphAfter()

        oPara1.Range.Font.Bold = False
oPara1.Range.InsertFile(TBox_DriveLetter.Text &":\StudContactReport.docx", missing, False, False, False)
Dim MaxRows AsInteger
        MaxRows = DGrid_TMA.Rows.Count

'Insert a MaxRows x 4 table, fill it with data
Dim r AsInteger, c AsInteger
oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, MaxRows + 1, 4)
        oTable.Range.ParagraphFormat.SpaceAfter = 6
        oTable.Borders.Enable = True

oTable.Cell(1, 1).Range.Text = "TMA Letter"
oTable.Cell(1, 2).Range.Text = "TMA Grade"
oTable.Cell(1, 3).Range.Text = "Date In"
oTable.Cell(1, 4).Range.Text = "Date Out"
For r = 2 To MaxRows + 1
For c = 1 To 4
                oTable.Cell(r, c).Range.Text = DGrid_TMA.Rows(r - 2).Cells.Item(c - 1).Value.ToString
Next
Next
oTable.Rows.Item(1).Range.Font.Bold = True


'Add some text after the table.
        oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara2.Range.Text = "Subject: "& TBox_tmaCourse.Text
        oPara2.Format.SpaceAfter = 1
oPara2.Range.InsertParagraphAfter()
oPara2.Range.InsertFile(TBox_DriveLetter.Text &":\StudContactReport.docx", missing, False, False, False)

        oPara3 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara3.Range.Text = "Tutor: Lesley Mumford"
        oPara3.Format.SpaceAfter = 1
oPara3.Range.InsertParagraphAfter()
oPara3.Range.InsertFile(TBox_DriveLetter.Text &":\StudContactReport.docx", missing, False, False, False)
        oPara4 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara4.Range.Text = "Intended Exam Date: "& TBox_tmaExamDate.Text
        oPara4.Format.SpaceAfter = 1
oPara4.Range.InsertParagraphAfter()
oPara4.Range.InsertFile(TBox_DriveLetter.Text &":\StudContactReport.docx", missing, False, False, False)
EndSub


'Contact details
PrivateSub Btn_ViewContact_Click(sender As System.Object, e As System.EventArgs) Handles Btn_ViewContact.Click
'CourseName
        TBox_contCourse.Text = TBox_allstudCourse.Text

'Find Exam Date
Dim tempStartDate AsDateTime
ForEach row InDTStudentCourses.Select("StudID = "& SSStudentID &" AND CourseID = "& SSCourseID)
tempStartDate = row.Item("StartDate")
If TBox_contCourse.Text = "Geography 2"Then
                TBox_contExamDate.Text = tempStartDate.AddYears(2).Year.ToString("d")
Else
                TBox_contExamDate.Text = tempStartDate.AddYears(1).Year.ToString("d")
EndIf
Next

'Fill Phone and email
ForEach row InDTStudents.Select("StudID = "& SSStudentID)
            TBox_PhoneNo.Text = row.Item("PhoneNo")
            TBox_EmailAddress.Text = row.Item("EmailAddress")
Next

'DTPicker
        DTP_FutureContTime.Value = Date.Today.AddDays(7)

'Border
        TBox_BorderContact.Text = TBox_allstudSIDNO.Text &" -- "& TBox_allstudStudName.Text &" Contact Details"

'Moves the display over to the Contact tab
        TCtrl_Student.SelectTab(2)

PopulateContactDGrid()
EndSub
SubPopulateContactDGrid()
'Resets the datatable so errors wont occur
DTofContact.Reset()
        DTofContact.TableName = "SelectedStudent"

'Creates a new table with the below columns
Dim DateOfContactColumn AsDataColumn = NewDataColumn
        DateOfContactColumn.DataType = System.Type.GetType("System.DateTime")
        DateOfContactColumn.ColumnName = "DateOfContact"

Dim TypeColumn AsDataColumn = NewDataColumn
        TypeColumn.DataType = System.Type.GetType("System.String")
        TypeColumn.ColumnName = "Type"

Dim DurationColumn AsDataColumn = NewDataColumn
        DurationColumn.DataType = System.Type.GetType("System.Single")
        DurationColumn.ColumnName = "Duration"
Dim DateForNextContactColumn AsDataColumn = NewDataColumn
        DateForNextContactColumn.DataType = System.Type.GetType("System.DateTime")
        DateForNextContactColumn.ColumnName = "DateForNextContact"

Dim StudIDColumn AsDataColumn = NewDataColumn
        StudIDColumn.DataType = System.Type.GetType("System.Single")
        StudIDColumn.ColumnName = "StudID"

Dim ContactIDColumn AsDataColumn = NewDataColumn
        ContactIDColumn.DataType = System.Type.GetType("System.Single")
        ContactIDColumn.ColumnName = "ContactID"


Dim ContactCommentColumn AsDataColumn = NewDataColumn
        ContactCommentColumn.DataType = System.Type.GetType("System.String")
        ContactCommentColumn.ColumnName = "ContactComment"


With DTofContact
            .Columns.Add(DateOfContactColumn)
            .Columns.Add(TypeColumn)
            .Columns.Add(DurationColumn)
            .Columns.Add(DateForNextContactColumn)
            .Columns.Add(ContactCommentColumn)
            .Columns.Add(StudIDColumn)
            .Columns.Add(ContactIDColumn)
EndWith
'Creates a new view and fill the above columns with data from the pure data tables
Dim view AsDataView = NewDataView
Dim ContactRow AsDataRow

ForEach ContactRow In DTContacts.Rows
Dim addingrow AsDataRow
addingrow = DTofContact.NewRow()
Dim tempDateForNextContact AsDateTime = ContactRow.Item("DateForNextContact")
Dim tempDateOfContact AsDateTime = ContactRow.Item("DateOfContact")
With addingrow
                .Item("DateOfContact") = tempDateOfContact.ToString("d")
                .Item("Type") = ContactRow.Item("Type")
                .Item("Duration") = ContactRow.Item("Duration")
                .Item("DateForNextContact") = tempDateForNextContact.ToString("d")
                .Item("ContactComment") = ContactRow.Item("ContactComment")
                .Item("StudID") = ContactRow.Item("StudID")
                .Item("ContactID") = ContactRow.Item("ContactID")
EndWith
DTofContact.Rows.Add(addingrow)
Next
'Point the newly created datatable to the view which points to the datagrid
        view.Table = DTofContact
        DGrid_Contact.DataSource = view
'Hide the index columns to the user
With DGrid_Contact
            .Columns.Item("StudID").Visible = False
            .Columns.Item("ContactID").Visible = False
            .Columns.Item("ContactComment").Visible = False
EndWith
EndSub






'Contact select
PrivateSub DGrid_Contact_CellClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGrid_Contact.CellClick
'If the user selects a column to sort the data by that column then an error wont occur
If e.RowIndex < 0 Then
Else
            lb_contTypeBig.Text = DGrid_Contact.Rows(e.RowIndex).Cells("Type").Value
            lb_contDurationBig.Text = DGrid_Contact.Rows(e.RowIndex).Cells("Duration").Value &"m"
            TBox_contComment.Text = DGrid_Contact.Rows(e.RowIndex).Cells("ContactComment").Value
EndIf
EndSub

'Contact add
PrivateSub Btn_ContTimeAdd_Click(sender As System.Object, e As System.EventArgs) Handles Btn_ContTimeAdd.Click
'Get the maximum contactID from the contact table
Dim MaxContactID AsInteger
Dim SQL AsString = "SELECT MAX(ContactID) FROM tblContact"
Dim oleComScalar AsNewOleDbCommand(SQL, oleCon)
oleComScalar.Connection.Open()
Try
            MaxContactID = oleComScalar.ExecuteScalar()
Catch ex AsException
            MaxContactID = 0
EndTry
        MaxContactID += 1
        oleCom.Connection = oleCon

'Add to Dataset
Dim Contacttblrow AsDataRow
        Contacttblrow = DTContacts.NewRow()
Dim Type AsString

'See what type is selected
If Rbtn_contEmailAdd.Checked Then
            Type = "E"
Else
            Type = "T"
EndIf
Contacttblrow.Item("StudID") = SSStudentID
Contacttblrow.Item("ContactID") = MaxContactID
Contacttblrow.Item("DateOfContact") = Date.Today.ToString("d")
Contacttblrow.Item("Type") = Type
Contacttblrow.Item("Duration") = CInt(TBox_Duration.Text)
Contacttblrow.Item("ContactComment") = "Enter Comment here"
Contacttblrow.Item("DateForNextContact") = DTP_FutureContTime.Value.ToString("d")
DTContacts.Rows.Add(Contacttblrow)

'Add to Database
        DA.InsertCommand = NewOleDbCommand("INSERT INTO tblContact (StudID, ContactID, DateOfContact, Type, Duration, ContactComment, DateForNextContact) " _
&"VALUES ("& SSStudentID &","& MaxContactID &", #"&Date.Today.ToString("d") &"#, '"& Type &"', "&CInt(TBox_Duration.Text) &", 'Enter Comment here', #"& DTP_FutureContTime.Value.ToString("d") &"#)", oleCon)
DA.Update(DTContacts)
DS.AcceptChanges()
oleCon.Close()
PopulateContactDGrid()
EndSub
'Contact Print
PrivateSub Btn_contPrintContact_Click(sender As System.Object, e As System.EventArgs) Handles Btn_contPrintContact.Click
Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oTable As Word.Table
Dim oPara1, oPara2, oPara3, oPara4, oPara5, oPara6 As Word.Paragraph
Dim missing AsObject = System.Reflection.Missing.Value

'Start Word and open the document template.
oWord = CreateObject("Word.Application")
        oWord.Visible = True
oDoc = oWord.Documents.Add

'Insert a paragraph at the beginning of the document.
        oPara1 = oDoc.Content.Paragraphs.Add
        oPara1.Range.Text = "Contact Details for Student: "& TBox_allstudStudName.Text &" -- "& TBox_allstudSIDNO.Text
        oPara1.Range.Font.Bold = True
        oPara1.Format.SpaceAfter = 12    '12 pt spacing after paragraph.
oPara1.Range.InsertParagraphAfter()

        oPara1.Range.Font.Bold = False
oPara1.Range.InsertFile(TBox_DriveLetter.Text &":\StudContactReport.docx", missing, False, False, False)
Dim MaxRows AsInteger
ForEach row InDTContacts.Select("StudID = "& SSStudentID)
            MaxRows += 1
Next

'Insert a MaxRows x 4 table, fill it with data
Dim r AsInteger, c AsInteger
oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, MaxRows + 1, 4)
        oTable.Range.ParagraphFormat.SpaceAfter = 6
        oTable.Borders.Enable = True

oTable.Cell(1, 1).Range.Text = "Date Of Contact"
oTable.Cell(1, 2).Range.Text = "Type"
oTable.Cell(1, 3).Range.Text = "Duration"
oTable.Cell(1, 4).Range.Text = "Date for Next Contact"
For r = 2 To MaxRows + 1
For c = 1 To 4
                oTable.Cell(r, c).Range.Text = DGrid_Contact.Rows(r - 2).Cells.Item(c - 1).Value.ToString
Next
Next
oTable.Rows.Item(1).Range.Font.Bold = True

'Add some text after the table.
        oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
oPara2.Range.InsertParagraphBefore()
        oPara2.Range.Text = "Subject: "& TBox_contCourse.Text
        oPara2.Format.SpaceAfter = 24
oPara2.Range.InsertParagraphAfter()
oPara2.Range.InsertFile(TBox_DriveLetter.Text &":\StudContactReport.docx", missing, False, False, False)

        oPara3 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
oPara3.Range.InsertParagraphBefore()
        oPara3.Range.Text = "Tutor: Lesley Mumford"
        oPara3.Format.SpaceAfter = 30
oPara3.Range.InsertParagraphAfter()
oPara3.Range.InsertFile(TBox_DriveLetter.Text &":\StudContactReport.docx", missing, False, False, False)

        oPara4 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
oPara4.Range.InsertParagraphBefore()
        oPara4.Range.Text = "Intended Exam Date: "& TBox_contExamDate.Text
        oPara4.Format.SpaceAfter = 30
oPara4.Range.InsertParagraphAfter()
oPara4.Range.InsertFile(TBox_DriveLetter.Text &":\StudContactReport.docx", missing, False, False, False)

        oPara5 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
oPara5.Range.InsertParagraphBefore()
        oPara5.Range.Text = "Phone Number : "& TBox_EmailAddress.Text
        oPara5.Format.SpaceAfter = 30
oPara5.Range.InsertParagraphAfter()
oPara5.Range.InsertFile(TBox_DriveLetter.Text &":\StudContactReport.docx", missing, False, False, False)

        oPara6 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
oPara6.Range.InsertParagraphBefore()
        oPara6.Range.Text = "Email Address : "& TBox_PhoneNo.Text
        oPara6.Format.SpaceAfter = 30
oPara6.Range.InsertParagraphAfter()
oPara6.Range.InsertFile(TBox_DriveLetter.Text &":\StudContactReport.docx", missing, False, False, False)

EndSub

'Create New student
PrivateSub Btn_CreateNewStudent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_CreateNewStudent.Click
Dim addFName AsString = TBox_NewFN.Text
Dim addSName AsString = TBox_NewSN.Text
Dim addCourse AsString = CBox_NewCourse.Text
Dim addPhoneNum AsString = TBox_NewPhoneNo.Text
Dim addStudComment AsString = "Enter a Comment here"
Dim addEmailAddress AsString = TBox_NewEmail.Text
Dim addStudIDno AsString = TBox_NewIDNO.Text
'If the textboxes = what they originally = then error message
If addFName = "First Name"Or addSName = "Second Name"Or addCourse = "Course"Or addPhoneNum = "Phone Number"Or addEmailAddress = "Email Address"Or addStudIDno = "Student ID number"Then
MsgBox("Please Fill all fields")
Return
EndIf


Dim DatabasePathName AsString = TBox_DriveLetter.Text &":\"& TBox_DatabaseName.Text &".accdb"
Dim ConStr AsString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="& DatabasePathName &";"
'Get the maximum stud ID
Dim SQL AsString = "SELECT MAX(StudID) FROM tblStudent"
oleCon = NewOleDbConnection(ConStr)
Dim oleComScalar AsNewOleDbCommand(SQL, oleCon)
oleComScalar.Connection.Open()
Dim NewStudID AsInteger
Try
            NewStudID = oleComScalar.ExecuteScalar()
Catch ex AsException
            NewStudID = 0
EndTry
        NewStudID += 1
'Get the maximum CourseID
        SQL = "SELECT MAX(CourseID) FROM tblCourse"
oleComScalar = NewOleDbCommand(SQL, oleCon)
Dim NewCourseID AsInteger
Try
            NewCourseID = oleComScalar.ExecuteScalar()
Catch ex AsException
            NewCourseID = 0
EndTry
        NewCourseID += 1

'Form connection string and open the database connection
oleCom = NewOleDbCommand
        oleCom.Connection = oleCon

        DA = NewOleDbDataAdapter
'tblStudent
'[ROW-x]'
        DA.InsertCommand = NewOleDbCommand("INSERT INTO tblStudent (StudID, FName, SName, IDNumber, StudComment, PhoneNo, EmailAddress, DateCreated) " _
&"VALUES ("& NewStudID &", '"& addFName &"', '"& addSName &"', '"& addStudIDno &"', "& _
"'"& addStudComment &"', '"& addPhoneNum &"', '"& addEmailAddress &"', #"&Date.Today.ToString("d") &"#)", oleCon)
Dim StudRow1 AsDataRow = DTStudents.NewRow()
StudRow1("StudID") = NewStudID
StudRow1("FName") = addFName
StudRow1("SName") = addSName
StudRow1("IDNumber") = addStudIDno
StudRow1("StudComment") = addStudComment
StudRow1("PhoneNo") = addPhoneNum
StudRow1("EmailAddress") = addEmailAddress
StudRow1("DateCreated") = Date.Today.ToString("d")
DTStudents.Rows.Add(StudRow1)
Try
DA.Update(DTStudents)

Catch ex AsException
MsgBox(ex.Message)
            NewStudID = 0
            NewCourseID = 0
Return
EndTry
DS.AcceptChanges()

'tblCourse
        DA.InsertCommand = NewOleDbCommand("INSERT INTO tblCourse (CourseID, CourseName) " _
&"VALUES ("& NewCourseID &", '"& CBox_NewCourse.Text &"')", oleCon)
Dim CourseRow1 AsDataRow = DS.Tables(1).NewRow()
CourseRow1("CourseID") = NewCourseID
CourseRow1("CourseName") = CBox_NewCourse.Text
DTCourses.Rows.Add(CourseRow1)
DA.Update(DTCourses)
DS.AcceptChanges()

'tblStudentCourse
Dim EndDate AsString
If CBox_NewCourse.Text = "Geography 2"Then
            EndDate = Date.Today.AddYears(2).ToString("d")
Else
            EndDate = Date.Today.AddYears(1).ToString("d")
EndIf

        DA.InsertCommand = NewOleDbCommand("INSERT INTO tblStudentCourse (StudID, CourseID, StartDate, EndDate, Dormant) " _
&"VALUES ("& NewStudID &", "& NewCourseID &", #"&Date.Today.ToString("d") &"#, #"& EndDate &"#, 'False')", oleCon)
Dim SCRow1 AsDataRow = DS.Tables(2).NewRow()
SCRow1("StudID") = NewStudID
SCRow1("CourseID") = NewCourseID
SCRow1("StartDate") = Date.Today.ToString("d")
SCRow1("EndDate") = EndDate
SCRow1("Dormant") = "False"
DTStudentCourses.Rows.Add(SCRow1)
DA.Update(DTStudentCourses)
DS.AcceptChanges()
oleCon.Close()
MsgBox("Added "& addFName &" "& addSName)
EndSub


'##Others
PrivateSub Btn_tmaBack_Click(sender As System.Object, e As System.EventArgs) Handles Btn_tmaBack.Click
        TCtrl_Student.SelectTab(0)
EndSub
PrivateSub Btn_contBack_Click(sender As System.Object, e As System.EventArgs) Handles Btn_contBack.Click
        TCtrl_Student.SelectTab(0)
EndSub
PrivateSub TBox_GradeAdd_Click(sender As System.Object, e As System.EventArgs) Handles TBox_GradeAdd.Click
        TBox_GradeAdd.Text = ""
EndSub
PrivateSub TPage_AllStud_Click(sender As System.Object, e As System.EventArgs) Handles TPage_AllStud.Click


EndSub
PrivateSub TCtrl_Student_SelectedIndexChanged(ByVal sender AsObject, ByVal e As System.EventArgs) Handles TCtrl_Student.SelectedIndexChanged
Dim i AsInteger
        i = TCtrl_Student.SelectedIndex
If i = 0 Then
Dim ShowDormatnOnly AsBoolean = False
DisplayStudents(ShowDormatnOnly)
'Gets the number of rows in the students table and displays it for the user
            TBox_allstudNoRows.Text = DTStudents.Rows.Count
EndIf
EndSub
PrivateSub TBox_NewFN_Click(sender As System.Object, e As System.EventArgs) Handles TBox_NewFN.Click
        TBox_NewFN.SelectAll()

EndSub
PrivateSub TBox_NewIDNO_Click(sender As System.Object, e As System.EventArgs) Handles TBox_NewIDNO.Click
        TBox_NewIDNO.SelectAll()
EndSub
PrivateSub TBox_NewPhoneNo_Click(sender As System.Object, e As System.EventArgs) Handles TBox_NewPhoneNo.Click
        TBox_NewPhoneNo.SelectAll()
EndSub
PrivateSub TBox_NewEmail_Click(sender As System.Object, e As System.EventArgs) Handles TBox_NewEmail.Click
        TBox_NewEmail.SelectAll()
EndSub
PrivateSub TBox_NewSN_Click(sender As System.Object, e As System.EventArgs) Handles TBox_NewSN.Click
        TBox_NewSN.SelectAll()
EndSub
PrivateSub TBox_Duration_Click(sender As System.Object, e As System.EventArgs) Handles TBox_Duration.Click
        TBox_Duration.SelectAll()
EndSub

'Add Stud Comment
PrivateSub Btn_SaveStudComment_Click(sender As System.Object, e As System.EventArgs) Handles Btn_SaveStudComment.Click
Dim addStudComment AsString = TBox_StudComment.Text
Dim DatabasePathName AsString = TBox_DriveLetter.Text &":\"& TBox_DatabaseName.Text &".accdb"
Dim ConStr AsString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="& DatabasePathName &";"

'Form connection string and open the database connection
oleCon = NewOleDbConnection(ConStr)
oleCon.Open()
oleCom = NewOleDbCommand
        oleCom.Connection = oleCon

'Update the comment
Dim oleStr AsString = "UPDATE tblStudent SET StudComment = '"& TBox_StudComment.Text &"' WHERE StudID = "& SSStudentID
oleCom = New OleDb.OleDbCommand(oleStr, oleCon)
oleCom.ExecuteNonQuery()
oleCon.Close()
MsgBox("CommentChanged")
Dim ShowDormantOnly AsBoolean = False
DisplayStudents(ShowDormantOnly)

EndSub


'Add Contact comment
PrivateSub Btn_SaveContCom_Click(sender As System.Object, e As System.EventArgs) Handles Btn_SaveContCom.Click
Dim addContComment AsString = TBox_contComment.Text
Dim DatabasePathName AsString = TBox_DriveLetter.Text &":\"& TBox_DatabaseName.Text &".accdb"
Dim ConStr AsString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="& DatabasePathName &";"

'Form connection string and open the database connection
oleCon = NewOleDbConnection(ConStr)
oleCon.Open()
oleCom = NewOleDbCommand
        oleCom.Connection = oleCon





'Update the comment
Dim oleStr AsString = "UPDATE tblContact SET ContactComment = '"& TBox_contComment.Text &"' WHERE CourseID = "& SSCourseID &"AND StudID = "& SSStudentID
oleCom = New OleDb.OleDbCommand(oleStr, oleCon)
oleCom.ExecuteNonQuery()
oleCon.Close()
MsgBox("CommentChanged")
Dim ShowDormantOnly AsBoolean = False
DisplayStudents(ShowDormantOnly)
EndSub


'Add TMA Comment
PrivateSub Btn_SaveTMACom_Click(sender As System.Object, e As System.EventArgs) Handles Btn_SaveTMACom.Click
Dim addTMAComment AsString = TBox_TMAComment.Text
Dim DatabasePathName AsString = TBox_DriveLetter.Text &":\"& TBox_DatabaseName.Text &".accdb"
Dim ConStr AsString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="& DatabasePathName &";"
oleCom = NewOleDbCommand
'Form connection string and open the database connection
oleCon = NewOleDbConnection(ConStr)
oleCon.Open()
        oleCom.Connection = oleCon
Dim oleStr AsString = "UPDATE tblTMA SET TMAComment='"& TBox_TMAComment.Text &"' WHERE CourseID="& SSCourseID &" AND TMAID="& SSTMAID
oleCom = NewOleDbCommand(oleStr, oleCon)
oleCom.ExecuteNonQuery()
oleCon.Close()
MsgBox("CommentChanged")
Dim ShowDormantOnly AsBoolean = False
DisplayStudents(ShowDormantOnly)
EndSub


'Dormant Buttons
PrivateSub Btn_MakeDormant_Click(sender As System.Object, e As System.EventArgs) Handles Btn_MakeDormant.Click
Dim DatabasePathName AsString = TBox_DriveLetter.Text &":\"& TBox_DatabaseName.Text &".accdb"
Dim ConStr AsString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="& DatabasePathName &";"
'Form connection string and open the database connection
oleCon = NewOleDbConnection(ConStr)
oleCon.Open()
oleCom = NewOleDbCommand
        oleCom.Connection = oleCon
Dim oleStr AsString

'Update dormant for the selected student
oleStr = "UPDATE tblStudentCourse SET Dormant = 'True' WHERE StudID = "& SSStudentID &" And CourseID = "& SSCourseID
MsgBox("Student is now dormant")
oleCom = New OleDb.OleDbCommand(oleStr, oleCon)
oleCom.ExecuteNonQuery()
oleCon.Close()
Dim ShowDormantOnly AsBoolean = False
        Btn_DisplayStudents.PerformClick()
EndSub
PrivateSub Btn_MakeUNDormant_Click(sender As System.Object, e As System.EventArgs) Handles Btn_MakeUNDormant.Click
Dim DatabasePathName AsString = TBox_DriveLetter.Text &":\"& TBox_DatabaseName.Text &".accdb"
Dim ConStr AsString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="& DatabasePathName &";"
'Form connection string and open the database connection
oleCon = NewOleDbConnection(ConStr)
oleCon.Open()
oleCom = NewOleDbCommand
        oleCom.Connection = oleCon
Dim oleStr AsString
'Update dormant for the selected student
oleStr = "UPDATE tblStudentCourse SET Dormant = 'False' WHERE StudID = "& SSStudentID &" And CourseID = "& SSCourseID
MsgBox("Student is now NOT dormant")
oleCom = New OleDb.OleDbCommand(oleStr, oleCon)
oleCom.ExecuteNonQuery()
oleCon.Close()
        Btn_DisplayStudents.PerformClick()

EndSub
EndClass
