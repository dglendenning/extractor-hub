#Benchmark Extractor.py

#import easygui as eg
import pyodbc
import numpy as np
import pandas as pd
import os.path
import sqlalchemy as sa
import datetime as dt
import mal_data as mal

def parcc_or_pssa(cnxn, districtID):
    df = pd.read_sql("SELECT StateID from District with (nolock) WHERE DistrictID = "+str(districtID), cnxn)
    if (df["StateID"][0]==49):
        print("New Jersey District, StateID:",df["StateID"][0])
        return 'PARCC', 217
    print("Non-New Jersey District, StateID:",df["StateID"][0])
    return 'PSSA', 109

#Checks first for an 18 in the term name, then for a 16. If neither, actually checks the date.
def clean_term(row):
    if '18' in row.DistrictTerm:
        row.DistrictTerm = '2017-18'
    elif '16' in row.DistrictTerm:
        row.DistrictTerm = '2016-17'
    else:
        #print("Non-Fatal Error: Could not parse DistrictTerm ",row.DistrictTerm)
        if row.ResultDate > dt.datetime(2017, 7, 30):
            row.DistrictTerm = '2017-18'
            #print(" -Changed to",row.DistrictTerm,"based on date",row.ResultDate)
        elif row.ResultDate > dt.datetime(2016, 7, 30):
            row.DistrictTerm = '2016-17'
            #print(" -Changed to",row.DistrictTerm,"based on date",row.ResultDate)

    return row.DistrictTerm

#Checks first for an 18 in the term name, then for a 16. If neither, actually checks the date.
def clean_term2(row):
    if '18' in row.TermName:
        row.TermName = '2017-18'
    elif '16' in row.TermName:
        row.TermName = '2016-17'
    else:
        #print("Non-Fatal Error: Could not parse TermName",row.TermName)
        if row.MostRecentDate > dt.datetime(2017, 7, 30):
            row.TermName = '2017-18'
            #print(" -Changed to",row.TermName,"based on date",row.MostRecentDate)
        elif row.MostRecentDate > dt.datetime(2016, 7, 30):
            row.TermName = '2016-17'
            #print(" -Changed to",row.TermName,"based on date",row.MostRecentDate)

    return row.TermName

def extract(districtID):

    districtID = str(districtID)
    cnxn = mal.setup_FTP()
    box = pd.read_sql("select Name from District with (nolock) WHERE DistrictID = \'"+districtID+"\'", cnxn)
    dname = box.at[0, 'Name']
    state_test, achievement_level = parcc_or_pssa(cnxn, districtID)
    extracts = os.path.join(os.getcwd(),"Extracts")
    if not os.path.exists(extracts):
        os.makedirs(extracts)
    n = os.path.join(extracts,dname+' Form B Data 2017-18.xlsx')
    writer = pd.ExcelWriter(n)

    box = pd.read_sql("""declare @districtid int, @resultdate datetime
set @districtid="""+districtID+"""
set @resultdate='2016-08-01'

---local assessment---
select tr.UpdatedDate as [ResultDate], dt.name as DistrictTerm, vt.name as TestName, sub.name as Subject, gr.name as Grade, sch.Name as School, u.UserID, u.code as TeacherCode, u.NameFirst as TeacherFirstName, u.NameLast as TeacherLastName, c.ClassID, c.Name as ClassName, s.StudentID, s.code as StudentCode, s.FirstName as StudentFirstName, s.LastName as StudentLastName, trs.ScoreRaw as TotalPointsEarned, trs.PointsPossible as TotalPointsPossible from testresultscore trs With (nolock)
join testresult tr With (nolock) on tr.testresultid=trs.testresultid
join virtualtest vt With (nolock) on vt.VirtualTestID=tr.VirtualTestID
join bank b With (nolock) on b.BankID =vt.bankid
join subject sub With (nolock) on sub.subjectid=b.SubjectID
join grade gr With (nolock) on gr.GradeID=sub.GradeID
join student s With (nolock) on s.studentid=tr.StudentID
join school sch With (nolock) on sch.SchoolID=tr.SchoolID
join class c With (nolock) on c.classid=tr.classid
join DistrictTerm dt With (nolock) on dt.DistrictTermID=c.DistrictTermID
join [user] u With (nolock) on u.userid=tr.userid
where dt.DistrictID=@districtid and tr.ResultDate>@resultdate and vt.Name like '%LinkIt%form%'
and not (vt.Name like '%Link%it%form%CR%' or vt.Name like '%Retake%')
order by tr.UpdatedDate desc""",cnxn, parse_dates = ['ResultDate'])


    if(not box.empty):
        box['Form'] = box.apply(lambda row: row.TestName[row.TestName.find("Form") + 5] if ("Form" in row.TestName) else " ", axis = 1)
        box['DistrictTerm'] = box.apply(lambda row: clean_term(row), axis = 1)
        box['ResultDate'] = box.apply(lambda row: f'{row.ResultDate.month}/{row.ResultDate.day}/{row.ResultDate.year}', axis = 1)
        box = box[['ResultDate','DistrictTerm','TestName','Subject','Grade','Form','School','UserID','TeacherCode','TeacherFirstName', 'TeacherLastName','ClassID','ClassName','StudentID', 'StudentCode','StudentFirstName','StudentLastName','TotalPointsEarned','TotalPointsPossible']]
        box.to_excel(writer, sheet_name = 'Linkit Benchmarks', index = False)


    box2 = pd.read_sql("""declare @districtid int, @resultdate datetime
set @districtid="""+districtID+"""
set @resultdate='2016-08-01'

select vt.Name as TestName, sub.name as Subject, gr.Name as Grade, sch.SchoolID, sch.name as SchoolName, s.StudentID, s.code as StudentCode, s.FirstName as StudentFirstName, s.LastName as StudentLastName, trs.ScoreScaled as ScaledScore, trs.AchievementLevel  from TestResultScore trs With (nolock)
join testresult tr With (nolock) on tr.testresultid=trs.testresultid
join virtualtest vt With (nolock) on vt.VirtualTestID=tr.VirtualTestID
join bank b With (nolock) on b.BankID =vt.bankid
join subject sub With (nolock) on sub.subjectid=b.SubjectID
join grade gr With (nolock) on gr.GradeID=sub.GradeID
join student s With (nolock) on s.studentid=tr.StudentID
join school sch With (nolock) on sch.SchoolID=tr.SchoolID
where sch.DistrictID=@districtid and vt.Name like '20%-20%"""+state_test+"""%' and vt.achievementlevelsettingid="""+str(achievement_level), cnxn)

    box2['Year'] = box2.apply(lambda row: row.TestName[5:9], axis = 1)
    box2 = box2[['Year','TestName','Subject','Grade','SchoolID','SchoolName','StudentID','StudentCode','StudentFirstName','StudentLastName','ScaledScore','AchievementLevel']]
    box2.to_excel(writer, sheet_name = state_test, index = False)
    #writer.save()
    cnxn = pyodbc.connect('DRIVER=***REMOVED***;SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)

    #Standards

    box3 = pd.read_sql("""declare @districtid int, @resultdate datetime
set @districtid="""+districtID+"""
set @resultdate='2017-08-01'



----Standards---
DECLARE @testdata3 TABLE (ID int IDENTITY(1,1), VirtualTestID INT, TestName VARCHAR(1000), Subject VARCHAR(1000), Grade VARCHAR(25), SchoolID INT,
SchoolName VARCHAR(1000), TestResultID INT, UserID INT, TeacherCode VARCHAR(1000), TeacherFirstName VARCHAR (100), TeacherLastName VARCHAR(1000), ClassID INT, ClassName VARCHAR(1000), TermName VARCHAR(1000), updatedate datetime)

DECLARE @testdata3a TABLE (ID int IDENTITY(1,1), StandardNbr VARCHAR(1200), TotalPointsEarned INT, TotalPointsPossible INT, ClassID INT, UserID INT, VirtualTestID INT, TestResultID INT)
SET NOCOUNT ON

INSERT INTO @testdata3 (VirtualTestID, SchoolID, UserID, ClassID, TestResultID, updatedate)
SELECT tr.VirtualTestID, tr.SchoolID, tr.UserID, tr.ClassID, tr.TestResultID, tr.UpdatedDate
FROM TestResult AS tr INNER JOIN
	[User] AS u ON tr.UserID = u.UserID
WHERE    (u.DistrictID = @DistrictID) AND (tr.ResultDate > @resultdate)

SET NOCOUNT ON
UPDATE @testdata3
SET SchoolName = sc.Name
FROM @testdata3 AS td INNER JOIN
	School AS sc ON td.SchoolID = sc.SchoolID

SET NOCOUNT ON
UPDATE @testdata3
SET Subject = sb.Name, Grade = g.Name, TestName = vt.Name
FROM @testdata3 AS td INNER JOIN
	VirtualTest AS vt ON td.VirtualTestID = vt.VirtualTestID INNER JOIN
	Bank AS b ON vt.BankID = b.BankID INNER JOIN
	Subject AS sb ON b.SubjectID = sb.SubjectID INNER JOIN
	Grade AS g ON sb.GradeID = g.GradeID

SET NOCOUNT ON
UPDATE @testdata3
SET ClassName = c.Name, TermName = dt.Name
FROM @testdata3 AS td INNER JOIN
	Class AS c ON td.ClassID = c.ClassID INNER JOIN
	DistrictTerm AS dt ON c.DistrictTermID = dt.DistrictTermID

SET NOCOUNT ON
UPDATE @testdata3
SET TeacherCode = u.Code, TeacherFirstName = u.NameFirst, TeacherLastName = u.NameLast
FROM @testdata3 AS td INNER JOIN
	[User] AS u ON td.UserID = u.UserID


SET NOCOUNT ON
INSERT INTO @testdata3a (StandardNbr, ClassID, UserID, VirtualTestID, TotalPointsEarned, TotalPointsPossible, TestResultID)
SELECT ms.Number, td.ClassID, td.UserID, td.VirtualTestID, a.PointsEarned, a.PointsPossible, td.TestResultID
FROM @testdata3 AS td INNER JOIN
	Answer AS a ON td.TestResultID = a.TestResultID INNER JOIN
	VirtualQuestion AS vq ON a.VirtualQuestionID = vq.VirtualQuestionID INNER JOIN
	QTIItem AS q ON vq.QTIItemID = q.QTIItemID INNER JOIN
	VirtualQuestionStateStandard AS vqss ON vq.VirtualQuestionID = vqss.VirtualQuestionID INNER JOIN
	MasterStandard AS ms ON vqss.StateStandardID = ms.MasterStandardID
--where q.QTISchemaID!=10

SELECT min(td.updatedate) as EarliestDate,max(td.updatedate) as MostRecentDate, td.TermName, td.TestName, td.Subject, td.Grade, td.SchoolName, td.UserID, td.TeacherCode, td.TeacherFirstName, td.TeacherLastName, td.ClassID, td.ClassName,  tda.StandardNbr, SUM(tda.TotalPointsEarned)As TotalPointsEarned, SUM(tda.TotalPointsPossible) AS TotalPointsPossible
FROM @testdata3 AS td INNER JOIN
	@testdata3a AS tda ON td.TestResultID = tda.TestResultID
	where td.TestName like '%linkit%form%' and not (td.TestName like '%Link%it%form%CR%' or td.TestName like '%Retake%')
GROUP BY td.TermName, td.TestName, td.Subject, td.Grade, td.SchoolName, td.UserID, td.TeacherCode, td.TeacherFirstName, td.TeacherLastName, td.ClassID, td.ClassName,  tda.StandardNbr""", cnxn, parse_dates = ["MostRecentDate"])
    if(not box3.empty):
        box3['Form'] = box3.apply(lambda row: row.TestName[row.TestName.find("Form") + 5], axis = 1)
        box3['TermName'] = box3.apply(lambda row: clean_term2(row), axis = 1)
        box3 = box3[['TermName','TestName','Subject','Grade','Form','SchoolName', 'UserID','TeacherCode','TeacherFirstName', 'TeacherLastName','ClassID','ClassName','StandardNbr', 'TotalPointsEarned', 'TotalPointsPossible']]
        box3.to_excel(writer, sheet_name = 'Standards', index = False)

    #writer.save()

    #Skills

    box4 = pd.read_sql("""declare @districtid int, @resultdate datetime
set @districtid="""+districtID+"""
set @resultdate='2017-08-01'
SET NOCOUNT ON

--- Skills---
DECLARE @testdata3 TABLE (ID int IDENTITY(1,1), VirtualTestID INT, TestName VARCHAR(1000), Subject VARCHAR(1000), Grade VARCHAR(25), SchoolID INT,
SchoolName VARCHAR(1000), TestResultID INT, UserID INT, TeacherCode VARCHAR(1000), TeacherFirstName VARCHAR (100), TeacherLastName VARCHAR(1000), ClassID INT, ClassName VARCHAR(1000), TermName VARCHAR(1000), updateddate datetime)

DECLARE @testdata3a TABLE (ID int IDENTITY(1,1), StandardNbr VARCHAR(1200), TotalPointsEarned INT, TotalPointsPossible INT, ClassID INT, UserID INT, VirtualTestID INT, TestResultID INT)

INSERT INTO @testdata3 (VirtualTestID, SchoolID, UserID, ClassID, TestResultID, updateddate)
SELECT tr.VirtualTestID, tr.SchoolID, tr.UserID, tr.ClassID, tr.TestResultID, tr.updateddate
FROM TestResult AS tr INNER JOIN
	[User] AS u ON tr.UserID = u.UserID
WHERE    (u.DistrictID = @districtid) AND (tr.ResultDate > @resultdate)


UPDATE @testdata3
SET SchoolName = sc.Name
FROM @testdata3 AS td INNER JOIN
	School AS sc ON td.SchoolID = sc.SchoolID

UPDATE @testdata3
SET Subject = sb.Name, Grade = g.Name, TestName = vt.Name
FROM @testdata3 AS td INNER JOIN
	VirtualTest AS vt ON td.VirtualTestID = vt.VirtualTestID INNER JOIN
	Bank AS b ON vt.BankID = b.BankID INNER JOIN
	Subject AS sb ON b.SubjectID = sb.SubjectID INNER JOIN
	Grade AS g ON sb.GradeID = g.GradeID

UPDATE @testdata3
SET ClassName = c.Name, TermName = dt.Name
FROM @testdata3 AS td INNER JOIN
	Class AS c ON td.ClassID = c.ClassID INNER JOIN
	DistrictTerm AS dt ON c.DistrictTermID = dt.DistrictTermID

UPDATE @testdata3
SET TeacherCode = u.Code, TeacherFirstName = u.NameFirst, TeacherLastName = u.NameLast
FROM @testdata3 AS td INNER JOIN
	[User] AS u ON td.UserID = u.UserID


INSERT INTO @testdata3a (StandardNbr, ClassID, UserID, VirtualTestID, TotalPointsEarned, TotalPointsPossible, TestResultID)
SELECT ms.Name, td.ClassID, td.UserID, td.VirtualTestID, a.PointsEarned, a.PointsPossible, td.TestResultID
FROM @testdata3 AS td INNER JOIN
	Answer AS a ON td.TestResultID = a.TestResultID INNER JOIN
	VirtualQuestion AS vq ON a.VirtualQuestionID = vq.VirtualQuestionID INNER JOIN
	QTIItem AS q ON vq.QTIItemID = q.QTIItemID INNER JOIN
	VirtualQuestionLessonOne AS vqss ON vq.VirtualQuestionID = vqss.VirtualQuestionID INNER JOIN
	LessonOne AS ms ON vqss.LessonOneID = ms.LessonOneID
	--where q.QTISchemaID!=10



SELECT min(td.updateddate) as EarliestDate ,max(td.updateddate) as MostRecentDate,  td.TermName, td.TestName, td.Subject, td.Grade, td.SchoolName, td.UserID, td.TeacherCode, td.TeacherFirstName, td.TeacherLastName, td.ClassID, td.ClassName,  replace(replace(tda.StandardNbr, char(10), ''), char(13), '') as Skills , SUM(tda.TotalPointsEarned)As TotalPointsEarned, SUM(tda.TotalPointsPossible) AS TotalPointsPossible
FROM @testdata3 AS td INNER JOIN
	@testdata3a AS tda ON td.TestResultID = tda.TestResultID
	where td.TestName like '%linkit%form%' and not (td.TestName like '%Link%it%form%CR%' or td.TestName like '%Retake%')
GROUP BY td.TermName, td.TestName, td.Subject, td.Grade, td.SchoolName, td.UserID, td.TeacherCode, td.TeacherFirstName, td.TeacherLastName, td.ClassID, td.ClassName,  tda.StandardNbr""", cnxn, parse_dates = ["MostRecentDate"])

    if(not box4.empty):
        box4['Form'] = box4.apply(lambda row: row.TestName[row.TestName.find("Form") + 5], axis = 1)
        box4['TermName'] = box4.apply(lambda row: clean_term2(row), axis = 1)
        box4 = box4[['TermName','TestName','Subject','Grade','Form','SchoolName', 'UserID','TeacherCode','TeacherFirstName', 'TeacherLastName','ClassID','ClassName','Skills', 'TotalPointsEarned', 'TotalPointsPossible']]
        box4.to_excel(writer, sheet_name = 'Skills', index = False)


    #Gender
    boxg = pd.read_sql("""declare @district int
set @district="""+districtID+"""

select s.StudentID, g.name as Gender from student s With (nolock)
join gender g With (nolock) on g.GenderID=s.GenderID
where s.DistrictID=@district
""", cnxn)

    boxg.to_excel(writer, sheet_name = 'Gender', index = False)

    #Race
    boxg = pd.read_sql("""declare @district int
set @district="""+districtID+"""

select s.StudentID, r.name as Race from student s With (nolock)
join race r With (nolock) on r.raceid=s.raceid
where s.districtid=@district
""", cnxn)

    boxg.to_excel(writer, sheet_name = 'Race', index = False)

    #Program
    boxg = pd.read_sql("""declare @district int
set @district="""+districtID+"""

select sp.StudentID, p.name as Program from studentprogram sp With (nolock)
join program p With (nolock) on p.programid=sp.programid
where p.districtid=@district
""", cnxn)

    boxg.to_excel(writer, sheet_name = 'Program', index = False)

    if state_test == 'PARCC':
        #Standards by Gender
        boxg = pd.read_sql("""declare @district int
set @district="""+districtID+"""

SELECT	SCHOOL.NAME AS 'School',
		Virtualtest.name as 'Test Name',
		Gender.name as 'Gender',
		MasterStandard.Number,
		SUM(ANSWER.POINTSEARNED) AS 'Points Earned',
		sum(answer.PointsPossible) as 'Points Possible'
FROM VIRTUALTEST WITH (NOLOCK)
	INNER JOIN TestResult WITH (NOLOCK) ON TESTRESULT.VIRTUALTESTID = VIRTUALTEST.VIRTUALTESTID
	INNER JOIN SCHOOL WITH (NOLOCK) ON SCHOOL.SCHOOLID = TESTRESULT.SCHOOLID
	INNER JOIN STUDENT WITH (NOLOCK) ON STUDENT.STUDENTID = TESTRESULT.STUDENTID
	INNER JOIN GENDER WITH (NOLOCK) ON STUDENT.GenderID = GENDER.GenderID
	INNER JOIN VirtualQuestion WITH (NOLOCK) ON VIRTUALQUESTION.VIRTUALTESTID = VIRTUALTEST.VIRTUALTESTID
	INNER JOIN VIRTUALQUESTIONSTATESTANDARD WITH (NOLOCK) ON VIRTUALQUESTIONSTATESTANDARD.VIRTUALQUESTIONID = VIRTUALQUESTION.VIRTUALQUESTIONID
	INNER JOIN MASTERSTANDARD WITH (NOLOCK) ON MASTERSTANDARD.MASTERSTANDARDID = VIRTUALQUESTIONSTATESTANDARD.STATESTANDARDID
	INNER JOIN ANSWER WITH (NOLOCK) ON ANSWER.TESTRESULTID = TESTRESULT.TESTRESULTID AND ANSWER.VIRTUALQUESTIONID = VIRTUALQUESTION.VIRTUALQUESTIONID
WHERE VIRTUALTEST.BANKID IN (
59164,59163,59075,59089,59076,59091,59077,59090,59079,59092,59081,59093,59084,
59094,59085,59095,59086,59096,59074,59088,59087,59098,60246,59123,59134,59124,
59135,59125,59136,59126,59137,59127,59138,59128,59139,59129,59140,59130,59141,
59131,59133,59132,59142,59171,59170)
AND STUDENT.DISTRICTID = @district
GROUP BY SCHOOL.Name,
			VIRTUALTEST.Name,
			GENDER.Name,
			MasterStandard.Number
order by School.Name,
			virtualtest.Name,
			MasterStandard.Number,
			Gender.Name
""", cnxn)

        boxg.to_excel(writer, sheet_name = 'Standards by Gender', index = False)

        #Standards by Race
        boxg = pd.read_sql("""declare @district int
set @district="""+districtID+"""



SELECT	SCHOOL.NAME AS 'School',
		Virtualtest.name as 'Test Name',
		Race.name as 'Race',
		MasterStandard.Number,
		SUM(ANSWER.POINTSEARNED) AS 'Points Earned',
		sum(answer.PointsPossible) as 'Points Possible'
FROM VIRTUALTEST WITH (NOLOCK)
	INNER JOIN TestResult WITH (NOLOCK) ON TESTRESULT.VIRTUALTESTID = VIRTUALTEST.VIRTUALTESTID
	INNER JOIN SCHOOL WITH (NOLOCK) ON SCHOOL.SCHOOLID = TESTRESULT.SCHOOLID
	INNER JOIN STUDENT WITH (NOLOCK) ON STUDENT.STUDENTID = TESTRESULT.STUDENTID
	INNER JOIN Race WITH (NOLOCK) ON STUDENT.RaceID = RACE.RaceID
	INNER JOIN VirtualQuestion WITH (NOLOCK) ON VIRTUALQUESTION.VIRTUALTESTID = VIRTUALTEST.VIRTUALTESTID
	INNER JOIN VIRTUALQUESTIONSTATESTANDARD WITH (NOLOCK) ON VIRTUALQUESTIONSTATESTANDARD.VIRTUALQUESTIONID = VIRTUALQUESTION.VIRTUALQUESTIONID
	INNER JOIN MASTERSTANDARD WITH (NOLOCK) ON MASTERSTANDARD.MASTERSTANDARDID = VIRTUALQUESTIONSTATESTANDARD.STATESTANDARDID
	INNER JOIN ANSWER WITH (NOLOCK) ON ANSWER.TESTRESULTID = TESTRESULT.TESTRESULTID AND ANSWER.VIRTUALQUESTIONID = VIRTUALQUESTION.VIRTUALQUESTIONID
WHERE VIRTUALTEST.BANKID IN (
59164,59163,59075,59089,59076,59091,59077,59090,59079,59092,59081,59093,59084,
59094,59085,59095,59086,59096,59074,59088,59087,59098,60246,59123,59134,59124,
59135,59125,59136,59126,59137,59127,59138,59128,59139,59129,59140,59130,59141,
59131,59133,59132,59142,59171,59170)
AND STUDENT.DISTRICTID = @district
GROUP BY SCHOOL.Name,
			VIRTUALTEST.Name,
			Race.Name,
			MasterStandard.Number
ORDER BY SCHOOL.Name,
			VIRTUALTEST.Name,
			MASTERSTANDARD.Number,
			RACE.Name
""", cnxn)

        boxg.to_excel(writer, sheet_name = 'Standards by Race', index = False)

        #Standards by Program
        boxg = pd.read_sql("""declare @district int
set @district="""+districtID+"""



SELECT	SCHOOL.NAME AS 'School',
		Virtualtest.name as 'Test Name',
		PROGRAM.name as 'Program',
		MasterStandard.Number,
		SUM(ANSWER.POINTSEARNED) AS 'Points Earned',
		sum(answer.PointsPossible) as 'Points Possible'
FROM VIRTUALTEST WITH (NOLOCK)
	INNER JOIN TestResult WITH (NOLOCK) ON TESTRESULT.VIRTUALTESTID = VIRTUALTEST.VIRTUALTESTID
	INNER JOIN SCHOOL WITH (NOLOCK) ON SCHOOL.SCHOOLID = TESTRESULT.SCHOOLID
	INNER JOIN STUDENT WITH (NOLOCK) ON STUDENT.STUDENTID = TESTRESULT.STUDENTID
	INNER JOIN TestResultProgram WITH (NOLOCK) ON TestResultProgram.TestResultID = testresult.TestResultID
	INNER JOIN Program WITH (NOLOCK) ON TestResultProgram.PROGRAMID = PROGRAM.PROGRAMID
	INNER JOIN VirtualQuestion WITH (NOLOCK) ON VIRTUALQUESTION.VIRTUALTESTID = VIRTUALTEST.VIRTUALTESTID
	INNER JOIN VIRTUALQUESTIONSTATESTANDARD WITH (NOLOCK) ON VIRTUALQUESTIONSTATESTANDARD.VIRTUALQUESTIONID = VIRTUALQUESTION.VIRTUALQUESTIONID
	INNER JOIN MASTERSTANDARD WITH (NOLOCK) ON MASTERSTANDARD.MASTERSTANDARDID = VIRTUALQUESTIONSTATESTANDARD.STATESTANDARDID
	INNER JOIN ANSWER WITH (NOLOCK) ON ANSWER.TESTRESULTID = TESTRESULT.TESTRESULTID AND ANSWER.VIRTUALQUESTIONID = VIRTUALQUESTION.VIRTUALQUESTIONID
WHERE VIRTUALTEST.BANKID IN (
59164,59163,59075,59089,59076,59091,59077,59090,59079,59092,59081,59093,59084,
59094,59085,59095,59086,59096,59074,59088,59087,59098,60246,59123,59134,59124,
59135,59125,59136,59126,59137,59127,59138,59128,59139,59129,59140,59130,59141,
59131,59133,59132,59142,59171,59170)
AND STUDENT.DISTRICTID = @district
GROUP BY SCHOOL.Name,
			VIRTUALTEST.Name,
			Program.Name,
			MasterStandard.Number
ORDER BY SCHOOL.Name,
			VIRTUALTEST.Name,
			MASTERSTANDARD.Number,
			PROGRAM.Name
""", cnxn)


        boxg.to_excel(writer, sheet_name = 'Standards by Program', index = False)
    elif state_test == "PSSA":
        #Standards by Gender
        boxg = pd.read_sql("""declare @district int
set @district="""+districtID+"""

SELECT	SCHOOL.NAME AS 'School',
		Virtualtest.name as 'Test Name',
		Gender.name as 'Gender',
		MasterStandard.Number,
		SUM(ANSWER.POINTSEARNED) AS 'Points Earned',
		sum(answer.PointsPossible) as 'Points Possible'
FROM VIRTUALTEST WITH (NOLOCK)
	INNER JOIN TestResult WITH (NOLOCK) ON TESTRESULT.VIRTUALTESTID = VIRTUALTEST.VIRTUALTESTID
	INNER JOIN SCHOOL WITH (NOLOCK) ON SCHOOL.SCHOOLID = TESTRESULT.SCHOOLID
	INNER JOIN STUDENT WITH (NOLOCK) ON STUDENT.STUDENTID = TESTRESULT.STUDENTID
	INNER JOIN GENDER WITH (NOLOCK) ON STUDENT.GenderID = GENDER.GenderID
	INNER JOIN VirtualQuestion WITH (NOLOCK) ON VIRTUALQUESTION.VIRTUALTESTID = VIRTUALTEST.VIRTUALTESTID
	INNER JOIN VIRTUALQUESTIONSTATESTANDARD WITH (NOLOCK) ON VIRTUALQUESTIONSTATESTANDARD.VIRTUALQUESTIONID = VIRTUALQUESTION.VIRTUALQUESTIONID
	INNER JOIN MASTERSTANDARD WITH (NOLOCK) ON MASTERSTANDARD.MASTERSTANDARDID = VIRTUALQUESTIONSTATESTANDARD.STATESTANDARDID
	INNER JOIN ANSWER WITH (NOLOCK) ON ANSWER.TESTRESULTID = TESTRESULT.TESTRESULTID AND ANSWER.VIRTUALQUESTIONID = VIRTUALQUESTION.VIRTUALQUESTIONID
WHERE VIRTUALTEST.BANKID IN (
58520,58522,58523,58524,58525,58526,58527,58528,58529,58530,
58531,58532,58667,58668,58669,58670,58671,58672,58922,58923)
AND STUDENT.DISTRICTID = @district
GROUP BY SCHOOL.Name,
			VIRTUALTEST.Name,
			GENDER.Name,
			MasterStandard.Number
order by School.Name,
			virtualtest.Name,
			MasterStandard.Number,
			Gender.Name
""", cnxn)

        boxg.to_excel(writer, sheet_name = 'Standards by Gender', index = False)

        #Standards by Race
        boxg = pd.read_sql("""declare @district int
set @district="""+districtID+"""



SELECT	SCHOOL.NAME AS 'School',
		Virtualtest.name as 'Test Name',
		Race.name as 'Race',
		MasterStandard.Number,
		SUM(ANSWER.POINTSEARNED) AS 'Points Earned',
		sum(answer.PointsPossible) as 'Points Possible'
FROM VIRTUALTEST WITH (NOLOCK)
	INNER JOIN TestResult WITH (NOLOCK) ON TESTRESULT.VIRTUALTESTID = VIRTUALTEST.VIRTUALTESTID
	INNER JOIN SCHOOL WITH (NOLOCK) ON SCHOOL.SCHOOLID = TESTRESULT.SCHOOLID
	INNER JOIN STUDENT WITH (NOLOCK) ON STUDENT.STUDENTID = TESTRESULT.STUDENTID
	INNER JOIN Race WITH (NOLOCK) ON STUDENT.RaceID = RACE.RaceID
	INNER JOIN VirtualQuestion WITH (NOLOCK) ON VIRTUALQUESTION.VIRTUALTESTID = VIRTUALTEST.VIRTUALTESTID
	INNER JOIN VIRTUALQUESTIONSTATESTANDARD WITH (NOLOCK) ON VIRTUALQUESTIONSTATESTANDARD.VIRTUALQUESTIONID = VIRTUALQUESTION.VIRTUALQUESTIONID
	INNER JOIN MASTERSTANDARD WITH (NOLOCK) ON MASTERSTANDARD.MASTERSTANDARDID = VIRTUALQUESTIONSTATESTANDARD.STATESTANDARDID
	INNER JOIN ANSWER WITH (NOLOCK) ON ANSWER.TESTRESULTID = TESTRESULT.TESTRESULTID AND ANSWER.VIRTUALQUESTIONID = VIRTUALQUESTION.VIRTUALQUESTIONID
WHERE VIRTUALTEST.BANKID IN (
58520,58522,58523,58524,58525,58526,58527,58528,58529,58530,
58531,58532,58667,58668,58669,58670,58671,58672,58922,58923)
AND STUDENT.DISTRICTID = @district
GROUP BY SCHOOL.Name,
			VIRTUALTEST.Name,
			Race.Name,
			MasterStandard.Number
ORDER BY SCHOOL.Name,
			VIRTUALTEST.Name,
			MASTERSTANDARD.Number,
			RACE.Name
""", cnxn)

        boxg.to_excel(writer, sheet_name = 'Standards by Race', index = False)

        #Standards by Program
        boxg = pd.read_sql("""declare @district int
set @district="""+districtID+"""



SELECT	SCHOOL.NAME AS 'School',
		Virtualtest.name as 'Test Name',
		PROGRAM.name as 'Program',
		MasterStandard.Number,
		SUM(ANSWER.POINTSEARNED) AS 'Points Earned',
		sum(answer.PointsPossible) as 'Points Possible'
FROM VIRTUALTEST WITH (NOLOCK)
	INNER JOIN TestResult WITH (NOLOCK) ON TESTRESULT.VIRTUALTESTID = VIRTUALTEST.VIRTUALTESTID
	INNER JOIN SCHOOL WITH (NOLOCK) ON SCHOOL.SCHOOLID = TESTRESULT.SCHOOLID
	INNER JOIN STUDENT WITH (NOLOCK) ON STUDENT.STUDENTID = TESTRESULT.STUDENTID
	INNER JOIN TestResultProgram WITH (NOLOCK) ON TestResultProgram.TestResultID = testresult.TestResultID
	INNER JOIN Program WITH (NOLOCK) ON TestResultProgram.PROGRAMID = PROGRAM.PROGRAMID
	INNER JOIN VirtualQuestion WITH (NOLOCK) ON VIRTUALQUESTION.VIRTUALTESTID = VIRTUALTEST.VIRTUALTESTID
	INNER JOIN VIRTUALQUESTIONSTATESTANDARD WITH (NOLOCK) ON VIRTUALQUESTIONSTATESTANDARD.VIRTUALQUESTIONID = VIRTUALQUESTION.VIRTUALQUESTIONID
	INNER JOIN MASTERSTANDARD WITH (NOLOCK) ON MASTERSTANDARD.MASTERSTANDARDID = VIRTUALQUESTIONSTATESTANDARD.STATESTANDARDID
	INNER JOIN ANSWER WITH (NOLOCK) ON ANSWER.TESTRESULTID = TESTRESULT.TESTRESULTID AND ANSWER.VIRTUALQUESTIONID = VIRTUALQUESTION.VIRTUALQUESTIONID
WHERE VIRTUALTEST.BANKID IN (
58520,58522,58523,58524,58525,58526,58527,58528,58529,58530,
58531,58532,58667,58668,58669,58670,58671,58672,58922,58923)
AND STUDENT.DISTRICTID = @district
GROUP BY SCHOOL.Name,
			VIRTUALTEST.Name,
			Program.Name,
			MasterStandard.Number
ORDER BY SCHOOL.Name,
			VIRTUALTEST.Name,
			MASTERSTANDARD.Number,
			PROGRAM.Name
""", cnxn)




    #do all .to_excel calls (USE WRITER) before calling this
    writer.save()
    return (dname+""" Benchmark Extract created and saved sucessfully.
Location: """+n)



#if __name__ == "__main__":
#    cont = 1
#    while cont:
#        districtID = eg.integerbox("Please enter the District ID:", upperbound = 9999999)
#        if districtID == None:
#            break
#        msg = extract(districtID)
#        cont = eg.ccbox(msg+"""
#    Do you wish to enter another ID?""")
