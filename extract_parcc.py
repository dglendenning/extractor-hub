"""Create PARCC report data extract."""
import pandas as pd
import os
import os.path
import mal_data as mal


def extract(districtID):
    """Create the extract and save it as a .xlsx file."""
    # Get a connection to the database.
    database = mal.setup_SQL()
    # Query the database for the district name.
    district_name = mal.get_district_name(districtID, database)
    # Make sure /Extracts directory exists and sets the output path.
    extracts = os.path.join(os.getcwd(), "Extracts")
    if not os.path.exists(extracts):
        os.makedirs(extracts)
    file_name = os.path.join(
        extracts, '{} 3-Year PARCC Data.xlsx'.format(district_name))
    # Get our output file set up for writing.
    file = mal.setup_writer(name=file_name)

    # Query the database and store it in a pandas DataFrame.
    score = pd.read_sql(
        """declare @districtid int
        set @districtid={}
        select vt.Name as TestName, sub.name as Subject, gr.Name as
        Grade, sch.Name as SchoolID, s.StudentID, s.code as StudentCode,
        s.FirstName as StudentFirstName, s.LastName as StudentLastName,
        trs.ScoreScaled as ScaledScore, trs.AchievementLevel as
        ProfLevel from TestResultScore trs With (nolock)
        join testresult tr With (nolock) on
        tr.testresultid=trs.testresultid
        join virtualtest vt With (nolock) on
        vt.VirtualTestID=tr.VirtualTestID
        join bank b With (nolock) on b.BankID =vt.bankid
        join subject sub With (nolock) on sub.subjectid=b.SubjectID
        join grade gr With (nolock) on gr.GradeID=sub.GradeID
        join student s With (nolock) on s.studentid=tr.StudentID
        join school sch With (nolock) on sch.SchoolID=tr.SchoolID
        where sch.DistrictID=@districtid
        and (vt.Name like '20%-20%PARCC%' and vt.Name not like '%N/A%')
        and vt.achievementlevelsettingid=217""".format(districtID),
        database)

    score['Year'] = score.apply(lambda row: row.TestName[5:9], axis=1)
    score['NAVGrade'] = score.apply(lambda row:
                                    11 if 'Alg II' in row.TestName else
                                    10 if 'Geo' in row.TestName else
                                    9 if 'Alg I' in row.TestName else
                                    row.Grade, axis=1)
    score.loc[score.Subject == 'Language Arts', 'Subject'] = 'ELA'
    score = score[[
        'Year', 'TestName', 'Subject', 'Grade', 'NAVGrade',
        'SchoolID', 'StudentID', 'StudentCode', 'StudentFirstName',
        'StudentLastName', 'ScaledScore', 'ProfLevel']]
    score.to_excel(file, sheet_name='Score', index=False)

    del score

    cluster = pd.read_sql(
        """Declare @districtid int
        set @districtid = """
        + districtID +
        """


        Select
        District.Name as [District],
        TestResult.ResultDate as [Date],
        VirtualTest.Name as [TestName],
        School.Name as [School],
        Class.ClassID AS [ClassID],
        Class.Name AS [Class Name],
        TestResultSubScore.Name AS [ClusterName],
        TestResultSubScore.ScoreScaled AS [Score],
        TestResultSubScore.AchievementLevel AS [Prof]

        from TestResult with (nolock)
            inner join TestResultScore as trs with (nolock) on
            trs.TestResultID=TestResult.TestResultID
            inner join VirtualTest with (nolock) on
            TestResult.VirtualTestID=VirtualTest.VirtualTestID
            inner join DistrictTerm with (nolock) on
            TestResult.DistrictTermID=DistrictTerm.DistrictTermID
            inner join District with (nolock) on
            District.DistrictID=DistrictTerm.DistrictID
            inner join Student with (nolock) on
            TestResult.StudentID=Student.StudentID
            inner join School with (nolock) on
            TestResult.SchoolID=School.SchoolID
            inner join Class with (nolock) on
            TestResult.ClassID=Class.ClassID
            inner join [User] with (nolock) on
            TestResult.UserID=[User].UserID
            inner join TestResultSubScore with (nolock) on
            TestResultSubScore.TestResultScoreID=trs.TestResultScoreID

        where District.DistrictID=@districtid
        AND VirtualTest.Name LIKE '20%-20%PARCC%'
        AND VirtualTest.achievementlevelsettingid=217""", database)

    cluster['one'] = cluster["Score"] == 1
    grouped = cluster.groupby(by=["TestName", "School", "ClusterName"])
    for key in grouped.groups.keys():
        (t, s, c) = key[0][:], key[1][:], key[2][:]
        num_col = "Score" if "Scale Score" in c else "one"
        group = grouped.get_group(key)
        num = group[num_col].sum()
        div = group["Score"].count()

        cluster.loc[(cluster.TestName == t) &
                    (cluster.School == s) &
                    (cluster.ClusterName == c),
                    "NUM"] = num
        cluster.loc[(cluster.TestName == t) &
                    (cluster.School == s) &
                    (cluster.ClusterName == c),
                    "DIV"] = div
    cluster = cluster[[
        'TestName', 'School', 'ClassID', 'Class Name', 'ClusterName',
        'NUM', 'DIV']]
    cluster.to_excel(file, sheet_name='Cluster', index=False)
    del cluster
    # Gender
    boxg = pd.read_sql(
        """declare @district int
        set @district={}
        select s.StudentID, g.name as Gender
        from student s With (nolock)
        join gender g With (nolock) on g.GenderID=s.GenderID
        where s.DistrictID=@district
        """.format(districtID), database)
    boxg.to_excel(file, sheet_name='Gender', index=False)
    # Race
    boxg = pd.read_sql(
        """declare @district int
        set @district={}
        select s.StudentID, r.name as Race from student s With (nolock)
        join race r With (nolock) on r.raceid=s.raceid
        where s.districtid=@district
        """.format(districtID), database)
    boxg.to_excel(file, sheet_name='Race', index=False)
    # Program
    boxg = pd.read_sql(
        """declare @district int
        set @district={}
        select sp.StudentID, p.name as Program
        from studentprogram sp With (nolock)
        join program p With (nolock) on
        p.programid=sp.programid
        where p.districtid=@district
        """.format(districtID), database)
    boxg.to_excel(file, sheet_name='Program', index=False)
    del boxg

    # do all .to_excel calls before calling this
    file.save()
    # We're done! Send the user a message letting them know this.
    return (district_name
            + " PARCC Extract created and saved sucessfully."
            + "\nLocation: "
            + file_name)
