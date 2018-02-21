"""Extract Benchmark Status data to send to client."""
import mal_data as mal
import datetime
import pandas as pd
import os
import os.path
_students = """ declare  @districtid int = '{}'

SELECT  DISTINCT

School.Name AS [School], Student.LastName,
Student.FirstName, Student.MiddleName, Student.Code, Grade.Name as [Grade]
FROM DistrictTerm WITH (NOLOCK)
INNER JOIN Class WITH (NOLOCK) ON
DistrictTerm.DistrictTermID = Class.DistrictTermID
INNER JOIN ClassUser WITH
(NOLOCK) ON Class.ClassID = ClassUser.ClassID
INNER JOIN School WITH (NOLOCK)
ON Class.SchoolID = School.SchoolID
INNER JOIN [User] WITH (NOLOCK) ON
ClassUser.UserID = [User].UserID
INNER JOIN ClassUserLOE WITH (NOLOCK) ON
ClassUser.ClassUserLOEID = ClassUserLOE.ClassUserLOEID
INNER JOIN ClassStudent
WITH (NOLOCK) ON Class.ClassID = ClassStudent.ClassID
INNER JOIN Student WITH
(NOLOCK) ON ClassStudent.StudentID = Student.StudentID
LEFT JOIN Grade WITH
(NOLOCK) ON Student.CurrentGradeID = Grade.GradeID
WHERE STUDENT.DISTRICTID = @districtid AND DistrictTerm.Active = 1
order by School, [Grade], LastName, FirstName, Student.Code
"""

_benchmarks = """declare  @districtid int = '{}', @resultdate datetime =
'{}-08-01'
select tr.UpdatedDate as [ResultDate], dt.name as DistrictTerm,
vt.name as TestName, sub.name as Subject, gr.name as Grade,
sch.Name as School, u.UserID, u.code as TeacherCode,
u.NameFirst as TeacherFirstName, u.NameLast as TeacherLastName,
c.ClassID, c.Name as ClassName, s.StudentID, s.code as StudentCode,
s.FirstName as StudentFirstName, s.LastName as StudentLastName,
trs.ScoreRaw as TotalPointsEarned, trs.PointsPossible as TotalPointsPossible
from testresultscore trs With (nolock)
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
where dt.DistrictID=@districtid and dt.Active = 1
and vt.Name like '%LinkIt%form%' and not
(vt.Name like '%retake%' or vt.Name like '%LinkIt%form%CR%')
"""


def find_score(code, subj, df, form_letter):
    """Check df for a score matching subj, form_letter and row."""
    expr = ("StudentCode == '{}' "
            "and Subject == '{}' "
            "and Form == '{}'").format(code, subj, form_letter)
    query = df.query(expr)
    if not query.empty:
        return True
    return False


def main(districtID, form):
    """Extract Benchmark Status data and save as spreadsheet."""
    global _students
    global _benchmarks
    connection = mal.setup_SQL()
    df_s = pd.read_sql(_students.format(districtID), connection)
    sql_b = _benchmarks.format(districtID,
                               datetime.date.today().year - 1)
    df_b = pd.read_sql(sql_b, connection)
    df_b["Form"] = df_b.apply(
        lambda row: row.TestName[row.TestName.find("Form") + 5] if
        ("Form" in row.TestName) else " ", axis=1)
    df_s["ELA"] = df_s.apply(
        lambda row: "Yes" if find_score(
            row["Code"],
            "Language Arts", df_b, form) else "No", axis=1)
    df_s["Math"] = df_s.apply(
        lambda row: "Yes" if find_score(
            row["Code"],
            "Math", df_b, form) else "No", axis=1)

    file_name = "Benchmark Status\\{} Form {} Benchmark Status {}".format(
        mal.get_district_name(districtID),
        form,
        str(datetime.date.today()))

    if not os.path.exists("Benchmark Status"):
        os.makedirs("Benchmark Status")
    file_path = mal.path_to(file_name)
    wb, excel = mal.get_excel(file_path)
    ws_s = mal.get_sheet(wb, "Students")
    mal.df_to_excel(df_s, ws_s)
    # ws_b = mal.get_sheet(wb, "Benchmarks")
    # mal.df_to_excel(df_b, ws_b)
    for ws in wb.Worksheets:
        ws.Cells.EntireColumn.AutoFit()
    mal.excel_save_quit(wb, excel, file_name)
    return (f"File created in {file_path}.")


if __name__ == "__main__":
    main(3757, "B")
