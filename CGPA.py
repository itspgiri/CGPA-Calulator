# We use pandas for our excel sheet
import pandas as pd
from pandas import ExcelWriter

# Import the excel file that contains all the details
data = pd.read_excel("ParticipantList.xlsx")

code = data['Code'].to_list()
course = data['Course'].to_list()
course_type = data['Type'].to_list()
credit = data['Credits'].to_list()
reg = data['No'].to_list()
name = data['Name'].to_list()
score = data['Score'].to_list()
school = data['School'].to_list()
branch = data['Branch'].to_list()

credit1 = data['Credits'].to_list()
reg1 = data['No'].to_list()
score1 = data['Score'].to_list()

final_credit = []
final_reg = []
final_name = []
final_score = []
final_school = []
final_branch = []
final_cgpa = []


def calls(reg11):
    t = 0
    c = 0
    for reg_i, score_i, credit_i in zip(reg1, score1, credit1):
        if reg_i == reg11:
            t = t + score_i
            c = c + credit_i
    return t, c


for code, course, course_type, credit, reg, name, score, school, branch in zip(code, course, course_type, credit, reg,
                                                                               name, score, school, branch):
    if reg not in final_reg:
        totalscore, totalcredit = calls(reg)
        if totalscore != 0 and totalcredit != 0:
            cgpa = totalscore / totalcredit
        else:
            cgpa = 0
        final_credit.append(totalcredit)
        final_reg.append(reg)
        final_name.append(name)
        final_score.append(totalscore)
        final_school.append(school)
        final_branch.append(branch)
        final_cgpa.append(cgpa)

df = pd.DataFrame({'Total Credit': final_credit,
                   'Reg No': final_reg,
                   'Name': final_name,
                   'Total Score': final_score,
                   'School': final_school,
                   'Branch': final_branch,
                   'CGPA': final_cgpa})

writer = ExcelWriter('CGPA.xlsx')
df.to_excel(writer, 'Sheet1', index=False)
writer.save()
print("Sheet Created")
