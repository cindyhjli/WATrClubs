import pandas as pd

curr = pd.read_excel('members.xlsx')
form = pd.read_excel('WATrClubs.xlsx')

women = []
men = []
either = []

form = form.sort_values(['Faculty (optional but may be used for matching)', 'Year (optional but may be used for matching)'])

for x in range(len(form.index)):
    name = form.iloc[x]['Name']
    email = form.iloc[x]['Email (school email)']
    if form.iloc[x]['Do you have a preference for your group?'] == "Same gender":
        if form.iloc[x]['Gender (optional but will be used for the next question)'] == "Women+":
            women.append([name, email, 'New Member'])
        else:
            men.append([name, email, 'New Member'])
    else:
        either.append([name, email, 'New Member'])

members = []
groups = 0

if len(women) < 3:
    groups += 1
else:
    groups += (len(women) // 3)

if len(men) < 3:
    groups += 1
else:
    groups += (len(men) // 3)

if len(either) < 3:
    groups += 1
else:
    groups += (len(either) // 3)

for x in range(len(curr.index)):
    name = curr.iloc[x]["Name"]
    email = curr.iloc[x]['Email']
    members.append([name, email, 'Current Member'])

if len(members) < groups:
    diff = groups - len(members)
    for x in range(diff):
        members.append(members[x])

count = 1
while len(women) > 6:
    new_df = pd.DataFrame(columns=['Name', 'Email', 'Role'])
    new_df.loc[0] = members[count-1]
    for x in range(3):
        new_df.loc[x+1] = women[0]
        women.pop(0)
    with pd.ExcelWriter('WATrClubs.xlsx', mode='a', engine="openpyxl", if_sheet_exists="replace") as writer:
        new_df.to_excel(writer, sheet_name='Group '+str(count), index=False)
    count += 1

new_df = pd.DataFrame(columns=['Name', 'Email', 'Role'])
new_df.loc[0] = members[count-1]
for x in range(len(women)):
    new_df.loc[x+1] = women[x]

with pd.ExcelWriter('WATrClubs.xlsx', mode='a', engine="openpyxl", if_sheet_exists="replace") as writer:
    new_df.to_excel(writer, sheet_name='Group '+str(count), index=False)
count += 1

while len(men) > 6:
    new_df = pd.DataFrame(columns=['Name', 'Email', 'Role'])
    new_df.loc[0] = members[count-1]
    for x in range(3):
        new_df.loc[x+1] = men[0]
        men.pop(0)
    
    with pd.ExcelWriter('WATrClubs.xlsx', mode='a', engine="openpyxl", if_sheet_exists="replace") as writer:
        new_df.to_excel(writer, sheet_name='Group '+str(count), index=False)
    count += 1

new_df = pd.DataFrame(columns=['Name', 'Email', 'Role'])
new_df.loc[0] = members[count-1]
for x in range(len(men)):
    new_df.loc[x+1] = men[x]

with pd.ExcelWriter('WATrClubs.xlsx', mode='a', engine="openpyxl", if_sheet_exists="replace") as writer:
    new_df.to_excel(writer, sheet_name='Group '+str(count), index=False)
count += 1

while len(either) > 6:
    new_df = pd.DataFrame(columns=['Name', 'Email', 'Role'])
    new_df.loc[0] = members[count-1]
    for x in range(3):
        new_df.loc[x+1] = either[0]
        either.pop(0)

    with pd.ExcelWriter('WATrClubs.xlsx', mode='a', engine="openpyxl", if_sheet_exists="replace") as writer:
        new_df.to_excel(writer, sheet_name='Group '+str(count), index=False)
    count += 1

new_df = pd.DataFrame(columns=['Name', 'Email', 'Role'])
new_df.loc[0] = members[count-1]
for x in range(len(either)):
    new_df.loc[x+1] = either[x]

with pd.ExcelWriter('WATrClubs.xlsx', mode='a', engine="openpyxl", if_sheet_exists="replace") as writer:
    new_df.to_excel(writer, sheet_name='Group '+str(count), index=False)
count += 1
