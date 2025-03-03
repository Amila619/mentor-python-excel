import pandas as pd

mentee_excel_file = 'mentee_votes.xlsx'
mentor_excel_file = 'mentor.xlsx'

# Opening necessary files
in_df = pd.read_excel(mentee_excel_file)
op_df = pd.read_excel(mentor_excel_file)

# Generating required data structures
mentors = {}
left_out_mentors = [];

for index, row in op_df.iterrows():
    mentors[row['Name']] = []

# Assign Mentors
for index, row in in_df.iterrows():
    Mentor_op_1 = row['vote 1']
    Mentor_op_2 = row['vote 2']
    Mentor_op_3 = row['vote 3']
    Mentee_ = row['tg number']

    assigned = False
    for mentor in [Mentor_op_1, Mentor_op_2, Mentor_op_3]:
        if mentor in mentors and len(mentors[mentor]) < 2 and Mentee_ not in mentors[mentor]:
            mentors[mentor].append(Mentee_)
            assigned = True
            break 
    if not assigned:
        left_out_mentors.append(Mentee_)

# Assign Left Out Mentees
for mentee in left_out_mentors:
    for k in mentors.keys():
        if len(mentors[k]) < 2:
            mentors[k].append(mentee)
        
# Generating mentees assigned to mentor file
df = pd.DataFrame(columns=["Mentor", "Mentee1", "Mentee2"])
file_path = "./mentor_mentee_assigned.xlsx"

for k, v in mentors.items():
    df.loc[len(df)] = [k, v[0], v[1]]

df.to_excel(file_path, index=False)