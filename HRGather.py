# Imports
import pandas as pd
import datetime
import nameparser
from nameparser import HumanName

# New Hires Path
path = ("path_to_HR_New_Hire.xlsx")

# distinguishedNames
CSG = ("OU=Users,OU=Dept,OU=Comp,DC=DC,DC=DC")
HCG = ("OU=Users,OU=Dept,OU=Comp,DC=DC,DC=DC")
LCH = ("OU=Users,OU=Dept,OU=Comp,DC=DC,DC=DC")
ICS = ("OU=Users,OU=Dept,OU=Comp,DC=DC,DC=DC")
LCA = ("OU=Users,OU=Dept,OU=Comp,DC=DC,DC=DC")
LSC = ("OU=Users,OU=Dept,OU=Comp,DC=DC,DC=DC")
NBG = ("OU=Users,OU=Dept,OU=Comp,DC=DC,DC=DC")
ESH = ("OU=Users,OU=Dept,OU=Comp,DC=DC,DC=DC")


# Loads in New hires sheet as Pandas Dataframe
pd.set_option('display.max_rows', 500, 'display.max_columns', 0)
df = pd.read_excel(path).dropna(how='all')

# Filters Dataframe to only include Users who are not yet added
df = df[df["AD/Rights Given"] != "ok"]

# Pulls Name column from New hire list, extracts only string values fromt it and uses a library to seperate first from last names in seperate lists.
items = []
results = df.iloc[:,1]
for result in results:
        items.append(result)
items = [n for n in items if isinstance(n, str)]



first_names = []   
items = [n for n in items if isinstance(n, str)]
for item in items:
        name = HumanName(item)
        first_names.append(name.first)
last_names = []   
for item in items:
        name = HumanName(item)
        last_names.append(name.last)


# Pulls titles section from New hire list and extracts only string values from it into a new list.
TITLE_LIST = []
titles = df.iloc[:,3]
for title in titles:
        TITLE_LIST.append(title)
TITLE_LIST = [n for n in TITLE_LIST if isinstance(n, str)]


# Pulls SBU's (Business Units) from new hire sheet and extracts only string values.
DEPT_LIST = []
depts = df.iloc[:,5]
for dept in depts:
        SBU_LIST.append(sbu)
DEPT_LIST = [n for n in DEPT_LIST if isinstance(n, str)]

# Gets email addresses from names
EMAIL_LIST = []
for name in items:
        name = name.replace(" ", ".").lower()
        name = (f"{name}@domain.com")      
        EMAIL_LIST.append(name)

# Iterates through SBU_LIST and appends the cooresponding DN attribute to to OU_LIST
OU_LIST = []
for dept in DEPT_LIST:
        if 'string' in dept:
                OU_LIST.append(DEPT1)
        elif 'string' in dept:
                OU_LIST.append(DEPT2)
        elif 'string' in dept:
                OU_LIST.append(DEPT3)
        elif 'string' in dept:
                OU_LIST.append(DEPT4)
        elif 'string' in dept:
                OU_LIST.append(DEPT5)
        elif 'string' in dept:
                OU_LIST.append(DEPT6)
        elif 'string' in dept:
                OU_LIST.append(DEPT7)
        elif 'string' in dept:
                OU_LIST.append(DEPT8)
        elif 'na' in dept:
                OU_LIST.append("pass")
        


# Puts each list into a dictionary object
list_dict = {
        'First': first_names,
        'Last': last_names,
        'Title': TITLE_LIST,
        'SBU': DEPT_LIST,
        'EMAIL': EMAIL_LIST,
        'DN': OU_LIST}

# Converts Python dictionary to a Pandas Dataframe
df_new = pd.DataFrame.from_dict(list_dict, orient="columns")
# Writes dataframe to .csv to then be parsed by Powershell
try:
        df_new.to_csv('filepath\\new_users.csv', index=False)
        print("Printed to CSV")

except:
        print("Could not print to CSV!")
        
