# importing important libraries
import pandas as pd

# reading spreadsheet containing employee information
df = pd.read_csv('colonyproject.csv')
print(df.to_sql)

# creating basic variables for the employee class
firstname = ""
lastname = ""
CORS = ""
tasks = []

# creating of class of Employee, which contains [name, system name, and tasks they can do]
class Employee:
    def __init__(self, lastname, CORS, tasks) -> None:
        self.lastname = lastname
        self.CORS = CORS
        self.tasks = tasks




