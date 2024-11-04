# importing important libraries
import pandas as pd 

# creating basic variables for the employee class, as well as placeholder for csv file
firstname = ""
lastname = ""
fullname = ""
CORS = ""
tasks = []
datafile = 'outbound tasks LOCAL.xlsx'



# creating of class of Employee, which contains [name, system name, and tasks they can do]
class Employee:
    def __init__(self, firstname, lastname, CORS, tasks) -> None:
        self.firstname = firstname
        self.lastname = lastname
        self.fullname = f"{firstname} {lastname}"
        self.CORS = CORS #system name
        self.tasks = tasks

    def __repr__(self) -> str:
        #explains how to return the class and its variables in a string
        return f"Employee: {self.firstname} {self.lastname}, CORS: {self.CORS}, Tasks: {', '.join(self.tasks)}"

def load_spreadsheet(datafile):
    # reads excel file and makes them as lists of strings and what not
    database = pd.read_excel(datafile)

    employee_list = []

    #iterates through each row header and pulls the strings below in each row 
    for index, row in database.iterrows():
        lastname = row['LastName']
        firstname = row['FirstName']
        CORS = row['CORSName']
        tasks = []
        
        #goes through the rows containing tasks and appends them to a list
        for col in database.columns[3:]:
            if pd.notna(row[col]):
                tasks.append(row[col])
        
        employee_list.append(Employee(firstname, lastname, CORS, tasks))
        
    return employee_list

def find_employees(employee_list, firstname=None, lastname=None, fullname=None):
    results = []

    for employee in employee_list:
        if fullname and employee.fullname.lower() == fullname.lower():
            results.append(employee)
        elif firstname and employee.firstname.lower() == firstname.lower():
            results.append(employee)
        elif lastname and employee.lastname.lower() == lastname.lower():
            results.append(employee)

    return results

def create_employee(datafile):

    # Load the existing Excel file into a DataFrame
    df = pd.read_excel(datafile)

    firstname = input("Enter the first name:   ").strip()
    lastname = input("Enter the last name:   ").strip()
    CORS = input("Enter the CORS (system name):   ").strip()

    if firstname in df and lastname in df:
        print("This employee is already in the database, start over")
        
    
    exit()
    
    

    # Create a new row to append to the DataFrame
    new_row = {
        'FirstName': firstname,
        'LastName': lastname,
        'CORSName': CORS,
    }

    task_columns = df.columns[3:]

    tasks = []
    while True:
        task = input("Enter a task (or press enter to finish): ").strip().upper()
        if not task:
            break
        if task in task_columns:
            tasks.append(task)
        else:
            print("Task is not valid, make sure you are spelling correctly, and maybe peep at the file to double check")

    # Add the tasks to the appropriate columns (assuming task columns start at the 4th column)
    for task in tasks:
        new_row[task] = task

    new_row_df = pd.DataFrame([new_row])

    # Append the new employee to the DataFrame
    df = pd.concat([df, new_row_df], ignore_index=True)

    # Save the updated DataFrame back to the Excel file
    df.to_excel(datafile, index=False)
    
    print(f"Employee {firstname} {lastname} has been added to the file.")

   

def delete_employee(datafile):
    df = pd.read_excel(datafile)

    # Search for the employee to delete
    firstname = input("Enter the first name of the employee to delete: ").strip()
    lastname = input("Enter the last name of the employee to delete: ").strip()

    # Find the index of the employee in the DataFrame
    index_to_delete = df[(df['FirstName'].str.lower() == firstname.lower()) & (df['LastName'].str.lower() == lastname.lower())].index

    # If the employee exists, delete the row
    if not index_to_delete.empty:
        df = df.drop(index_to_delete)
        df.to_excel(datafile, index=False)
        print(f"Employee {firstname} {lastname} has been deleted from the file.")
    else:
        print("Employee not found.")

    

def update_employee(datafile):
    # Load the existing Excel file into a DataFrame
    df = pd.read_excel(datafile)
    
    # Prompt the user for the employee's current details
    firstname = input("Enter the first name of the employee to update: ").strip().upper()
    lastname = input("Enter the last name of the employee to update: ").strip().upper()
    
    # Search for the employee in the DataFrame
    employee_row = df[(df['FirstName'] == firstname) & (df['LastName'] == lastname)]
    
    if employee_row.empty:
        print("Employee not found.")
        return
    
    # Get the index of the employee to update their details
    employee_index = employee_row.index[0]
    
    # Prompt for new details, allowing the user to skip updating a field
    new_firstname = input(f"Enter new first name (or press Enter to keep '{firstname}'): ").strip().upper()
    new_lastname = input(f"Enter new last name (or press Enter to keep '{lastname}'): ").strip().upper()
    new_CORS = input(f"Enter new CORS (or press Enter to keep '{df.at[employee_index, 'CORSName']}'): ").strip().upper()

    # Update the employee's name and CORS if the user provided new input
    if new_firstname:
        df.at[employee_index, 'FirstName'] = new_firstname
    if new_lastname:
        df.at[employee_index, 'LastName'] = new_lastname
    if new_CORS:
        df.at[employee_index, 'CORSName'] = new_CORS

    # Update tasks
    task_columns = df.columns[3:]  # Task columns start from the 4th column onward
    print("Current tasks: ")
    print(employee_row[task_columns].dropna(axis=1, how='all').to_string(index=False))  # Display only non-empty task columns

    # Ask if the user wants to update tasks
    update_tasks = input("Do you want to update tasks? (yes/no): ").strip().lower()
    
    if update_tasks == 'yes':
        # Clear existing tasks
        for task_col in task_columns:
            df.at[employee_index, task_col] = None
        
        # Prompt for new tasks
        tasks = []
        while True:
            task = input("Enter a new task (or press Enter to finish): ").strip().upper()
            if not task:
                break
            if task in task_columns:
                tasks.append(task)
            else:
                print(f"Task '{task}' is not a valid task column. Please enter a valid task.")
        
        # Assign new tasks
        for task in tasks:
            df.at[employee_index, task] = task

    # Save the updated DataFrame back to the Excel file
    df.to_excel(datafile, index=False)
    
    print(f"Employee {firstname} {lastname} has been updated.")

def search_by_task(datafile):
    df = pd.read_excel(datafile)
    taskInput = input("Enter the task to search for: ").strip().upper()
    
    task_columns = df.columns[3:]
    if taskInput in task_columns:
        print("here")
        results = df[df[taskInput].str.lower() == taskInput.lower()]
        

        if not results.empty:
            print("Employees assigned to this task:")
            for index, emp in results.iterrows():
                print(emp["FirstName"] + ' ' + emp["LastName"])
        else:
            print("No employees found for this task.")

    else:
        print(f"Task '{taskInput}' was not found")

    



if __name__ == "__main__":

    print("Employee Datasheet Loaded")
    print("Hello!, I am Adi, a 2024 Operations Intern. This program was designed to help you with task delegation, hope you enjoy!")


    employee_list = load_spreadsheet(datafile)
    
    while True:
        action = input("Do you want to SEARCH an employee, ADD an employee, UPDATE an employee, DELETE an employee, or QUIT? You can also search by job: ").strip().lower()

        if action == "search": 
        # prompts user with input of name
            search_input = input("Do you want to search by 'FIRSTNAME', 'LASTNAME', OR 'FULLNAME'?   ").strip().lower()

            if search_input == 'fullname':
                fullname = input("Enter the full name: ").strip()
                results = find_employees(employee_list, fullname=fullname)
            elif search_input == 'firstname':
                firstname = input("Enter the first name: ").strip()
                results = find_employees(employee_list, firstname=firstname)
            elif search_input == 'lastname':
                lastname = input("Enter the last name: ").strip()
                results = find_employees(employee_list, lastname=lastname)
            else:
                print("Invalid option. Please choose 'firstname', 'lastname', or 'fullname'.")
                results = []

            # Display the results
            if results:
                print(f"Found {len(results)} employee(s):")
                for employee in results:
                    print(employee)
            else:
                print("No employees found with the provided name.")

        elif action == "add":
            create_employee(datafile)

        elif action == "update": 
            update_employee(datafile)

        elif action == "delete":
            delete_employee(datafile)

        elif action == "quit":
            exit()

        elif action == "job":
            search_by_task(datafile)

        else:
            print("Invalid option, please try again")



        




                


    
    


    

    





