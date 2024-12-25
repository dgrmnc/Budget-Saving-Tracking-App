from datetime import datetime
import time
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns 

# budget-saving app
class Budget:

    date_format = "%d-%m-%Y"
    excel_file = "budget_report.xlsx"

    def __init__(self, budget=0):
        if budget < 0:
            raise ValueError("can not be negative.")
        self.original_budget = budget
        self.remaining_budget = budget
        self.income_type = []
        self.income = []
        self.expenses_type = []
        self.expenses = []
        self.saving_amount = []
        self.current_saving = 0
    
    # ask user their budget
    def get_budget(self):
        while True:
            try:
                self.original_budget = float(input("Please enter your current budget: "))
                if self.original_budget < 0:
                    raise ValueError("Budget cannot be negative.")
                self.remaining_budget = self.original_budget
                print(f"The budget you just entered is: {self.original_budget}")
                break
            except ValueError:
                print("Please enter a valid number.")
    
    # ask user their budget type
    def get_income_type(self):
        while True:
            try:
                income_type = input("Please enter your income type: ").strip().lower().capitalize()
                if not isinstance(income_type, str): #or not income_type.isalpha():
                    raise ValueError("Income type must be a non-empty string containing only letters.")
                if len(income_type) > 45:
                    raise ValueError("Income type must be within 25 characters.")
                self.income_type.append(income_type) 
                print(f"Your income type is: {income_type}")
                break
            except ValueError as e:
                print(e)
    # ask expenses if there is any
    def get_expenses(self):
        while True:
            try:
                no_expenses = input("Would you like to skeep adding expenses?(y/n): ").lower().strip()
                if no_expenses in ['y','n']:
                    self.ask_for_savings()
                    self.extract_to_table()
                    print("Skipped adding expenses. Income has been saved to the Excel file.")
                    return
                elif no_expenses not in ['y','n']:
                    print("Please enter y or no in order to skip this process.")
                    continue
                break
            except Exception as e:
                print(f"An error occued: {e}")
        # for multiple expenses
        while True:
            try:
                    num_of_expenses = int(input("How many expenses would you like to add?: "))
                    if num_of_expenses < 0:
                        raise ValueError("Expenses number must be positive")
                    break
            except ValueError:
                    print(ValueError)

        for i in range(num_of_expenses):
                    while True:
                        try:
                            expense = float(input(f"Please enter {i + 1} expenses: "))
                            if expense < 0:
                                raise ValueError("Can not be negative value")        
                            self.expenses.append(expense)
                            self.remaining_budget -= expense
                            break
                        except ValueError:
                            print(ValueError)
                    while True:
                        try:
                            expense_type = input(f"Please enter {i + 1} expense type here: ").strip().lower().capitalize()
                            if not isinstance(expense_type,str): #or not expense_type.isalpha():
                                raise ValueError("must be string value and special characters are not allowed.")
                            self.expenses_type.append(expense_type)
                            print(f"Expenses type is : {expense_type}")
                            break
                        except ValueError as e:
                            print(e)               
        self.ask_for_savings()
        # here goes to calc the current budget and write it down to excel
        self.extract_to_table()

    # ask user if they save any money
    def ask_for_savings(self):
        while True:
            save_money = input("Would you like to add money into your saving account for this month?(yes / no): ").strip().lower()
            if save_money in ['yes', 'y']:
                while True:
                    try:
                        saving_amount = float(input("How much money would you like to put?: "))
                        if saving_amount < 0:
                            raise ValueError("Can not be negative")
                        if saving_amount > self.remaining_budget:
                            raise ValueError("You cannot put aside more than your remaining budget.")
                        self.remaining_budget -= saving_amount
                        self.current_saving += saving_amount
                        print(f"You have saved: {saving_amount}")
                        break
                    except ValueError as e:
                        print(e)
                break
            elif save_money in ['no', 'n']:
                print("No saving for this month. No problem, there is always tomorrow!")
                break
            else:
                print("Please enter 'yes' or 'no'.")

    # calculate the total money after expenses
    def calculate_current_money(self):
        total_expenses = sum(self.expenses)
        current_money = self.original_budget - total_expenses
        print(f"Your remaining budget after calculating expenses is : {current_money:.2f}")
        return current_money
   
    
    # to extract the input in to the excel
    def extract_to_table(self):
   
        current_date = datetime.now()
        #current_month = current_date.month
        
        # if there is no expenses but income type
        if not self.expenses:
            data = {
                "Income": [self.original_budget],
                "Income Type": [self.income_type[0] if self.income_type else "N/A"],
                "Expenses": ["No Expenses"],
                "Expenses Type": ["No Expenses Type"],
                "Remaining Budget": [self.original_budget],
                "Insertion_Date": [current_date.strftime("%d-%m-%Y")],
                "Insertion_Month": [current_date.strftime("%B")],
                "Saving": [self.current_saving],
                "Rem Budget(After Saving)": [self.remaining_budget]
            }
        # if there is both expenses and income 
        else:
            data = {
                "Income": [self.original_budget] * len(self.expenses),
                "Income Type": [self.income_type[0]] * len(self.expenses),
                "Expenses": self.expenses,
                "Expenses Type": self.expenses_type,
                "Remaining Budget": [self.original_budget - sum(self.expenses[:i+1]) for i in range(len(self.expenses))],
                "Insertion_Date": [datetime.now().strftime("%d-%m-%Y")] * len(self.expenses),
                "Insertion_Month": [current_date.strftime("%B")] * len(self.expenses),
                "Saving": [self.current_saving] * len(self.expenses),
                "Rem Budget(After Saving)": [self.remaining_budget] * len(self.expenses)
            }

        try:
            existing_data = pd.read_excel(Budget.excel_file)
            df = pd.DataFrame(data)
            df = pd.concat([existing_data, df], ignore_index=True)  
        except FileNotFoundError:
            df = pd.DataFrame(data)

        df.to_excel(Budget.excel_file, index=False)
        print("Data has been successfully exported to 'budget_report.xlsx'.")

    @classmethod
    def plot_expenses(cls, start_date, end_date):
        try:
            df = pd.read_excel(cls.excel_file)
            df['Insertion_Date'] = pd.to_datetime(df['Insertion_Date'], format=cls.date_format)
            start_date = datetime.strptime(start_date, cls.date_format)
            end_date = datetime.strptime(end_date, cls.date_format)

            filtered_df = df[(df['Insertion_Date'] >= start_date) & (df['Insertion_Date'] <= end_date)]

            if filtered_df.empty:
                print("No data available for the given date range.")
            else:
                print(f"Transactions from {start_date.strftime(Budget.date_format)} to {end_date.strftime(Budget.date_format)}")
                print(filtered_df.to_string(index=False, formatters={"Insertion_Date": lambda x : x.strftime(Budget.date_format)}))
            
            filtered_data = filtered_df.groupby(['Insertion_Month','Income Type'])['Saving'].sum().reset_index()
            plt.figure(figsize=(12,8))
            sns.barplot(data=filtered_data, x = 'Insertion_Month', hue = 'Income Type'
                         , y = 'Saving',palette= 'Set3',ci=None, alpha=0.8)
            plt.title("Monthly Saving Performance")
            plt.xlabel('Saving Insertion Month')
            plt.ylabel('Saving Total Amount')
            plt.xticks()
            plt.tight_layout()
            plt.show()

        except Exception as e:
            print(f"An error occurred: {e}")      

    #cleaning excel sheet
    def clear_excel_sheet(self,file_name, sheet_name):
            
        import openpyxl

        web = openpyxl.load_workbook(file_name)
        sheet = web[sheet_name]

        # the rows that you want to terminate
        total_rows = sheet.max_row

        for row in range(total_rows, 0, -1):
            sheet.delete_rows(row)

        web.save(file_name)
        print(f"{total_rows} rows has been deleted successfully.")

        # the way of using it:
        #clear_excel_sheet("budget_report.xlsx",'Sheet1')

    def cleaning_excel_sheet(self):
        print("Your data from your Sheet1 will be deleted")
        time.sleep(1)
        print("Processing")
        time.sleep(1.1)
        g1 = Budget()
        g1.clear_excel_sheet(Budget.excel_file, 'Sheet1')
        time.sleep(2)
        print("Deletion has been completed.")

    """
    def show_all_budgets(self):
        print(f"All entered budgets are: {Budget.income}")

    def show_all_income_types(self):
        print("All entered income types:", Budget.income_type) 
    
    def show_all_expenses(self):
        print(f"All entered budgets are: {Budget.expenses}") 

    def show_all_expenses_types(self):
        print(f"All entered budgets are: {Budget.expenses_type}")
    """


def main():
    g1 = Budget()
    while True:
        print("\n1. Add a new transaction: ")
        print("2. View transactions and plot within a given date range")
        print("3. Exit")
        print("4. Cleaning excel sheet")
        choice = input("Enter your choice( 1 - 4): ")

        if choice == "1":
            g1.get_budget()
            g1.get_income_type()
            g1.get_expenses()
            g1.calculate_current_money()
        elif choice == "2":
            start_date = input("Enter the start date(dd-mm-yyyy): ").strip()
            end_date = input("Enter the end date(dd-mm-yyyy): ").strip()
            Budget.plot_expenses(start_date, end_date)
        elif choice == "3":
            print("Exiting..")
            time.sleep(1)
            print("Exited.")
            break
        elif choice == '4':
            g1.cleaning_excel_sheet()
        else:
            print("Invalid choice.")

if __name__ == "__main__":
    main()

# for clearing the excel sheet

# def cleaning_excel_sheet():
#     print("Your data from your Sheet1 will be deleted")
#     time.sleep(1)
#     print("Processing")
#     time.sleep(1.1)
#     g1 = Budget()
#     g1.clear_excel_sheet(Budget.excel_file, 'Sheet1')
#     time.sleep(2)
#     print("Deletion has been completed.")

# if __name__ == "__main__":
#     cleaning_excel_sheet()  

 