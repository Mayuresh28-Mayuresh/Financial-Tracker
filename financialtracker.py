import csv
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font


def add_expense():
    date = input("Enter date (YYYY-MM-DD): ")
    category = input("Enter category: ")
    item = input("Enter item: ")
    amount = float(input("Enter amount: "))

    with open('expenses.csv', mode='a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow([date, category, item, amount])


def generate_report():
    total_expenses = 0
    category_expenses = {}

    with open('expenses.csv', mode='r') as file:
        reader = csv.reader(file)
        for row in reader:
            total_expenses += float(row[3])
            category = row[1]
            amount = float(row[3])
            category_expenses[category] = category_expenses.get(category, 0) + amount

    print("Total Expenses: $", total_expenses)
    print("\nCategory Wise Expenses:")
    for category, amount in category_expenses.items():
        print(category, ": $", amount)

    # Generate pie chart for category-wise expenses
    plt.figure(figsize=(8, 6))
    plt.pie(list(category_expenses.values()), labels=list(category_expenses.keys()), autopct='%1.1f%%')
    plt.title('Category Wise Expenses')
    plt.show()

    # Store data in Excel
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Expense Summary"
    sheet.append(["Category", "Amount"])
    for category, amount in category_expenses.items():
        sheet.append([category, amount])
    sheet.append(["Total Expenses", total_expenses])

    # Format headers
    for cell in sheet["1:1"]:
        cell.font = Font(bold=True)

    # Save Excel file
    workbook.save('expense_summary.xlsx')
    print("Expense summary saved to 'expense_summary.xlsx'")


def main():
    while True:
        print("\nExpense Tracker")
        print("1. Add Expense")
        print("2. Generate Report")
        print("3. Exit")
        choice = input("Enter your choice: ")

        if choice == '1':
            add_expense()
        elif choice == '2':
            generate_report()
        elif choice == '3':
            print("Exiting...")
            break
        else:
            print("Invalid choice. Please try again.")


if __name__ == "__main__":
    main()
