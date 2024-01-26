import pandas as pd
from tabulate import tabulate

def analyze_employee_schedule(file_path):
    try:
        employee_schedule_data_frame = pd.read_excel(file_path)
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")
        return
    except pd.errors.EmptyDataError:
        print(f"Error: File '{file_path}' is empty.")
        return
    except pd.errors.ParserError:
        print(f"Error: Unable to parse the CSV file '{file_path}'.")
        return
    employee_schedule_data_frame['Time'] = pd.to_datetime(employee_schedule_data_frame['Time'])
    employee_schedule_data_frame['Time'] = pd.to_datetime(employee_schedule_data_frame['Time'])
    employee_schedule_data_frame['Time Out'] = pd.to_datetime(employee_schedule_data_frame['Time Out'])
    employee_schedule_data_frame.sort_values(by=['Employee Name', 'Time'], inplace=True)
    employee_schedule_data_frame['DaysDiff'] = employee_schedule_data_frame.groupby('Employee Name')['Time'].diff().dt.days
    employee_schedule_data_frame['ShiftDuration'] = (employee_schedule_data_frame['Time'].shift(-1) - employee_schedule_data_frame['Time Out']).dt.total_seconds() / 3600
    consecutive_days_result = employee_schedule_data_frame[employee_schedule_data_frame['DaysDiff'] == 1].groupby('Employee Name').filter(lambda x: len(x) >= 6)
    short_time_between_shifts_result = employee_schedule_data_frame[(employee_schedule_data_frame['ShiftDuration'] < 10) & (employee_schedule_data_frame['ShiftDuration'] > 1)]
    long_single_shift_result = employee_schedule_data_frame[employee_schedule_data_frame['ShiftDuration'] > 14]

    print("Employees who have worked for 7 consecutive days:")
    print(tabulate(consecutive_days_result[['Position ID','Employee Name', 'Position Status']].drop_duplicates(), headers='keys', tablefmt='pretty'))
    print("\nEmployees who have less than 10 hours between shifts but greater than 1 hour:")
    print(tabulate(short_time_between_shifts_result[['Position ID','Employee Name', 'Position Status']], headers='keys', tablefmt='pretty'))
    print("\nEmployees who have worked for more than 14 hours in a single shift:")
    print(tabulate(long_single_shift_result[['Position ID','Employee Name', 'Position Status']], headers='keys', tablefmt='pretty'))

    with open('output.txt', 'w') as output_file:
        output_file.write("Employees who have worked for 7 consecutive days:\n")
        output_file.write(tabulate(consecutive_days_result[['Position ID','Employee Name', 'Position Status']].drop_duplicates(), headers='keys', tablefmt='pretty') + '\n\n')

        output_file.write("Employees who have less than 10 hours between shifts but greater than 1 hour:\n")
        output_file.write(tabulate(short_time_between_shifts_result[['Position ID','Employee Name', 'Position Status']], headers='keys', tablefmt='pretty') + '\n\n')

        output_file.write("Employees who have worked for more than 14 hours in a single shift:\n")
        output_file.write(tabulate(long_single_shift_result[['Position ID','Employee Name', 'Position Status']], headers='keys', tablefmt='pretty') + '\n\n')

if __name__ == "__main__":
    file_path = 'C:\\Users\\swast\\OneDrive\\Desktop\\DataAnalyzerProject\\Assignment_Timecard.xlsx'
    analyze_employee_schedule(file_path)







