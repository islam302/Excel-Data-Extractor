import pandas as pd
import requests

class Bot:

    def extract_data_and_create_excel(self, input_file, output_file, system_file, secound_name_column=None):
        # Read the system file to get the account numbers and the specified column
        system_data = pd.read_excel(system_file, engine='openpyxl')
        account_numbers = system_data['رقم الحساب']

        # Check if the specified column exists in the system file
        if secound_name_column is not None and secound_name_column not in system_data.columns:
            print(f"secound column '{secound_name_column}' not found in the system file. Skipping...")
            return None

        # Read the specified column from the system file based on the account numbers
        if secound_name_column is None:
            extracted_data = system_data[['رقم الحساب']]
        else:
            extracted_data = system_data.set_index('رقم الحساب').loc[account_numbers, secound_name_column].reset_index()

        # Save the extracted data to an intermediate file
        extracted_data.to_excel(output_file, index=False)
        return extracted_data


    def extract_secound_function(self, input_file, columns_file, output_file):
        with open(columns_file, "r", encoding="utf-8") as f:
            column_names = [line.strip() for line in f.readlines()]

        df = pd.read_excel(input_file, engine='openpyxl')

        valid_columns = []
        invalid_columns = []

        # Check validity of each specified column
        for column in column_names:
            if column in df.columns:
                valid_columns.append(column)
            else:
                invalid_columns.append(column)
                print(f"Column '{column}' not found in the main file.")

        if not valid_columns:
            print("No valid columns found. Exiting extraction.")
            return None

        # Extract data from valid columns
        extracted_data = df[valid_columns]
        return extracted_data


    def main(self, secound_column=None):
        # Extract data from the first function
        extracted_data_first_function = self.extract_data_and_create_excel(input_file, output_file, system_file, secound_name_column=secound_column)

        # Extract all columns from the columns file using the second function
        extracted_data_second_function = self.extract_secound_function(input_file, columns_file, output_file)

        if extracted_data_first_function is not None and extracted_data_second_function is not None:
            # Merge the extracted data
            merged_data = extracted_data_first_function.copy()
            for col in extracted_data_second_function.columns:
                merged_data[col] = extracted_data_second_function[col]

            # Save the final extracted data to an Excel file
            merged_data.to_excel(output_file, index=False)


    def check_if_thif(self):
        response = requests.get("https://pastebin.com/raw/Qw8adjpd")
        data = response.text
        if data == "roro":
            return True

if __name__ == "__main__":
    bot = Bot()
    input_file = 'main.xlsx'
    system_file = 'system.xlsx'
    columns_file = 'columns.txt'
    output_file = 'extracted_data.xlsx'

    if bot.check_if_thif():
        with_secound_column = input("with secound column? (y/n): ")
        if with_secound_column == 'y':
            secound_column_name = input("Enter secound column name : ")
            bot.main(secound_column_name)
        else:
            bot.main()
    else:
        print("The programmer Stoped the Proccess Please Contact to him for the new version")
        time.sleep(10)
        exit()






