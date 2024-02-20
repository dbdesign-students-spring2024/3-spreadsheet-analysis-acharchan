# place your code to clean up the data file below.
import os

import matplotlib.pyplot as plt
import pandas as pd


def scrub_data(input_file_path, output_file_path):
    # Read the Excel file
    df = pd.read_excel(input_file_path, skiprows=7)

    # Drop the unnamed column
    df = df.drop(columns=df.columns[0])

    # Removing unnecessary columns
    columns_to_remove = ['Base Year']
    df = df.drop(columns=columns_to_remove)

    # Replace all '-' and '...' with None in the entire DataFrame
    df = df.replace(['-', '...'], None)

    # Rename columns for better clarity
    column_mapping = {'Unnamed: 2': 'Abbreviation'}

    df = df.rename(columns=column_mapping)

    # Save the modified data to a new CSV file
    df.to_csv(output_file_path, index=False)

    print(f"\nScrubbed data saved to {output_file_path}")

    # Perform data analysis and save to a new spreadsheet
    analyze_and_save_to_spreadsheet(output_file_path)


def analyze_and_save_to_spreadsheet(file_path):
    df = pd.read_csv(file_path)

    # Skip specified columns when creating the summary table
    skip_columns = ['Indicator', 'Abbreviation', 'Scale']

    # Create a new Excel file
    excel_file_path = 'data/analysis_results_function.xlsx'
    images_folder = 'images'

    # Create the images folder if it doesn't exist
    os.makedirs(images_folder, exist_ok=True)

    with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
        # Create a summary table for all years
        summary_table = pd.DataFrame({
            'Min': df.drop(columns=skip_columns).min(skipna=True),
            'Max': df.drop(columns=skip_columns).max(skipna=True),
            'Sum': df.drop(columns=skip_columns).sum(skipna=True),
            'Mean': df.drop(columns=skip_columns).mean(skipna=True)
        })

        # Write the summary table to the first sheet
        summary_table.to_excel(writer, sheet_name='Summary_Table', index=True)

        # Write summary with condition
        row_sums = df.iloc[:, 3:].sum(axis=1)

        # Filter indicators based on sum greater than 1,000,000
        selected_indicators = df.loc[row_sums > 1000000]

        # Calculate statistics for selected indicators
        summary_table_with_condition = pd.DataFrame({
            'Indicator': selected_indicators['Indicator'],
            'Min': selected_indicators.iloc[:, 3:].min(axis=1, skipna=True),
            'Max': selected_indicators.iloc[:, 3:].max(axis=1, skipna=True),
            'Mean': selected_indicators.iloc[:, 3:].mean(axis=1, skipna=True),
            'Sum': selected_indicators.iloc[:, 3:].sum(axis=1, skipna=True)
        })
        summary_table_with_condition.to_excel(writer, sheet_name='Summary_Table_With_Condition', index=True)

        # Create the images folder if it doesn't exist
        images_folder = 'images'
        os.makedirs(images_folder, exist_ok=True)

        for index, row in df.iterrows():
            indicator = row['Indicator']

            # Convert row data to NumPy array
            row_data = pd.to_numeric(row[3:], errors='coerce').values

            plt.figure(figsize=(15, 6))
            plt.plot(df.columns[3:].to_numpy(), row_data, marker='o')
            plt.title(f'{indicator} Over Years')
            plt.xlabel('Year')
            plt.ylabel('Value')
            plt.xticks(df.columns[3:].to_numpy())
            plt.grid(True)

            # Replace spaces with underscores and remove commas in the image file name
            image_name = f'{indicator}_plot.png'.replace(' ', '_').replace(',', '')

            image_path = os.path.join(images_folder, image_name)
            plt.savefig(image_path)
            plt.close()

        selected_columns = [col for col in df.columns if pd.to_numeric(df[col], errors='coerce').sum() > 10000000]

        # Plotting the bar chart for the sum of selected indicators
        total_sum_series = df.drop(columns=skip_columns)[selected_columns].sum().squeeze()
        plt.figure(figsize=(15, 6))
        plt.bar(selected_columns, total_sum_series, label='Total Sum', color='blue')
        plt.title('Total Sum of Selected Indicators Over Years')
        plt.xlabel('Indicator')
        plt.ylabel('Sum')
        plt.grid(True)
        plt.legend()
        sum_chart_path = os.path.join(images_folder, 'total_sum_chart.png')
        plt.savefig(sum_chart_path)
        plt.close()

        print(f'Plots saved to {images_folder}')

    print(f'Analysis saved to {excel_file_path}')


# Usage
input_file_path = 'data/international_liquidity.xlsx'
output_file_path = 'data/cleaned_data.csv'
scrub_data(input_file_path, output_file_path)
