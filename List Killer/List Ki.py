import pandas as pd

def find_and_remove_duplicates(input_file, output_file, column_name='Name'):
    """
    Load an Excel file, identify duplicate values in a specified column,
    print duplicate counts, remove duplicates (keeping first occurrence),
    and save the deduplicated data to a new Excel file.

    Parameters:
    input_file (str): Path to the input Excel file.
    output_file (str): Path where the deduplicated Excel file will be saved.
    column_name (str): Name of the column to check for duplicates (default 'Name').
    """
    # Load the Excel file into a pandas DataFrame
    df = pd.read_excel(input_file)

    # Verify that the required column exists in the file
    if column_name not in df.columns:
        raise ValueError(f"The Excel file must contain a column named '{column_name}'")

    # Count occurrences of each unique value in the specified column
    # value_counts() returns a Series with counts sorted descending
    duplicate_counts = df[column_name].value_counts()
    # Keep only those with count > 1 (duplicates)
    duplicates = duplicate_counts[duplicate_counts > 1]

    # Report duplicates found (if any)
    if not duplicates.empty:
        print("Duplicate names found:\n")
        for name, count in duplicates.items():
            print(f"{name}: {count} times")
    else:
        print("No duplicate names found.")

    # Remove duplicate rows based on the specified column
    # keep='first' retains the first occurrence, drops later duplicates
    df_unique = df.drop_duplicates(subset=column_name, keep='first')

    # Save the cleaned DataFrame to a new Excel file
    df_unique.to_excel(output_file, index=False)
    print(f"\nDeduplicated names saved to '{output_file}'")


# Run the function with default values
if __name__ == "__main__":
    find_and_remove_duplicates(
        input_file='Book1.xlsx',
        output_file='names_deduplicated.xlsx',
        column_name='Name'
    )
