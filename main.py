import pandas as pd
import time
from googletrans import Translator

# Initialize google translate
translator = Translator()

# Change source and destination language if needed: 
def translate_cell(text, src_lang='nl', dest_lang='en'):
    try:
        # Translate the text
        translated_text = translator.translate(text, src=src_lang, dest=dest_lang).text
    except Exception as e:
        print(f"Error during translation: {e}")
        translated_text = text  # Return the original text if translation fails
    return translated_text


def select_columns_for_translation(df):
    # Display all column names
    print("The following columns are available for translation:")
    for idx, column in enumerate(df.columns):
        print(f"{idx + 1}. {column}")

    # User selects which columns to translate
    selected_indices = input(
        "Enter the numbers of the columns you wish to translate, separated by commas (e.g., 1,3,4): ")

    # Convert input into list of column names
    try:
        selected_indices = [int(i.strip()) - 1 for i in selected_indices.split(',')]
        selected_columns = [df.columns[i] for i in selected_indices]
    except (IndexError, ValueError):
        print("Invalid selection. Please enter the correct column numbers separated by commas.")
        return None

    return selected_columns


def translate_columns(file_path, src_lang='nl', dest_lang='en'):
    # Read the Excel file
    df = pd.read_excel(file_path, header=[0,1])

    # Get user-selected columns
    columns_to_translate = select_columns_for_translation(df)
    if not columns_to_translate:
        return

    # Translate selected columns
    for column in columns_to_translate:
        print(f'Translating column: {column}')
        if df[column].dtype == object:  # Check if the column contains text
            # Translate each cell in the column
            for idx, cell in df[column].items():
                if pd.isnull(cell):
                    print(f"Empty row: {idx+1} Skipping")
                    continue  # Skip null or NaN values
                df.at[idx, column] = translate_cell(str(cell), src_lang, dest_lang)
                time.sleep(0.3)
                print(f"Translating row: {idx+1}")
        else:
            print(f"Skipping column {column} - not a text column.")

    # Save the translated DataFrame to a new Excel file
    translated_file_path = file_path.replace('.xlsx', '_translated.xlsx')
    df.to_excel(translated_file_path, index=True)
    print(f'Translation completed. Translated file saved as {translated_file_path}')


# Provide path to the input file:
translate_columns('Input_file.xlsx')
