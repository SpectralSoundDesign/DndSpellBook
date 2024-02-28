# Import the required libraries
import pandas as pd

def import_excel():
    # Set the path to the file
    file_path = '/home/willr1861/projects/spellbook/src/spells/DND Spells.xlsm'

    # Read the file
    data = pd.read_excel(file_path)

    # Strip whitespace from column names
    data.columns = data.columns.str.strip()

    # Clean up data by stripping whitespace from all string columns
    data = data.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # Convert the data to a dictionary
    data_dict = data.to_dict(orient='records')
    
    # Return the data dictionary
    return data_dict

def add_to_excel(spell):
    # Set the path to the file
    file_path = '/home/willr1861/projects/spellbook/src/spells/DND Spells.xlsm'

    # Read the file
    data = pd.read_excel(file_path)

    # Strip whitespace from column names
    data.columns = data.columns.str.strip()

    # Clean up data by stripping whitespace from all string columns
    data = data.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # Add the new spell to the dataframe
    spell_df = pd.DataFrame([spell])
    
    data = data._append(spell_df, ignore_index=True)

    # Write the data back to the file
    data.to_excel(file_path, index=False)
