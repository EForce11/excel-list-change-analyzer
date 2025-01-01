import pandas as pd

# Function to reformat phone numbers
def reformat_phone_number(phone_number):
    if pd.isna(phone_number) or not phone_number:
        return None  # Return None for missing or empty phone numbers
    phone_number = ''.join(filter(str.isdigit, str(phone_number)))
    if len(phone_number) < 10:  # Invalid length check
        return None
    return f"+90 {phone_number[:3]} {phone_number[3:6]} {phone_number[6:8]} {phone_number[8:]}"

# Function to get members and their phone numbers from a dataframe with formatted phone numbers
def get_members_with_formatted_phones(data, members):
    members_with_phones = []
    for _, row in data.iterrows():
        if pd.isna(row['Unnamed: 1']) or pd.isna(row['Unnamed: 2']):
            continue  # Skip empty values
        if row['Unnamed: 1'] in members:
            formatted_phone = reformat_phone_number(row['Unnamed: 2'])
            if formatted_phone is not None:  # Add valid phone numbers
                members_with_phones.append((row['Unnamed: 1'], formatted_phone))
    return members_with_phones

# Function to find members in multiple supervisor sheets
def find_members_in_multiple_supervisors(data, supervisors):
    multiple_supervisors = {}

    # Iterate over each supervisor and find their members
    for supervisor in supervisors:
        # Get the data for the current supervisor sheet
        supervisor_data = data[supervisor]

        # Extract members of the current supervisor
        supervisor_members = set(supervisor_data['Unnamed: 1'][4:])

        # Compare with other supervisor sheets
        for other_supervisor in supervisors:
            if other_supervisor != supervisor:
                other_data = data[other_supervisor]
                other_members = set(other_data['Unnamed: 1'][4:])

                # Find members who appear in both supervisor sheets
                common_members = supervisor_members.intersection(other_members)

                # Store in multiple_supervisors dictionary
                for member in common_members:
                    if member not in multiple_supervisors:
                        multiple_supervisors[member] = set()
                    multiple_supervisors[member].add(other_supervisor)

    # Convert set to sorted list for consistent output
    for member in multiple_supervisors:
        multiple_supervisors[member] = sorted(list(multiple_supervisors[member]))

    return multiple_supervisors

# Function to print multiple supervisors warning
def print_multiple_supervisors_warning(multiple_supervisors):
    if multiple_supervisors:
        print("### WARNING: Members Belonging to Multiple Supervisors")
        for member, supervisors in multiple_supervisors.items():
            print(f"- {member}: {', '.join(supervisors)}")

# Function to print supervisor changes
def print_supervisor_changes(added_members_with_phones, removed_members_with_phones, multiple_supervisors):
    # Supervisors list
    supervisors = added_members_with_phones.keys()

    # Print added and removed members for each supervisor
    for supervisor in supervisors:
        print(f"### {supervisor}")

        print("#### ADDED MEMBERS")
        if added_members_with_phones[supervisor]:
            for member, phone in added_members_with_phones[supervisor]:
                print(f"- {member}: {phone}")
        else:
            print("- No members added.")

        print("\n#### REMOVED MEMBERS")
        if removed_members_with_phones[supervisor]:
            for member, phone in removed_members_with_phones[supervisor]:
                print(f"- {member}: {phone}")
        else:
            print("- No members removed.")
        print("\n")

    # Print multiple supervisors warning
    print_multiple_supervisors_warning(multiple_supervisors)

# Main function to process Excel files and print supervisor changes
def main(file_before, file_after):
    # Read the sheet names to iterate over them
    sheets_before = pd.ExcelFile(file_before).sheet_names
    sheets_after = pd.ExcelFile(file_after).sheet_names

    # Initialize dictionaries to store the data from each sheet
    data_before = {}
    data_after = {}

    # Read the data from each sheet into the dictionaries
    for sheet in sheets_before:
        data_before[sheet] = pd.read_excel(file_before, sheet_name=sheet).dropna(subset=['Unnamed: 1', 'Unnamed: 2'])

    for sheet in sheets_after:
        data_after[sheet] = pd.read_excel(file_after, sheet_name=sheet).dropna(subset=['Unnamed: 1', 'Unnamed: 2'])

    # Initialize dictionaries to store differences for each supervisor
    added_members_with_phones = {}
    removed_members_with_phones = {}

    # Iterate over each sheet and get the members with their formatted phone numbers
    for sheet in sheets_before:
        # Get the data for the current sheet
        before_data = data_before[sheet]
        after_data = data_after[sheet]

        # Extract the names of members before and after the changes
        before_members = set(before_data['Unnamed: 1'][4:])
        after_members = set(after_data['Unnamed: 1'][4:])

        # Determine which members were added and which were removed
        added = after_members - before_members
        removed = before_members - after_members

        # Get the added and removed members with their formatted phone numbers
        added = get_members_with_formatted_phones(after_data, added)
        removed = get_members_with_formatted_phones(before_data, removed)

        # Store the results in the dictionaries
        added_members_with_phones[sheet] = added
        removed_members_with_phones[sheet] = removed

    # Find members in multiple supervisor sheets
    multiple_supervisors = find_members_in_multiple_supervisors(data_after, sheets_after)

    # Print supervisor changes
    print_supervisor_changes(added_members_with_phones, removed_members_with_phones, multiple_supervisors)

# Paths to Excel files
file_before = 'file_before.xlsx'
file_after = 'file_after.xlsx'

# Call the main function to process files and print supervisor changes
main(file_before, file_after)
