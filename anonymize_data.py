import pandas as pd
import random

# Load the Excel file
file_path = 'input_list_sheets.xlsx'
data = pd.ExcelFile(file_path)

# Load all sheets into a dictionary
sheets_data = {sheet_name: data.parse(sheet_name) for sheet_name in data.sheet_names}

# Function to generate a random English name
def generate_random_name():
    # Random English names and surnames 
    first_names = [
        "James", "John", "Robert", "Michael", "William", "David", "Richard", "Joseph", "Thomas", "Charles",
        "Christopher", "Daniel", "Matthew", "Anthony", "Mark", "Donald", "Paul", "Steven", "Andrew", "Kenneth",
        "George", "Joshua", "Kevin", "Brian", "Edward", "Ronald", "Timothy", "Jason", "Jeffrey", "Ryan",
        "Jacob", "Gary", "Nicholas", "Eric", "Jonathan", "Stephen", "Larry", "Justin", "Scott", "Brandon",
        "Benjamin", "Samuel", "Frank", "Gregory", "Raymond", "Alexander", "Patrick", "Jack", "Dennis", "Jerry",
        "Tyler", "Aaron", "Jose", "Adam", "Nathan", "Henry", "Douglas", "Zachary", "Peter", "Kyle", "Walter",
        "Ethan", "Jeremy", "Harold", "Keith", "Christian", "Roger", "Noah", "Gerald", "Carl", "Terry", "Sean",
        "Austin", "Arthur", "Lawrence", "Jesse", "Dylan", "Bryan", "Joe", "Jordan", "Billy", "Bruce", "Albert",
        "Willie", "Gabriel", "Logan", "Alan", "Juan", "Wayne", "Roy", "Ralph", "Randy", "Eugene", "Vincent"
    ] * 5  # Multiply to increase size

    last_names = [
        "Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis", "Rodriguez", "Martinez",
        "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "Martin",
        "Lee", "Perez", "Thompson", "White", "Harris", "Sanchez", "Clark", "Ramirez", "Lewis", "Robinson",
        "Walker", "Young", "Allen", "King", "Wright", "Scott", "Torres", "Nguyen", "Hill", "Flores", "Green",
        "Adams", "Nelson", "Baker", "Hall", "Rivera", "Campbell", "Mitchell", "Carter", "Roberts", "Gomez",
        "Phillips", "Evans", "Turner", "Diaz", "Parker", "Cruz", "Edwards", "Collins", "Reyes", "Stewart",
        "Morris", "Morales", "Murphy", "Cook", "Rogers", "Gutierrez", "Ortiz", "Morgan", "Cooper", "Peterson",
        "Bailey", "Reed", "Kelly", "Howard", "Ramos", "Kim", "Cox", "Ward", "Richardson", "Watson", "Brooks",
        "Chavez", "Wood", "James", "Bennett", "Gray", "Mendoza", "Alvarez", "Ruiz", "Hughes", "Price", "Myers"
    ] * 5  # Multiply to increase size
    return f"{random.choice(first_names)} {random.choice(last_names)}"

# Function to generate a random Turkish phone number
def generate_random_phone():
    return f"(5{random.randint(10, 59)}) {random.randint(100, 999)}-{random.randint(1000, 9999)}"

# Process each sheet
processed_sheets = {}
for sheet_name, sheet_data in sheets_data.items():
    sheet_data = sheet_data.copy()
    
    # Replace names and phone numbers
    if 'Unnamed: 1' in sheet_data.columns:
        sheet_data['Unnamed: 1'] = sheet_data['Unnamed: 1'].apply(
            lambda x: generate_random_name() if isinstance(x, str) and x.strip() else x
        )
    if 'Unnamed: 2' in sheet_data.columns:
        sheet_data['Unnamed: 2'] = sheet_data['Unnamed: 2'].apply(
            lambda x: generate_random_phone() if isinstance(x, str) and x.strip() else x
        )
    
    processed_sheets[sheet_name] = sheet_data

# Save to a new Excel file
output_path = 'output_list_sheets.xlsx'
with pd.ExcelWriter(output_path) as writer:
    for sheet_name, sheet_data in processed_sheets.items():
        sheet_data.to_excel(writer, index=False, sheet_name=sheet_name)
