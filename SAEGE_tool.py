import pandas as pd
import nltk
from nltk.tokenize import word_tokenize
from nltk.stem import WordNetLemmatizer

# Download NLTK resources
nltk.download('punkt_tab')
nltk.download('wordnet')

# Sample data structure for the Excel sheet (assuming it's already structured)
# The Excel sheet has columns: 'News Source', 'Fraud Type', 'Approach', 'Platform', 'Suspect Tactics', 'Impact'
excel_path = r"E:\SAEGE\Comprehensive_Cyber_Fraud_Cases.xlsx"  # Path to your Excel file

# Read the Excel sheet
fraud_data = pd.read_excel(excel_path)

# Sample text to analyze
input_text = "TD Fraud alert purchase $175.28 @ UKVI Credit card **13. Reply Y if this was you/ your add'l Cardholder, N if not Msg rates may apply"

# Preprocess the input text
def preprocess_text(text):
    lemmatizer = WordNetLemmatizer()
    tokens = word_tokenize(text)
    lemmatized_tokens = [lemmatizer.lemmatize(token.lower()) for token in tokens]
    return lemmatized_tokens

# Find Fraud Type
def find_fraud_type(input_text, fraud_data):
    # Preprocess the input text
    tokens = preprocess_text(input_text)

    # Create a list to store matching fraud types
    matched_fraud_types = []

    # Iterate through the fraud data
    for index, row in fraud_data.iterrows():
        # Tokenize the labels in the Excel sheet (fraud type, platform, tactics, etc.)
        fraud_type_tokens = preprocess_text(row['Fraud Type'])
        approach_tokens = preprocess_text(row['Approach'])
        platform_tokens = preprocess_text(row['Platform'])
        suspect_tactics_tokens = preprocess_text(row['Suspect Tactics'])

        # Check for matches
        common_tokens = set(tokens).intersection(fraud_type_tokens + approach_tokens + platform_tokens + suspect_tactics_tokens)
        if common_tokens:
            matched_fraud_types.append({
                'fraud_type': row['Fraud Type'],
                'common_tokens': common_tokens
            })

    return matched_fraud_types

# Analyze the input text and find possible fraud types
matched_fraud_types = find_fraud_type(input_text, fraud_data)

# Display the result
if matched_fraud_types:
    print("Matched Fraud Types:")
    for match in matched_fraud_types:
        print(f"Fraud Type: {match['fraud_type']} - Common Tokens: {', '.join(match['common_tokens'])}")
else:
    print("No fraud types matched.")
