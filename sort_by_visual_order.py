import pandas as pd
import locale
import unicodedata
import re

# Set Hebrew locale
try:
    locale.setlocale(locale.LC_ALL, 'he_IL.UTF-8')
except locale.Error:
    # Fallback to system default if Hebrew locale not available
    print("Warning: Hebrew locale not found, using system default")
    locale.setlocale(locale.LC_ALL, '')

def normalize_hebrew(text):
    """Normalize Hebrew text by removing control characters and normalizing unicode"""
    # Convert to string if not already
    text = str(text)
    # Remove control characters
    text = ''.join(char for char in text if not unicodedata.category(char).startswith('C'))
    # Normalize unicode
    return unicodedata.normalize('NFKC', text)

def strip_bidi_controls(text):
    # Remove all Unicode bidirectional control characters
    return re.sub(r'[\u200e\u200f\u202a-\u202e\u2066-\u2069]', '', str(text))

def get_last_word(text):
    # Split by whitespace and return the last non-empty word, stripped of trailing apostrophes/geresh/gershayim
    words = [w for w in text.strip().split() if w]
    if not words:
        return ''
    last = words[-1]
    # Remove trailing apostrophe, geresh, gershayim, or similar
    return re.sub(r"[׳״'\"`]+$", '', last)

# Load original Excel
df = pd.read_excel('prices-original.xlsx')

# Get the descriptions column (Column B)
descriptions = df.iloc[:, 1].astype(str).tolist()

# Clean and extract last word
sort_keys = [get_last_word(strip_bidi_controls(desc)) for desc in descriptions]

# Sort by the last word
sorted_indices = sorted(range(len(sort_keys)), key=lambda i: sort_keys[i])

# Reorder the dataframe based on the sorted indices
df_sorted = df.iloc[sorted_indices].reset_index(drop=True)

# Save output
df_sorted.to_excel('sorted_output.xlsx', index=False)
print("✅ Sorted and saved to sorted_output.xlsx")
