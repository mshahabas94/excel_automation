import pandas as pd
import re

def fix_product_ids(df):
    """
    Fix product IDs that weren't properly extracted from descriptions
    """
    fixed_count = 0
    
    for idx, row in df.iterrows():
        product_id = row['Product ID']
        description = row['Short Description']
        
        # Skip if description is empty or already has a good product ID
        if pd.isna(description) or description == '':
            continue
            
        desc_str = str(description).strip()
        
        # Pattern 1: Parentheses pattern - (82X700BTAX), (83DV00UPPS), etc.
        match = re.search(r'\(([A-Z0-9#\-]+)\)', desc_str)
        if match:
            extracted_id = match.group(1)
            # Only update if the current product ID doesn't match the extracted one
            if extracted_id != product_id and len(extracted_id) >= 6:
                df.at[idx, 'Product ID'] = extracted_id
                fixed_count += 1
                continue
        
        # Pattern 2: Dell monitor patterns - P3424WEB, U2724D, 210-BKVB
        if 'dell' in desc_str.lower() and 'monitor' in desc_str.lower():
            patterns = [
                r'([A-Z][0-9A-Z]{5,})',  # Like P3424WEB, U2724D
                r'(\d{3}-[A-Z]+)',       # Like 210-BKVB
            ]
            for pattern in patterns:
                match = re.search(pattern, desc_str)
                if match:
                    extracted_id = match.group(1)
                    if extracted_id != product_id:
                        df.at[idx, 'Product ID'] = extracted_id
                        fixed_count += 1
                        break
        
        # Pattern 3: HP product patterns - 65P58AS#ABV, 9M9D7AT
        if 'hp' in desc_str.lower():
            # HP Monitor pattern: 65P58AS#ABV – HP V24i G5
            match = re.search(r'^([A-Z0-9#]+)[\s–\-]', desc_str)
            if match:
                extracted_id = match.group(1)
                if extracted_id != product_id:
                    df.at[idx, 'Product ID'] = extracted_id
                    fixed_count += 1
                    continue
            
            # HP PC pattern: (9M9D7AT)
            match = re.search(r'\(([A-Z0-9]+)\)', desc_str)
            if match:
                extracted_id = match.group(1)
                if extracted_id != product_id:
                    df.at[idx, 'Product ID'] = extracted_id
                    fixed_count += 1
                    continue
        
        # Pattern 4: Microsoft/ASUS parentheses pattern
        if 'microsoft' in desc_str.lower() or 'asus' in desc_str.lower():
            match = re.search(r'\(([A-Z0-9\-]+)\)', desc_str)
            if match:
                extracted_id = match.group(1)
                if extracted_id != product_id:
                    df.at[idx, 'Product ID'] = extracted_id
                    fixed_count += 1
                    continue
        
        # Pattern 5: Generic parentheses pattern for any product
        match = re.search(r'\(([A-Z0-9#\-]{6,})\)', desc_str)
        if match:
            extracted_id = match.group(1)
            # Only update if it looks like a product ID (not too long, not just numbers)
            if (extracted_id != product_id and 
                len(extracted_id) <= 20 and 
                not extracted_id.isdigit()):
                df.at[idx, 'Product ID'] = extracted_id
                fixed_count += 1
    
    return df, fixed_count

def clean_cost_column(df):
    """
    Clean the Cost column - remove dollar signs and convert to numeric
    """
    # Remove dollar signs and any extra spaces
    df['Cost'] = df['Cost'].astype(str).str.replace('$', '', regex=False).str.strip()
    
    # Convert to numeric, setting errors to NaN
    df['Cost'] = pd.to_numeric(df['Cost'], errors='coerce')
    
    return df

# Read the current CSV
df = pd.read_csv("final_merged.csv")

print(f"Original data: {len(df)} rows")

# Fix product IDs
df, fixed_count = fix_product_ids(df)
print(f"Fixed {fixed_count} product IDs")

# Clean cost column
df = clean_cost_column(df)

# Remove rows with empty Product ID or Cost
df = df.dropna(subset=['Product ID', 'Cost'])
df = df[df['Product ID'] != '']
df = df.reset_index(drop=True)

print(f"Cleaned data: {len(df)} rows")

# Save the fixed CSV
df.to_csv("final_merged_fixed.csv", index=False)

print("✅ Fixed data saved to final_merged_fixed.csv")

# Show some examples of fixed product IDs
print("\n📋 Sample of fixed product IDs:")
sample_df = df.head(20)[['Product ID', 'Short Description']]
for idx, row in sample_df.iterrows():
    desc = str(row['Short Description'])[:80] + "..." if len(str(row['Short Description'])) > 80 else str(row['Short Description'])
    print(f"  {row['Product ID']} - {desc}")