#!/usr/bin/env python3
"""Diagnose script to verify the match counter updates in excel_transaction_matcher.py"""

import re

def check_counter_updates():
    """Check that all counter updates use the max() logic correctly."""
    
    with open('excel_transaction_matcher.py', 'r') as f:
        content = f.read()
    
    # Find all counter update patterns
    old_pattern = r'shared_match_counter = max\(int\(match\[\'match_id\'\]\[1:\]\) for match in \w+_matches\)'
    new_pattern = r'shared_match_counter = max\(shared_match_counter, max_\w+_counter\)'
    
    old_matches = list(re.finditer(old_pattern, content))
    new_matches = list(re.finditer(new_pattern, content))
    
    print("=== CHECKING MATCH COUNTER UPDATES ===\n")
    
    if old_matches:
        print(f"❌ Found {len(old_matches)} instances of OLD logic (direct assignment):")
        for match in old_matches:
            line_num = content[:match.start()].count('\n') + 1
            print(f"  Line {line_num}: {match.group()}")
    else:
        print("✅ No instances of OLD logic found (good!)")
    
    print(f"\n✅ Found {len(new_matches)} instances of NEW logic (max preservation):")
    for match in new_matches:
        line_num = content[:match.start()].count('\n') + 1
        print(f"  Line {line_num}: {match.group()}")
    
    # Check for each match type
    match_types = ['LC', 'PO', 'Interunit', 'USD']
    print("\n=== CHECKING EACH MATCH TYPE ===")
    
    for match_type in match_types:
        # Find the section for this match type
        if match_type == 'Interunit':
            section_pattern = rf'# Find interunit loan matches.*?print\(f"\\nInterunit Loan Matching Results:'
        else:
            section_pattern = rf'# Find {match_type} matches.*?print\(f"\\n{match_type} Matching Results:'
        
        section_match = re.search(section_pattern, content, re.DOTALL)
        if section_match:
            section = section_match.group()
            
            # Check if it has the counter update
            if f'max_{match_type.lower()}_counter' in section or f'max_interunit_counter' in section:
                print(f"✅ {match_type}: Counter update found with max() logic")
            else:
                print(f"❌ {match_type}: Missing counter update!")
        else:
            print(f"⚠️  {match_type}: Could not find section")

if __name__ == "__main__":
    check_counter_updates()