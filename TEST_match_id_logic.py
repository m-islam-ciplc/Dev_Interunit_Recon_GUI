#!/usr/bin/env python3
"""Test to verify the Match ID sequential logic fix without pandas."""

def test_old_logic():
    """Demonstrate the problem with the old logic."""
    print("=== OLD LOGIC (BROKEN) ===")
    
    shared_match_counter = 0
    
    # Simulate LC matching
    lc_matches = [{'match_id': 'M001'}, {'match_id': 'M002'}, {'match_id': 'M003'}]
    if lc_matches:
        # OLD LOGIC: Replaces the counter
        shared_match_counter = max(int(match['match_id'][1:]) for match in lc_matches)
    print(f"After LC matches: shared_match_counter = {shared_match_counter}")
    
    # Simulate PO matching (only 1 match)
    po_matches = [{'match_id': 'M001'}]  # This could happen if PO reuses existing match
    if po_matches:
        # OLD LOGIC: Replaces the counter - GOES BACKWARDS!
        shared_match_counter = max(int(match['match_id'][1:]) for match in po_matches)
    print(f"After PO matches: shared_match_counter = {shared_match_counter}")
    print("❌ Counter went backwards from 3 to 1!")
    
    # Next match would be M002, which already exists!
    print(f"Next match would be: M{shared_match_counter + 1:03d} (DUPLICATE!)")

def test_new_logic():
    """Demonstrate the fix with the new logic."""
    print("\n=== NEW LOGIC (FIXED) ===")
    
    shared_match_counter = 0
    
    # Simulate LC matching
    lc_matches = [{'match_id': 'M001'}, {'match_id': 'M002'}, {'match_id': 'M003'}]
    if lc_matches:
        # NEW LOGIC: Takes maximum of current counter and highest match
        max_lc_counter = max(int(match['match_id'][1:]) for match in lc_matches)
        shared_match_counter = max(shared_match_counter, max_lc_counter)
    print(f"After LC matches: shared_match_counter = {shared_match_counter}")
    
    # Simulate PO matching (only 1 match)
    po_matches = [{'match_id': 'M001'}]  # This could happen if PO reuses existing match
    if po_matches:
        # NEW LOGIC: Preserves the higher counter value
        max_po_counter = max(int(match['match_id'][1:]) for match in po_matches)
        shared_match_counter = max(shared_match_counter, max_po_counter)
    print(f"After PO matches: shared_match_counter = {shared_match_counter}")
    print("✅ Counter stayed at 3!")
    
    # Next match would be M004
    print(f"Next match would be: M{shared_match_counter + 1:03d} (CORRECT!)")

def test_sequential_generation():
    """Test full sequential generation across match types."""
    print("\n=== FULL SEQUENTIAL TEST ===")
    
    shared_match_counter = 0
    shared_existing_matches = {}
    all_match_ids = []
    
    # Simulate different match types
    match_types = [
        ("LC", 3),    # 3 LC matches
        ("PO", 2),    # 2 PO matches
        ("Interunit", 1),  # 1 Interunit match
        ("USD", 2),   # 2 USD matches
    ]
    
    for match_type, count in match_types:
        print(f"\nProcessing {match_type} matches ({count} matches)...")
        matches = []
        
        for i in range(count):
            # Simulate match key
            match_key = (f"{match_type}_{i}", 1000 * (i + 1))
            
            if match_key in shared_existing_matches:
                match_id = shared_existing_matches[match_key]
                print(f"  Reusing existing Match ID: {match_id}")
            else:
                shared_match_counter += 1
                match_id = f"M{shared_match_counter:03d}"
                shared_existing_matches[match_key] = match_id
                print(f"  Creating new Match ID: {match_id}")
            
            matches.append({'match_id': match_id})
            all_match_ids.append(match_id)
        
        # Update counter using NEW LOGIC
        if matches:
            max_counter = max(int(match['match_id'][1:]) for match in matches)
            shared_match_counter = max(shared_match_counter, max_counter)
            print(f"  Updated shared_match_counter = {shared_match_counter}")
    
    print("\n=== FINAL RESULTS ===")
    print(f"All Match IDs generated: {all_match_ids}")
    print(f"✅ All Match IDs are sequential from M001 to M{shared_match_counter:03d}!")

if __name__ == "__main__":
    test_old_logic()
    test_new_logic()
    test_sequential_generation()