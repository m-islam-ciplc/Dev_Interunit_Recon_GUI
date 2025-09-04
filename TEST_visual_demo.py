#!/usr/bin/env python3
"""Visual demonstration of the Match ID sequencing fix."""

def print_box(title, content):
    """Print content in a nice box."""
    width = max(len(title), max(len(line) for line in content)) + 4
    print("┌" + "─" * width + "┐")
    print(f"│ {title.center(width-2)} │")
    print("├" + "─" * width + "┤")
    for line in content:
        print(f"│ {line.ljust(width-2)} │")
    print("└" + "─" * width + "┘")

print("\n🔍 MATCH ID SEQUENCING ISSUE - BEFORE AND AFTER FIX\n")

# Scenario 1: Before Fix
print("❌ BEFORE FIX - Non-Sequential Match IDs:")
print_box("Step 1: LC Matching", [
    "Found 3 LC matches",
    "Generated: M001, M002, M003",
    "shared_match_counter = 3  ← Set to max ID"
])
print("          ↓")
print_box("Step 2: PO Matching", [
    "Found 1 PO match (reused M001)",
    "Max ID in PO matches: M001",
    "shared_match_counter = 1  ← OVERWRITES to 1!"
])
print("          ↓")
print_box("Step 3: Interunit Matching", [
    "Starting from counter = 1",
    "New match generates: M002  ← DUPLICATE!",
    "ERROR: M002 already exists from LC!"
])

print("\n" + "="*60 + "\n")

# Scenario 2: After Fix
print("✅ AFTER FIX - Sequential Match IDs:")
print_box("Step 1: LC Matching", [
    "Found 3 LC matches",
    "Generated: M001, M002, M003",
    "shared_match_counter = max(0, 3) = 3"
])
print("          ↓")
print_box("Step 2: PO Matching", [
    "Found 1 PO match (reused M001)",
    "Max ID in PO matches: M001 (numeric: 1)",
    "shared_match_counter = max(3, 1) = 3  ← PRESERVES 3!"
])
print("          ↓")
print_box("Step 3: Interunit Matching", [
    "Starting from counter = 3",
    "New match generates: M004  ← SEQUENTIAL!",
    "shared_match_counter = max(3, 4) = 4"
])
print("          ↓")
print_box("Step 4: USD Matching", [
    "Starting from counter = 4",
    "New matches generate: M005, M006",
    "shared_match_counter = max(4, 6) = 6"
])

print("\n📊 SUMMARY:")
print_box("Result", [
    "✅ All Match IDs are sequential",
    "✅ No duplicates possible",
    "✅ Counter never goes backwards",
    "✅ M001, M002, M003, M004, M005, M006..."
])

print("\n🔧 THE FIX:")
print("OLD: shared_match_counter = max(int(match['match_id'][1:]) for match in matches)")
print("NEW: shared_match_counter = max(shared_match_counter, max_counter)")
print("\nThe key difference: NEW logic preserves the higher value!")
print("="*60)