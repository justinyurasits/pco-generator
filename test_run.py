#!/usr/bin/env python3
"""
Test the change order generator with realistic messy field notes.
This is what a super actually sends — run this to validate output quality.
"""

from generator import generate

# ---------------------------------------------------------------------------
# Test Case 1: Unforeseen condition — the bread and butter scenario
# Messy field notes, voice-memo style, missing nothing critical
# ---------------------------------------------------------------------------

test_1 = {
    "company_name": "Beacon Hill Builders",
    "project_name": "Johnson Residence Addition",
    "project_address": "47 Maple Street, Lexington, MA 01773",
    "change_order_number": "003",
    "date": "June 19, 2026",
    "original_contract_date": "March 1, 2026",
    "scope_description": (
        "found rot in the subfloor under the master bath when we pulled up the tile. "
        "about 60 sq ft of 3/4 plywood needs to come out, plus we gotta sister 4 of "
        "the floor joists where they're soft. owner was on site and said go ahead. "
        "gonna need to rent a floor scraper too and pick up some joist hangers"
    ),
    "labor_cost": "2,400",
    "material_cost": "680",
    "schedule_days": "2",
    "reason_for_change": "unforeseen condition"
}

# ---------------------------------------------------------------------------
# Test Case 2: Owner-requested change — missing cost data
# Tests that [TBD] appears correctly and no numbers are fabricated
# ---------------------------------------------------------------------------

test_2 = {
    "company_name": "Beacon Hill Builders",
    "project_name": "Sullivan Kitchen Renovation",
    "project_address": "112 Commonwealth Ave, Boston, MA 02116",
    "change_order_number": "001",
    "date": "June 19, 2026",
    "original_contract_date": "April 15, 2026",
    "scope_description": (
        "owner wants to extend the kitchen island by 18 inches to add seating for 2 more. "
        "means we need more countertop material, extend the base cabinets, and move the "
        "electrical outlet on the island end. owner said they'll pick the countertop material "
        "this weekend so we don't have pricing yet"
    ),
    "labor_cost": "",          # missing — should produce [TBD]
    "material_cost": "",       # missing — should produce [TBD]
    "schedule_days": "3",
    "reason_for_change": "owner-requested"
}

# ---------------------------------------------------------------------------
# Test Case 3: Design revision — cleaner input, all fields present
# Tests professional language from already-formal field notes
# ---------------------------------------------------------------------------

test_3 = {
    "company_name": "Beacon Hill Builders",
    "project_name": "Westwood Custom Home",
    "project_address": "8 Pine Ridge Road, Westwood, MA 02090",
    "change_order_number": "007",
    "date": "June 19, 2026",
    "original_contract_date": "January 10, 2026",
    "scope_description": (
        "architect issued revised window schedule on 6/17. window W-4 on the east elevation "
        "changed from a 3068 double hung to a 4068 casement. requires framing modification "
        "to rough opening — widen by 12 inches and add a new king stud and jack stud on "
        "the south side. existing lintel is adequate."
    ),
    "labor_cost": "1,100",
    "material_cost": "340",
    "schedule_days": "1",
    "reason_for_change": "design revision"
}


if __name__ == "__main__":
    import sys

    # Run one test case or all three
    test_case = int(sys.argv[1]) if len(sys.argv) > 1 else 1

    tests = {1: test_1, 2: test_2, 3: test_3}

    if test_case not in tests:
        print(f"Usage: python test_run.py [1|2|3]")
        sys.exit(1)

    data = tests[test_case]
    print(f"\nRunning Test Case {test_case}: {data['project_name']}")
    print(f"Scenario: {data['reason_for_change']}")
    print("-" * 60)

    result = generate(data, output_dir="./output")

    print("\n--- GENERATED TEXT ---")
    print(result["generated_text"])
    print("\n--- PARSED SECTIONS ---")
    for k, v in result["sections"].items():
        print(f"\n[{k}]")
        print(v)
