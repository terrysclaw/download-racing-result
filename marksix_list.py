import itertools
import csv


red_balls = [1, 2, 7, 8, 12, 13, 18, 19, 23, 24, 29, 30, 34, 35, 40, 45, 46]
blue_balls = [3, 4, 9, 10, 14, 15, 20, 25, 26, 31, 36, 37, 41, 42, 47, 48]
green_balls = [5, 6, 11, 16, 17, 21, 22, 27, 28, 32, 33, 38, 39, 43, 44, 49]

last_drawn_numbers = [6, 23, 28, 31, 33, 34, 11]  # 假設這是最近一次開獎的號碼

def get_marksix_combinations():
    """
    Generates all combinations of 6 numbers from 1 to 49.
    Returns an iterator.
    """
    numbers = list(range(1, 50))
    return itertools.combinations(numbers, 6)

def is_valid_combination(combo):
    """
    Returns True if:
    1. The sum of the combination is between 100 and 200 (inclusive).
    2. The combination is NOT all odd numbers.
    3. The combination is NOT all even numbers.
    4. The combination includes at least one number from each color (Red, Blue, Green).
    5. The combination includes at least 2 consecutive numbers.
    6. The combination includes at least 2 numbers with the same last digit.
    7. Interval check:
       - Numbers distributed across 3-4 intervals (1-10, 11-20, 21-30, 31-40, 41-49).
       - Max 3 numbers per interval.
    8. Include 1-2 numbers from last_drawn_numbers.
    """
    # Check sum
    total = sum(combo)
    if not (100 <= total <= 200):
        return False
        
    # Check odd/even
    # Count odd numbers
    odd_count = sum(1 for n in combo if n % 2 != 0)
    
    # If all odd (6) or all even (0), return False
    if odd_count == 0 or odd_count == 6:
        return False

    # Check colors
    # Convert lists to sets for faster lookup if this function is called many times, 
    # but for small lists 'in' operator is fast enough.
    # Using the global lists defined at the top.
    has_red = any(n in red_balls for n in combo)
    if not has_red: return False
    
    has_blue = any(n in blue_balls for n in combo)
    if not has_blue: return False
    
    has_green = any(n in green_balls for n in combo)
    if not has_green: return False

    # Check consecutive numbers
    # combo is sorted tuple
    has_consecutive = False
    for i in range(len(combo) - 1):
        if combo[i+1] == combo[i] + 1:
            has_consecutive = True
            break
    if not has_consecutive:
        return False

    # Check same last digit
    last_digits = [n % 10 for n in combo]
    if len(set(last_digits)) == len(combo): # If all unique, then no duplicates
        return False

    # Check intervals
    # Intervals: 1-10, 11-20, 21-30, 31-40, 41-49
    interval_counts = [0] * 5 
    for n in combo:
        if 1 <= n <= 10: interval_counts[0] += 1
        elif 11 <= n <= 20: interval_counts[1] += 1
        elif 21 <= n <= 30: interval_counts[2] += 1
        elif 31 <= n <= 40: interval_counts[3] += 1
        elif 41 <= n <= 49: interval_counts[4] += 1
    
    # Condition 1: Numbers must be distributed across 3-4 intervals
    active_intervals = sum(1 for count in interval_counts if count > 0)
    if not (3 <= active_intervals <= 4):
        return False
        
    # Condition 2: Max 3 numbers per interval
    if any(count > 3 for count in interval_counts):
        return False

    # Check last drawn numbers overlap
    # Must include 1-2 numbers from last_drawn_numbers
    overlap_count = sum(1 for n in combo if n in last_drawn_numbers)
    if not (1 <= overlap_count <= 2):
        return False
        
    return True

def save_combinations_to_csv(filename='marksix_combinations_filtered.csv'):
    print(f"Saving filtered combinations to {filename}...")
    combos = get_marksix_combinations()
    count = 0
    with open(filename, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['N1', 'N2', 'N3', 'N4', 'N5', 'N6', 'Sum'])
        for i, combo in enumerate(combos):
            if is_valid_combination(combo):
                writer.writerow(list(combo) + [sum(combo)])
                count += 1
            
            if (i + 1) % 1000000 == 0:
                print(f"Processed {i + 1} combinations... (Found {count} valid so far)")
    print(f"Done. Total valid combinations: {count}")

if __name__ == "__main__":
    print("Generating Mark Six combinations (1-49)...")
    print("Filtering:")
    print("1. Sum must be between 100 and 200 (inclusive).")
    print("2. Exclude all-odd and all-even combinations.")
    print("3. Must include at least one Red, one Blue, and one Green ball.")
    print("4. Must include at least 2 consecutive numbers.")
    print("5. Must include at least 2 numbers with the same last digit.")
    print("6. Interval check: 3-4 active intervals, max 3 per interval.")
    print("7. Must include 1-2 numbers from last_drawn_numbers.")
    
    # Total combinations: C(49, 6) = 13,983,816
    total_combinations = 13983816 
    print(f"Total possible combinations (before filter): {total_combinations:,}")
    
    print("First 10 valid combinations:")
    combos = get_marksix_combinations()
    shown = 0
    for combo in combos:
        if is_valid_combination(combo):
            print(f"{combo}, Sum: {sum(combo)}")
            shown += 1
            if shown >= 10:
                break
        
    print("...")

    print("Calculating total valid combinations (this may take a moment)...")
    valid_count = 0
    combos = get_marksix_combinations()
    for i, combo in enumerate(combos):
        if is_valid_combination(combo):
            valid_count += 1
        if (i + 1) % 5000000 == 0:
            print(f"Scanned {i + 1} combinations...")
            
    print(f"Total valid combinations (All filters applied): {valid_count:,}")
    print(f"Excluded combinations: {total_combinations - valid_count:,}")
    
    # Uncomment the following line to generate the full CSV file
    # save_combinations_to_csv()
