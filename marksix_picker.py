import random

## 從 1-49 隨機選取 7 個不重複的號碼

red_balls = [1, 2, 7, 8, 12, 13, 18, 19, 23, 24, 29, 30, 34, 35, 40, 45, 46]
blue_balls = [3, 4, 9, 10, 14, 15, 20, 25, 26, 31, 36, 37, 41, 42, 47, 48]
green_balls = [5, 6, 11, 16, 17, 21, 22, 27, 28, 32, 33, 38, 39, 43, 44, 49]

before_drawn_numbers = [4, 6, 26, 28, 34, 40, 25]  # 假設這是前一次開獎的號碼
last_drawn_numbers = [1, 5, 6, 25, 30, 42, 43]  # 假設這是最近一次開獎的號碼


def has_consecutive_numbers(numbers, min_consecutive=2):
    """檢查號碼中是否包含至少指定數量的連續號碼"""
    sorted_nums = sorted(numbers)
    consecutive_count = 1
    
    for i in range(1, len(sorted_nums)):
        if sorted_nums[i] == sorted_nums[i-1] + 1:
            consecutive_count += 1
            if consecutive_count >= min_consecutive:
                return True
        else:
            consecutive_count = 1
    
    return False


def has_same_last_digit(numbers, min_count=2):
    """檢查號碼中是否包含至少指定數量的同尾號碼"""
    last_digits = [n % 10 for n in numbers]
    for digit in set(last_digits):
        if last_digits.count(digit) >= min_count:
            return True
    return False


def has_overlap_with_last_draw(numbers, last_draw, min_overlap=1, max_overlap=3):
    """檢查號碼中是否包含指定數量範圍的上期號碼"""
    overlap_count = len(set(numbers) & set(last_draw))
    return min_overlap <= overlap_count <= max_overlap


def has_all_colors(numbers):
    """檢查號碼中是否包含至少一個紅色、一個藍色和一個綠色號碼"""
    has_red = any(n in red_balls for n in numbers)
    has_blue = any(n in blue_balls for n in numbers)
    has_green = any(n in green_balls for n in numbers)
    return has_red and has_blue and has_green


def has_valid_sum(numbers, min_sum=100, max_sum=200):
    """檢查最小6個號碼總和及最大6個號碼總和是否在指定範圍內"""
    sorted_nums = sorted(numbers)
    sum_min_6 = sum(sorted_nums[:6])
    sum_max_6 = sum(sorted_nums[1:])
    return (min_sum <= sum_min_6 <= max_sum) and (min_sum <= sum_max_6 <= max_sum)


def check_intervals(numbers):
    """
    檢查區間條件:
    1. 7個號碼當中, 只能出現在其中3-4組
    2. 每個區間最多只能包括3個號碼
    區間: 1-10, 11-20, 21-30, 31-40, 41-49
    """
    groups = {
        '1-10': 0,
        '11-20': 0,
        '21-30': 0,
        '31-40': 0,
        '41-49': 0
    }
    
    for n in numbers:
        if 1 <= n <= 10:
            groups['1-10'] += 1
        elif 11 <= n <= 20:
            groups['11-20'] += 1
        elif 21 <= n <= 30:
            groups['21-30'] += 1
        elif 31 <= n <= 40:
            groups['31-40'] += 1
        elif 41 <= n <= 49:
            groups['41-49'] += 1
            
    # 條件1: 只能出現在其中3-4組 (即有號碼的組數為 3 或 4)
    active_groups = sum(1 for count in groups.values() if count > 0)
    if not (3 <= active_groups <= 4):
        return False
        
    # 條件2: 每個區間最多只能包括3個號碼
    if any(count > 3 for count in groups.values()):
        return False
        
    return True


def has_mixed_parity(numbers):
    """檢查號碼是否包含單雙數混合 (避免全單或全雙)"""
    odds = sum(1 for n in numbers if n % 2 != 0)
    evens = len(numbers) - odds
    return odds > 0 and evens > 0


def has_valid_range(numbers, min_diff=25):
    """檢查最大號碼和最小號碼的差距是否大於指定值"""
    return (max(numbers) - min(numbers)) > min_diff


def pick_marksix_numbers():
    """選取7個號碼，確保符合所有條件：
    1. 至少包含兩個連續號碼
    2. 至少包含兩個同尾號碼
    3. 包含1-3個上期號碼
    4. 至少包含紅、藍、綠波各一
    5. 最小6個及最大6個號碼總和介乎 100 - 200
    6. 區間分佈: 3-4組有號碼, 每組最多3個
    7. 包含1-3個前一次開獎號碼
    8. 包含單雙數混合
    9. 最大號碼和最小號碼差距大於25
    """
    count = 0
    while True:
        count += 1
        numbers = random.sample(range(1, 50), 7)
        if (has_consecutive_numbers(numbers, min_consecutive=2) and 
            has_same_last_digit(numbers, min_count=2) and
            has_overlap_with_last_draw(numbers, last_drawn_numbers, min_overlap=1, max_overlap=3) and
            has_overlap_with_last_draw(numbers, before_drawn_numbers, min_overlap=1, max_overlap=3) and
            has_all_colors(numbers) and
            has_valid_sum(numbers, min_sum=100, max_sum=200) and
            check_intervals(numbers) and
            has_mixed_parity(numbers) and
            has_valid_range(numbers, min_diff=25)):
            return numbers, count
        # 不符合條件，重新抽取


def print_analysis(numbers):
    sorted_nums = sorted(numbers)
    print("\n--- 符合條件詳情 ---")
    
    # 1. 連續號碼
    consecutive_groups = []
    current_group = [sorted_nums[0]]
    for i in range(1, len(sorted_nums)):
        if sorted_nums[i] == sorted_nums[i-1] + 1:
            current_group.append(sorted_nums[i])
        else:
            if len(current_group) >= 2:
                consecutive_groups.append(current_group)
            current_group = [sorted_nums[i]]
    if len(current_group) >= 2:
        consecutive_groups.append(current_group)
    print(f"1. 連續號碼: {consecutive_groups}")

    # 2. 同尾號碼
    same_digits = {}
    for n in sorted_nums:
        d = n % 10
        same_digits.setdefault(d, []).append(n)
    same_digits_list = [v for k, v in same_digits.items() if len(v) >= 2]
    print(f"2. 同尾號碼: {same_digits_list}")

    # 3. 上期重號
    overlap = sorted(list(set(numbers) & set(last_drawn_numbers)))
    print(f"3. 上期重號: {overlap}")

    # 3.1 前一次重號
    overlap_prev = sorted(list(set(numbers) & set(before_drawn_numbers)))
    print(f"3.1 前一次重號: {overlap_prev}")

    # 4. 顏色分佈
    reds = [n for n in sorted_nums if n in red_balls]
    blues = [n for n in sorted_nums if n in blue_balls]
    greens = [n for n in sorted_nums if n in green_balls]
    print(f"4. 顏色分佈: 紅{reds}, 藍{blues}, 綠{greens}")

    # 5. 總和
    sum_min_6 = sum(sorted_nums[:6])
    sum_max_6 = sum(sorted_nums[1:])
    print(f"5. 號碼總和: 最小6個({sum_min_6}), 最大6個({sum_max_6})")

    # 6. 區間分佈
    groups = {
        '1-10': 0, '11-20': 0, '21-30': 0, '31-40': 0, '41-49': 0
    }
    for n in sorted_nums:
        if 1 <= n <= 10: groups['1-10'] += 1
        elif 11 <= n <= 20: groups['11-20'] += 1
        elif 21 <= n <= 30: groups['21-30'] += 1
        elif 31 <= n <= 40: groups['31-40'] += 1
        elif 41 <= n <= 49: groups['41-49'] += 1
    
    active_groups = [k for k, v in groups.items() if v > 0]
    print(f"6. 區間分佈: {groups} (共 {len(active_groups)} 組有號碼)")

    # 7. 單雙分佈
    odds = [n for n in sorted_nums if n % 2 != 0]
    evens = [n for n in sorted_nums if n % 2 == 0]
    print(f"7. 單雙分佈: 單{odds}, 雙{evens}")

    # 8. 號碼差距
    diff = max(numbers) - min(numbers)
    print(f"8. 號碼差距: {diff} (最大{max(numbers)} - 最小{min(numbers)})")


if __name__ == "__main__":
    numbers, count = pick_marksix_numbers()
    print(f"經過 {count} 次抽選，選取的六合彩號碼是:", sorted(numbers))  # 將號碼排序後輸出
    print_analysis(numbers)