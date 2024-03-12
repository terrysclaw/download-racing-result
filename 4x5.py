# using random sampling to generate numbers respecting the criteria of unique bankers within a decade, and unique feet not overlapping with the bankers - doing this repeatedly for each defined decade range.

import random

decades = [range(1, 11), range(11, 21), range(21, 31), range(31, 41), range(41, 50)]
all_nums = range(1, 50)

for decade in decades:
    bankers = random.sample(decade, 3)

    feet = random.sample(list(set(all_nums) - set(bankers)), 5)

    print(bankers, feet)
