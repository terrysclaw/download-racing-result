# using random sampling to generate numbers respecting the criteria of unique bankers within a decade, and unique feet not overlapping with the bankers - doing this repeatedly for each defined decade range.

import random

columns = [range(1, 11), range(11, 21), range(21, 31), range(31, 41), range(41, 50)]
all_nums = range(1, 50)

for column in columns:
    column = list(column)
    random.shuffle(column)
    bankers = random.sample(column, 3)

    remaining = list(set(all_nums) - set(column))
    random.shuffle(remaining)
    feet = random.sample(remaining, 6)

    print(bankers, feet)
