import random

decades = [range(1,11), range(11,21), range(21,31), range(31,41), range(41,50)]

numbers_decade = random.choice(decades)
numbers = random.sample(numbers_decade, 4)

all_nums = []
for decade in decades:
    all_nums.extend(list(decade))

foots = random.sample(list(set(all_nums)-set(numbers)), 5)

print(numbers)
print(foots)
