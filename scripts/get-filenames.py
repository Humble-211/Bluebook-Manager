import os
import re

folder_path = r"X:\819 (Golden Windows)\JPG"

numbers = set()

for file in os.listdir(folder_path):
    name, ext = os.path.splitext(file)

    # extract numeric prefix at the start of filename
    match = re.match(r"(\d+)", name)
    if match:
        numbers.add(match.group(1))

# sort numerically
result_list = sorted(numbers, key=int)

# join with commas
result = ",".join(result_list)

print(result)