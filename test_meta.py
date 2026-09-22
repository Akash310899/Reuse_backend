import sys
# Just a quick check to ensure length of new_headers and old_headers
new_headers = ["New Spec", "New WL", "New WR", "New CODE", "New LL", "New LR", "New Flat", "New Location", "New Part", "New QTY", "New Unit"]
old_headers = ["Old Spec", "Old WL", "Old WR", "Old CODE", "Old LL", "Old LR", "Old Location", "Old QTY", "Old Unit Area"]
print("New len:", len(new_headers))
print("Old len:", len(old_headers))
print("Remainder Used index:", len(new_headers) + len(old_headers) + 1)
