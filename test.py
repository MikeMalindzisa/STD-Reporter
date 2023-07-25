import datetime

prefix = "0032023"
current_datetime = datetime.datetime.now()
unique_id = prefix + current_datetime.strftime("%S%H%M%m%d%y")

# Pad the unique ID to a length of 14 characters
unique_id = unique_id.ljust(14, "0")

print(unique_id)
