from datetime import datetime

# Convert the string '12-15-2023' into a date
date_string = '12-15-2023'
date_format = '%m-%d-%Y'
desired_date = datetime.strptime(date_string, date_format)

# Get the current date
current_date = datetime.now()

# Calculate the time difference in days
time_difference = desired_date - current_date

# Extract the number of days from the timedelta
days_difference = time_difference.days

print(f"The difference in days is: {days_difference} days")

