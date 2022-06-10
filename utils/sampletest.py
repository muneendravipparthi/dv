# Import Pandas package
import pandas as pd

# Create a sample dataframe
df = pd.DataFrame({'num_posts': [4, 6, 3, 9, 1, 14, 2, 5, 7, 2],
                   'date': ['2020-08-09', '2020-08-25',
                            '2020-09-05', '2020-09-12',
                            '2020-09-29', '2020-10-15',
                            '2020-11-21', '2020-12-02',
                            '2020-12-10', '2020-12-18']})

# Convert the date to datetime64
df['date'] = pd.to_datetime(df['date'], format='%Y-%m-%d')

# Filter data between two dates
filtered_df = df.query("date >= '2020-08-01' \
                       and date < '2020-09-01'")

# Display
print(filtered_df)