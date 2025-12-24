import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt

file_path = 'podatki/Rizana_Zaledje_INCA_dnev_2020_2021_N1.xlsx'
df = pd.read_excel(file_path, sheet_name='Izbrane', skiprows=[0, 1, 3, 4], usecols="A,C:IW")

df["Datum"] = pd.to_datetime(df["Datum"], format="%m/%d/%y")
df = df.set_index('Datum')


print(f"Date range: {df.index.min()} to {df.index.max()}")

# 2. LIMIT NUMBER OF ROWS (DATES) - NEW SECTION

# Select specific date range
start_date = '2020-06-01'  # Change to your desired start date
end_date = '2020-06-30'    # Change to your desired end date
date_mask = (df.index >= start_date) & (df.index <= end_date)
df_limited = df[date_mask]

print(f"\nLimited to date range: {df_limited.index.min()} to {df_limited.index.max()}")
print(f"Number of days shown: {len(df_limited)}")

# 3. Calculate total values for each column and select top N
top_n = 20  # Change this to plot more or fewer columns
col_totals = df_limited.sum().sort_values(ascending=False)
top_cols = col_totals.head(top_n).index.tolist()

print(f"\n=== Top {top_n} Columns by Total Value (in selected date range) ===")
print(f"{'Rank':<5} {'Column':<10} {'Total Value':<15}")
print("-" * 35)
for i, (col, total) in enumerate(col_totals.head(top_n).items(), 1):
    print(f"{i:<5} {str(col):<10} {total:>15.0f}")

# 4. Create the stacked bar chart
df_top = df_limited[top_cols]

fig, ax = plt.subplots(figsize=(16, 8))

# Create stacked bars
bottom_vals = np.zeros(len(df_top))

# Plot each column
for column in df_top.columns:
    ax.bar(df_top.index, df_top[column], bottom=bottom_vals, width=0.8)
    bottom_vals += df_top[column].values

# 5. Customize the plot
ax.set_title(f'Top {top_n} Columns by Total Value\n{df_top.index.min().date()} to {df_top.index.max().date()} ({len(df_top)} days)', 
             fontsize=14, fontweight='bold')
ax.set_xlabel('Date', fontsize=12)
ax.set_ylabel('Total Value', fontsize=12)

# Format x-axis dates based on number of days shown
num_days = len(df_top)
if num_days <= 30:
    # If <= 30 days, show every day
    ax.set_xticks(df_top.index)
    ax.set_xticklabels(df_top.index.strftime('%Y-%m-%d'), rotation=45, fontsize=9)
elif num_days <= 90:
    # If 31-90 days, show every 3rd day
    ax.set_xticks(df_top.index[::3])
    ax.set_xticklabels(df_top.index[::3].strftime('%Y-%m-%d'), rotation=45, fontsize=9)
else:
    # If > 90 days, show every 7th day (weekly)
    ax.set_xticks(df_top.index[::7])
    ax.set_xticklabels(df_top.index[::7].strftime('%Y-%m-%d'), rotation=45, fontsize=9)

# Add grid for better readability
ax.grid(axis='y', alpha=0.3, linestyle='--')

# 8. Adjust layout and show plot
plt.tight_layout()
plt.show()

# 9. Optional: Create a summary table
print("\n=== Summary Statistics for Top Columns (in selected date range) ===")
summary_df = pd.DataFrame({
    'Column': top_cols,
    'Total': [col_totals[col] for col in top_cols],
    'Average': [df_top[col].mean() for col in top_cols],
    'Max': [df_top[col].max() for col in top_cols],
    'Min': [df_top[col].min() for col in top_cols],
    'Days > 0': [(df_top[col] > 0).sum() for col in top_cols],
    '% Days > 0': [((df_top[col] > 0).sum() / len(df_top) * 100) for col in top_cols]
})

print(summary_df.to_string(index=False, float_format='{:,.1f}'.format))