import pandas as pd

# Load the data
df_sales = pd.read_excel('C:\\Users\\Lenovo\\Downloads\\shitjet orafol.xlsx')
df_purchases = pd.read_excel('C:\\Users\\Lenovo\\Downloads\\shitjet orafol.xlsx')

# Drop the first row
df_sales = df_sales.iloc[1:]

# Save the total row to a separate variable and drop it from the dataframe
total_sales_row = df_sales.iloc[-1]
df_sales = df_sales.iloc[:-1]

# Exclude rows where Unnamed: 2 is either "Roll" or "CP"
df_sales = df_sales[~df_sales['Unnamed: 2'].isin(['Roll', 'CP'])]


# Take a look at the first few rows of the data
print(df_sales.head())
print(df_purchases.head())

df_sales['01.01.2023-31.05.2023'] = pd.to_numeric(df_sales['01.01.2023-31.05.2023'], errors='coerce')


print ("TOP DHE BOTTOM PRODUKTET /n")
# Top 10 products in terms of sales
top_sales = df_sales.groupby('Unnamed: 1')['01.01.2023-31.05.2023'].sum().sort_values(ascending=False).head(11)
print(top_sales)

# Bottom 10 products in terms of sales
bottom_sales = df_sales.groupby('Unnamed: 1')['01.01.2023-31.05.2023'].sum().sort_values().head(11)
print(bottom_sales)

# Repeat the same for purchases

print("Quantiles")
# Quantiles of the data
print(df_sales['01.01.2023-31.05.2023'].quantile([0.25, 0.5, 0.75]))
 

print("Sales Concetration")
# Sales Concetration 
total_sales = df_sales['01.01.2023-31.05.2023'].sum()
top_10_sales = df_sales.groupby('Unnamed: 1')['01.01.2023-31.05.2023'].sum().sort_values(ascending=False).head(10).sum()
sales_concentration = (top_10_sales / total_sales) * 100
print(f"Top 10 products account for {sales_concentration}% of total sales.")

print("Statistics")
# Mean sales
mean_sales = df_sales['01.01.2023-31.05.2023'].mean()
print(f"Mean Sales: {mean_sales}")

# Median sales
median_sales = df_sales['01.01.2023-31.05.2023'].median()
print(f"Median Sales: {median_sales}")

# Mode sales
mode_sales = df_sales['01.01.2023-31.05.2023'].mode()
print(f"Mode Sales: {mode_sales}")

# Standard deviation of sales
std_sales = df_sales['01.01.2023-31.05.2023'].std()
print(f"Standard Deviation of Sales: {std_sales}")

# Variance of sales
var_sales = df_sales['01.01.2023-31.05.2023'].var()
print(f"Variance of Sales: {var_sales}")

# Minimum sales
min_sales = df_sales['01.01.2023-31.05.2023'].min()
print(f"Minimum Sales: {min_sales}")

# Maximum sales
max_sales = df_sales['01.01.2023-31.05.2023'].max()
print(f"Maximum Sales: {max_sales}")

print("Sales Variability")
# Sales Variability 
sales_variability = df_sales.groupby('Unnamed: 1')['01.01.2023-31.05.2023'].std() / df_sales.groupby('Unnamed: 1')['01.01.2023-31.05.2023'].mean()
print(sales_variability.sort_values(ascending=False))
 

print("Skewness and Kurtosis")
#Skewness and Kurtosis
print(df_sales['01.01.2023-31.05.2023'].skew())
print(df_sales['01.01.2023-31.05.2023'].kurt())

print("Pareto 80/20")
#Pareto 80/20
# First, let's calculate the total sales and the product share
total_sales = df_sales['01.01.2023-31.05.2023'].sum()
product_share = df_sales.groupby('Unnamed: 1')['01.01.2023-31.05.2023'].sum().sort_values(ascending=False)

# Then let's calculate the cumulative sales share
cumulative_sales_share = product_share.cumsum() / total_sales

# Let's print out the cumulative sales share for debugging
print(cumulative_sales_share)

# Now, let's find the point where cumulative sales share exceeds 80%
over_80 = cumulative_sales_share[cumulative_sales_share > 0.8]

if not over_80.empty:
    # The first index in this series is the product that pushed cumulative sales over 80%
    first_over_80 = over_80.index[0]

    # Let's find out how many products contribute up to 80% of sales
    products_upto_80 = product_share.index.get_loc(first_over_80)

    # Let's print out the product that pushed sales over 80%
    print(f"Product '{first_over_80}' pushed sales over 80%.")
else:
    # If there is no product that pushes sales over 80%, all products contribute up to 80%
    products_upto_80 = len(product_share)

percentage_of_products = products_upto_80 / len(product_share) * 100

print(f"{percentage_of_products}% of products make up 80% of sales.")

print("Sales Velocity")
#Sales Velocity Monthly
# Calculate monthly sales velocity
df_sales['sales_velocity'] = df_sales['01.01.2023-31.05.2023'] / 5  # monthly sales velocity

# Display top N products with highest sales velocity
N = 10  # Change this to the number of products you want to display
top_N_velocity_products = df_sales.nlargest(N, 'sales_velocity')[['Unnamed: 1', 'sales_velocity']]
print(top_N_velocity_products)





