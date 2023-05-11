import pandas as pd

def analyze_excel_data(file_path):
    # Read Excel file as a DataFrame
    df = pd.read_excel(file_path)

    # Find absolute value of negative numbers in SALES column
    df['SALES'] = df['SALES'].abs()

    # Convert the Period column to datetime type
    df['Period'] = pd.to_datetime(df['Period'])

    # Filter the dataset for the required periods
    unique_months = df['Period'].dt.month.unique()
    required_months = unique_months[-5:]  # Filter the last 5 months
    required_df = df[df['Period'].dt.month.isin(required_months)]

    # Compute sales value for each store
    required_df['Sales Value'] = required_df['SALES'] * required_df['PRICE']

    # Group data by Store Code and Period, and calculate total sales value for each store within each month
    total_sales_value = required_df.groupby(['Store Code', 'Period']).agg(
        Sum_Sales_Value=('Sales Value', 'sum')).reset_index()
    print(total_sales_value)

    # Group data by Store Code and calculate the average of total sales value for each store
    previous_months = total_sales_value[total_sales_value['Period'].dt.month.isin(unique_months[-5:-1])]
    average_sales_value = previous_months.groupby('Store Code')['Sum_Sales_Value'].mean()
    print(average_sales_value)

    # Calculate upper and lower thresholds
    upper_threshold = average_sales_value * 1.3
    lower_threshold = average_sales_value * 0.7

    # Check if total sales value in last month breaches the thresholds
    last_month_total_sales_value = total_sales_value[total_sales_value['Period'].dt.month == unique_months[-1]]
    last_month_sales_value = last_month_total_sales_value.reset_index()
    last_month_sales_value['Status'] = last_month_sales_value.apply(lambda row: 'Breach' if
    row['Store Code'] in upper_threshold and
    (row['Sum_Sales_Value'] > upper_threshold[row['Store Code']] or
     row['Sum_Sales_Value'] < lower_threshold[row['Store Code']])
    else 'Within the limit', axis=1)

    return last_month_sales_value


if __name__ == "__main__":
    file_path = "data.xlsx"
    last_month_sales_value_by_store = analyze_excel_data(file_path)
    print(last_month_sales_value_by_store)
