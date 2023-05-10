import pandas as pd


def analyze_excel_data(file_path):
    # Read Excel file as a DataFrame
    df = pd.read_excel(file_path)

    # Find absolute value of negative numbers in SALES column
    df['SALES'] = df['SALES'].abs()

    # Convert the Period column to datetime type
    df['Period'] = pd.to_datetime(df['Period'])

    # Filter the dataset for the required periods
    required_periods = [11, 12, 1, 2, 3]
    required_df = df[df['Period'].dt.month.isin(required_periods)]

    # Compute sales value for each store
    required_df['Sales Value'] = required_df['SALES'] * required_df['PRICE']

    # Group data by Store Code and Period, and calculate total sales value for each store within each month
    total_sales_value = required_df.groupby(['Store Code', 'Period'])['Sales Value'].sum()
    #print(total_sales_value)

    # Group data by Store Code and calculate the average of total sales value for each store
    average_sales_value = total_sales_value.groupby('Store Code').mean()
    #print(average_sales_value)

    # Calculate upper and lower thresholds
    upper_threshold = average_sales_value * 1.3
    lower_threshold = average_sales_value * 0.7

    # Check if total sales value in March breaches the thresholds
    march_sales = total_sales_value[total_sales_value.index.get_level_values('Period').month == 3]
    march_sales = march_sales.reset_index()
    march_sales['Status'] = march_sales.apply(lambda row: 'Breach' if
                                              row['Sales Value'] > upper_threshold[row['Store Code']] or
                                              row['Sales Value'] < lower_threshold[row['Store Code']]
                                              else 'Within the limit', axis=1)

    return march_sales


if __name__ == "__main__":
    file_path = "C:\\Users\\Flovianey ELION\\PycharmProjects\\Stock_checker\\data.xlsx"
    march_sales_by_store = analyze_excel_data(file_path)
    print(march_sales_by_store)
