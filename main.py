import pandas as pd


def check_stores(data_path):
    data = pd.read_excel(data_path)

    unique_months = data.Period.unique()
    last_month_data = data[data.Period == unique_months[-1]]

    def classify(row):
        return 'Y' if (row['sum'] / row['count']) >= 0.8 else 'N'

    last_month_data['Has_Stock'] = (last_month_data['CLOSING STOCK'] > 0) | (last_month_data['SALES'] > 0)
    last_month_agg_data = last_month_data.groupby(['Store Code', 'Store Name'])['Has_Stock'].agg(['sum', 'count']).reset_index()
    last_month_agg_data['Class'] = last_month_agg_data.apply(classify, axis=1)

    return last_month_agg_data[last_month_agg_data.Class == 'N'][['Store Code', 'Store Name']]


if __name__ == "__main__":
    data_path = "data.xlsx"
    stores = check_stores(data_path)
    print(stores)
