import pandas as pd


def check_stores(data_path):
    data = pd.read_excel(data_path)

    march_data = data[data['Period'].dt.month == 3]
    print(march_data.columns)

    def has_stock(row):
        return 1 if (row['CLOSING STOCK'] > 0) or (row['SALES'] > 0) else 0

    march_data['Has_Stock'] = march_data.apply(has_stock, axis=1)

    march_agg_data = march_data.groupby(['Store Code', 'Store Name'])['Has_Stock'].agg(['sum', 'count']).reset_index()
    print(march_agg_data.columns)

    def classify(row):
        return 'Y' if (row['sum'] / row['count']) >= 0.8 else 'N'

    march_agg_data['Class'] = march_agg_data.apply(classify, axis=1)

    return march_agg_data[march_agg_data.Class == 'N'][['Store Code', 'Store Name']]


if __name__ == "__main__":
    data_path = "C:\\Users\\Flovianey ELION\\PycharmProjects\\Stock_checker\\data.xlsx"
    stores = check_stores(data_path)
    print(stores)
