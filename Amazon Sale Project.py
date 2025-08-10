# -*- coding: utf-8 -*-
"""
Created on Sun Aug 10 09:04:42 2025

@author: MinhTok1oPC

Delivered Question:
    - How many sales have amount over 1000?
    - How many sales have Category 'Tops' and quantities of 3?
    - Average Amount by Category and Status, or Category and Fulfilment?
    - Total Sales by Shipment Type and Fulfilment?
"""
#============================================
# Get Data
#============================================
import pandas as pd

sales_data = pd.read_excel('sales_data.xlsx')

sales_data.info()
sales_data.describe()
print(sales_data.columns)

#============================================
# Clean Data
#============================================

# Missing Value
sales_data.isnull()
print(sales_data.isnull().sum())

# Drop any 'nan' variable in Amount column
sales_data_cleaned = sales_data.dropna(subset = ['Amount'])

# Filtering a subset based on Category Top
category_tops = sales_data_cleaned[sales_data_cleaned['Category'] == 'Top']
print(category_tops)
# and have quantity of 3
multi_filter = sales_data_cleaned[(sales_data_cleaned['Category'] == 'Top') & (sales_data_cleaned['Qty'] == 3)]
print(multi_filter)                                                               

# Subset where Amount > 1000
high_amount = sales_data[sales_data['Amount'] > 1000]
print(high_amount)

#============================================
# Aggregating Data
#============================================
category_totals = sales_data_cleaned.groupby('Category', as_index = False)['Amount'].sum()
category_totals = category_totals.sort_values('Amount', ascending = False)

# Avg amount by Category and Fulfilment, Category and Status
fulfilment_avg = sales_data_cleaned.groupby(['Category', 'Fulfilment'], as_index = False)['Amount'].mean()
fulfilment_avg = fulfilment_avg.sort_values('Amount', ascending = False)
status_avg = sales_data_cleaned.groupby(['Category', 'Status'], as_index = False)['Amount'].mean()
status_avg = status_avg.sort_values('Amount', ascending = False)

total_sales_shipandfulfil = sales_data.groupby(['Courier Status', 'Fulfilment'], as_index = False)['Amount'].sum()
total_sales_shipandfulfil = total_sales_shipandfulfil.sort_values('Amount', ascending = False)


#============================================
# Export Data
#============================================

total_sales_shipandfulfil.rename(columns={'Courier Status' : 'Shipment'}, inplace = True)

high_amount.to_excel('sales_have_over_1000_amount.xlsx', index = False)
multi_filter.to_excel('sales_of_tops_with_3_quantity.xlsx', index = False)
fulfilment_avg.to_excel('average_amounts_by_fulfilment.xlsx', index = False)
status_avg.to_excel('average_amounts_by_status.xlsx', index = False)
total_sales_shipandfulfil.to_excel('total_sales_by_shipment_and_fultiment.xlsx', index = False)


