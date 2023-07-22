import pandas as pd
import numpy as np
from datetime import datetime  , timedelta
import datetime as dt
#---------------------IMP ---------- CHaNGE DATE ALSO ------------
# reading file 
data=pd.read_excel(r"C:\Users\Tanishqa\Desktop\Partner\file.xlsx")

#setting up date column 
date_column = "collection_date"
#current date string
date_string="2023-07-22"
#changing it to datetime format 
current_date =np.datetime64(date_string)
# finding weeks number 
data["week"]= ((current_date - data[date_column]).dt.days )//7+1 
output_file_path = r"C:\Users\Tanishqa\Desktop\Partner\file.xlsx"
data.to_excel(output_file_path)



# chossing data less than or equal to 8 weeks 
filtered_data = data[data['week'] <= 8]

# usiing pivot_table of pandas 
pivot_table = pd.pivot_table(filtered_data, index=['booking_medium', 'center_name'], columns=['week'],  values='booking_id', aggfunc="sum",fill_value=0)
# setting up column names
new_columns = {col: f"Week {col}" for col in pivot_table.columns if isinstance(col, int)}
# renaming columns names
pivot_table = pivot_table.rename(columns=new_columns)
#finding grand total at the end 
pivot_table['Grand Total'] = pivot_table.sum(axis=1)
#finding max values 
max_values = pivot_table.drop('Grand Total', axis=1).max(axis=1)
# finding min values
min_values = pivot_table.drop('Grand Total', axis=1).apply(lambda x: x[x > 0].min(), axis=1)

booking_medium_sum = pivot_table.groupby('booking_medium').sum()
#adding min and max column 
pivot_table["max values"]=max_values
pivot_table["min values"]=min_values
#replacing 0 with ""
pivot_table = pivot_table.replace(0, '')

#pivot_table=pivot_table._append(booking_medium_sum)
print(booking_medium_sum)
print("done, please check your file ")
# exporting to excel file 
output_file_path = r"C:\Users\Tanishqa\Desktop\Partner\21th_july.xlsx"
pivot_table.to_excel(output_file_path)






