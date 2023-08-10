

# In[127]:


import pandas as pd
import requests


# In[177]:


df=pd.read_excel(r"C:\Users\Tanishqa\Desktop\project3\Phlebo data-manoj.xlsx")
df['Pincode'] = df['Pincode'].fillna(-1)

# Convert the 'Pincode' column to integer data type
df['Pincode'] = df['Pincode'].astype(int)
df['Pincode'] = df['Pincode'].astype(str)
list1=df["Pincode"]
list2=df["City"]

list2[1]


for i in range(0,len(list1)):
    PINCODE=list1[i]
    API_ENDPOINT = f'https://api.postalpincode.in/pincode/{PINCODE}'
    response = requests.get(API_ENDPOINT)
    if response.status_code == 200:
        print("ok")
        data = response.json()  # Assuming the API returns JSON data.
        if data[0]["Status"]=="Success":
             for post_office in data[0]['PostOffice']:
                    area_name = post_office['District']
                    df["Area"][i]=area_name
            
        else:
            print("Not found")
      



df.to_excel(r"C:\Users\Tanishqa\Desktop\project3\Phlebo_data_manoj.xlsx", index=False)

print("Area names updated successfully!")


print(df)




