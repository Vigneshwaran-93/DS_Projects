# -*- coding: utf-8 -*-

# # Reading files
# import os
# path="data"
# os.listdir(path)

#importing necessary libraries 
import pandas as pd

"""### Company X - Order Report.xlsx"""

comp_ord=pd.read_excel("data/Company X - Order Report.xlsx")

comp_ord

comp_ord.isnull().sum()

comp_ord.info()

comp_ord.describe()

[print(comp_ord[i].nunique()) for i in comp_ord.columns]

"""#Analysis 

1. There are toal of 400 rows of which we can get the following 
  - There are total of 124 orders 
  - There are total of 65 products
  - The maximum weight of a single product is 8 Nos
"""



"""### Company X - SKU Master.xlsx"""

comp_sku=pd.read_excel("data/Company X - SKU Master.xlsx")

comp_sku

comp_sku.shape

comp_sku.info()

comp_sku.describe()

comp_sku.isnull().sum()

[print(comp_sku[i].nunique()) for i in comp_sku.columns]

comp_sku.drop_duplicates(inplace=True)

comp_sku

"""#Analysis 

The total unique products matched with the no. of items with Orders Table 
"""



"""### Company X - Pincode Zones.xlsx"""

comp_pin_zone=pd.read_excel("data/Company X - Pincode Zones.xlsx")

comp_pin_zone

comp_pin_zone.isnull().sum()

comp_pin_zone.info()

[print(comp_pin_zone[i].nunique()) for i in comp_pin_zone.columns]

comp_pin_zone.drop_duplicates(inplace=True)

comp_pin_zone.shape

"""#Analysis 

There are a total of 108 unique pincodes to be serviced by the courier company 
"""



"""### Courier Company - Invoice.xlsx"""

cour_invoc=pd.read_excel("data/Courier Company - Invoice.xlsx")

cour_invoc

cour_invoc.isnull().sum()

cour_invoc.info()

cour_invoc.describe()

cour_invoc.shape

[print(i, ": ",cour_invoc[i].nunique()) for i in cour_invoc.columns]

cour_invoc

"""# Analysis 
     

*  The orderid is matchign with the company order Id's
*  The order pincode data is avilable with the Courier company data, It has to be mapped to company's orderid table for calulations. 
*  **There is a possibility that the courier company may have inccorect customer pincode, which cannot be verfierd as we dont have customer pincode from company's data.**

###  'Courier Company - Rates.xlsx',
"""

cour_rates=pd.read_excel("data/Courier Company - Rates.xlsx")

cour_rates

cour_rates=cour_rates.transpose()

cour_rates.shape

# Creating Additional comumns

cour_rates.rename(columns = {0:'rate'}, inplace = True)

cour_rates

cour_rates["desc"]=cour_rates.index

cour_rates

cour_rates["type"]=cour_rates["desc"].apply(lambda x:x[:3])

cour_rates["zone_code"]=cour_rates["desc"].apply(lambda x:x[4:5])

cour_rates["rate_type"]=cour_rates["desc"].apply(lambda x:x[6:])

cour_rates

"""# Analysis 
There are 5 zones with fixed and additional charges, also they have seperate charges for Forword and Return orders. 
"""



"""## Resultant Table
The result table sould have the following columns
 *   ● Order ID
 *  ● AWB Number
 *   ● Total weight as per X (KG)
 *   ● Weight slab as per X (KG)
 *   ● Total weight as per Courier Company (KG)
 *   ● Weight slab charged by Courier Company (KG)
 *   ● Delivery Zone as per X
 *   ● Delivery Zone charged by Courier Company
 *   ● Expected Charge as per X (Rs.)
 *   ● Charges Billed by Courier Company (Rs.)
 *   ● Difference Between Expected Charges and Billed Charges (Rs.)
"""

#Creating a new DataFrame for result data 
df=pd.DataFrame()

# Creating  Order ID and  AWB Number columns
df["Order ID"]=cour_invoc["Order ID"]
df["AWB Number"]=cour_invoc["AWB Code"]

# Creating  Total weight as per X KG column
def total_comp_wt(orderid):
  temp_df= comp_ord[comp_ord["ExternOrderNo"]==orderid]
  temp_df=pd.merge(temp_df,comp_sku,on='SKU',how='left')
  temp_df["Total_wt"]=temp_df["Weight (g)"]*temp_df["Order Qty"]
  return ( temp_df["Total_wt"].sum())/1000


df["Total weight as per X KG"]=df["Order ID"].apply(total_comp_wt)
df

# Creating  Weight slab as per X KG column

def get_wt_slab(weight):
 weight=weight*1000
 ct=0
 while(weight >500):
   ct+=1
   
   weight=weight-500
 ct=ct+1
 return (ct*500)/1000

df["Weight slab as per X KG"]=df["Total weight as per X KG"].apply(get_wt_slab)
df

# Creating  Total weight as per Courier Company KG column

def total_cour_wt(orderid):
  k=cour_invoc[cour_invoc["Order ID"] == orderid]["Charged Weight"].values
  return k[0]

df["Total weight as per Courier Company KG"]=df["Order ID"].apply(total_cour_wt)
df

# Creating  Weight slab charged by Courier Company KG column

df[" Weight slab charged by Courier Company KG"]=df["Total weight as per Courier Company KG"].apply(get_wt_slab)
df

comp_pin_zone=comp_pin_zone.rename(columns={"Zone":"Comap_Zone"})

comp_pin_zone.columns

# Creating  Company_Zone column

def company_zone(orderid):
  pincode=cour_invoc[cour_invoc["Order ID"]==orderid]["Customer Pincode"].iloc[0]
  return comp_pin_zone[comp_pin_zone["Customer Pincode"]==pincode ]["Comap_Zone"].iloc[0]

df["Company_Zone"]=df["Order ID"].apply(company_zone)
df

# Creating  Courier_Zone column

def courier_zone(orderid):
  return cour_invoc[cour_invoc["Order ID"]==orderid]["Zone"].iloc[0]

df["Courier_Zone"]=df["Order ID"].apply(company_zone)
df

# Creating  Expected Charge as per X (Rs.) column

def get_company_charges(orderid):
  order_type=cour_invoc[cour_invoc["Order ID"]==orderid]["Type of Shipment"].iloc[0]
  company_wt_slab=df[df["Order ID"]==orderid]["Weight slab as per X KG"].iloc[0]
  company_zone=df[df["Order ID"]==orderid]["Company_Zone"].iloc[0]
  weight_sec=company_wt_slab/0.5
  
  fwd_base_chrg=cour_rates[(cour_rates["type"]=="fwd" ) & (cour_rates["zone_code"]==company_zone) & (cour_rates["rate_type"]=="fixed")]["rate"].iloc[0]
  fwd_addl_chrg=cour_rates[(cour_rates["type"]=="fwd" ) & (cour_rates["zone_code"]==company_zone) & (cour_rates["rate_type"]=="additional")]["rate"].iloc[0]

  amount=0
  if order_type=="Forward charges":
    
    amount+=fwd_base_chrg
    amount+=(fwd_addl_chrg*(weight_sec-1))
    return amount
  elif order_type=="Forward and RTO charges":
    rto_base_chrg=cour_rates[(cour_rates["type"]=="rto" ) & (cour_rates["zone_code"]==company_zone) & (cour_rates["rate_type"]=="fixed")]["rate"].iloc[0]
    rto_addl_chrg=cour_rates[(cour_rates["type"]=="rto" ) & (cour_rates["zone_code"]==company_zone) & (cour_rates["rate_type"]=="additional")]["rate"].iloc[0]

    amount+=fwd_base_chrg
    amount+=(fwd_addl_chrg*(weight_sec-1))
    amount+=rto_base_chrg
    amount+=(rto_addl_chrg*(weight_sec-1))
    return amount

df["Expected Charge as per X (Rs.)"]=df["Order ID"].apply(get_company_charges)
df

# Creating Charges Billed by Courier Company (Rs.) column

def get_courier_charges(orderid):
  order_type=cour_invoc[cour_invoc["Order ID"]==orderid]["Type of Shipment"].iloc[0]
  Courier_wt_slab=df[df["Order ID"]==orderid][" Weight slab charged by Courier Company KG"].iloc[0]
  Courier_zone=df[df["Order ID"]==orderid]["Courier_Zone"].iloc[0]
  weight_sec=Courier_wt_slab/0.5
  
  fwd_base_chrg=cour_rates[(cour_rates["type"]=="fwd" ) & (cour_rates["zone_code"]==Courier_zone) & (cour_rates["rate_type"]=="fixed")]["rate"].iloc[0]
  fwd_addl_chrg=cour_rates[(cour_rates["type"]=="fwd" ) & (cour_rates["zone_code"]==Courier_zone) & (cour_rates["rate_type"]=="additional")]["rate"].iloc[0]

  amount=0
  if order_type=="Forward charges":
    
    amount+=fwd_base_chrg
    amount+=(fwd_addl_chrg*(weight_sec-1))
    return amount
  elif order_type=="Forward and RTO charges":
    rto_base_chrg=cour_rates[(cour_rates["type"]=="rto" ) & (cour_rates["zone_code"]==Courier_zone) & (cour_rates["rate_type"]=="fixed")]["rate"].iloc[0]
    rto_addl_chrg=cour_rates[(cour_rates["type"]=="rto" ) & (cour_rates["zone_code"]==Courier_zone) & (cour_rates["rate_type"]=="additional")]["rate"].iloc[0]

    amount+=fwd_base_chrg
    amount+=(fwd_addl_chrg*(weight_sec-1))
    amount+=rto_base_chrg
    amount+=(rto_addl_chrg*(weight_sec-1))
    return amount

df["Charges Billed by Courier Company (Rs.)"]=df["Order ID"].apply(get_courier_charges)
df

# Creating  Difference Between Expected Charges and Billed Charges (Rs.) column

df["Difference Between Expected Charges and Billed Charges (Rs.)"]=df["Expected Charge as per X (Rs.)"]-df["Charges Billed by Courier Company (Rs.)"]

# Checking the total difference amount
df["Difference Between Expected Charges and Billed Charges (Rs.)"].sum()



"""### Summary Table """

#Calculating necessory values 

total_corr_amount = df[df['Difference Between Expected Charges and Billed Charges (Rs.)'] == 0]['Expected Charge as per X (Rs.)'].sum()

total_overcharging_amount = df[df['Difference Between Expected Charges and Billed Charges (Rs.)'] < 0]['Difference Between Expected Charges and Billed Charges (Rs.)'].sum() * -1

total_undercharging_amount = df[df['Difference Between Expected Charges and Billed Charges (Rs.)'] > 0]['Difference Between Expected Charges and Billed Charges (Rs.)'].sum()

count_corr_charged = len(df[df['Difference Between Expected Charges and Billed Charges (Rs.)'] == 0])

count_over_charged = len(df[df['Difference Between Expected Charges and Billed Charges (Rs.)'] < 0])

count_under_charged = len(df[df['Difference Between Expected Charges and Billed Charges (Rs.)'] > 0])

# Create summary table

summary_table = pd.DataFrame({
    'Count': [count_corr_charged, count_over_charged, count_under_charged],
    'Amount (Rs.)': [total_corr_amount, total_overcharging_amount, total_undercharging_amount]
}, index=['Total Orders - Correctly Charged', 'Total Orders - Over Charged', 'Total Orders - Under Charged'])

# Display summary table
print(summary_table)



"""# Exporing to Excel """

import pandas as pd

with pd.ExcelWriter('data/output.xlsx') as writer:
    
    summary_table.to_excel(writer, sheet_name='Summary', index=True)
    df.to_excel(writer, sheet_name='Calculations', index=False)

