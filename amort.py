#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
from datetime import datetime


# In[2]:

# get data from cmd

file_path = input("Enter file path: ").strip()

def parse_rate(s: str) -> float:
    """Parse interest input like '14.5%', '14.5', or '0.145' to a decimal float (e.g. 0.145)."""
    s = s.strip()
    if not s:
        raise ValueError("empty rate")
    if s.endswith('%'):
        return float(s[:-1]) / 100.0
    val = float(s)
    # Heuristic: if user enters a number > 1 (e.g. 14.5) assume percent and divide by 100.
    return val / 100.0 if val > 1 else val


interest_rates = []
interest_dates = []

print("Enter entries. Leave the interest input blank to finish at any time.\n")

# First entry: only interest rate
while True:
    first_rate = input("First entry — interest rate (e.g. 0.145): ").strip()
    if first_rate == "":
        print("No entries provided. Exiting.")
        break
    try:
        r = parse_rate(first_rate)
    except Exception as e:
        print("Invalid interest rate:", e)
        continue
    interest_rates.append(r)
    break

# Subsequent entries: interest + date
entry_num = 2
while interest_rates:  # continue only if we got the first entry
    rate_in = input(f"Entry #{entry_num} — interest rate (blank to finish): ").strip()
    if rate_in == "":
        break
    try:
        r = parse_rate(rate_in)
    except Exception as e:
        print("Invalid interest rate:", e)
        continue

    # get a valid date
    while True:
        date_in = input(f"Entry #{entry_num} — date (YYYY-MM-DD) ").strip()
        try:
            d = datetime.strptime(date_in, "%Y-%m-%d").date()
            break
        except Exception as e:
            print("Invalid date format. Please try again.")

    interest_rates.append(r)
    interest_dates.append(d)

    entry_num += 1


#interest_rates = [0.145, 0.1425]

# add dates (date of starting interests is not needed)
#interest_effective = ['2025-01-01']


# In[3]:


#interest_dates = [datetime.strptime(s, "%Y-%m-%d").date() for s in interest_effective]


# In[4]:


# automated amort writer
#file_path = 'SPVL158.xlsx'

df = pd.read_excel(file_path)

print("Read Excel sheet")

# In[5]:


df_reversed = df[::-1]
df_reindexed = df_reversed.reset_index(drop=True)


# In[6]:


# all drawdowns, repayments and fees
col4 = df_reindexed['Statement Text']            # 4th col values
col10 = df_reindexed['Amount'].astype(float)  # may raise if conversion impossible

col4_str = np.asarray(col4, dtype=str)
col4_lower = np.char.lower(np.char.strip(col4_str))


mask = (np.char.find(col4_lower, 'drawdown') >= 0) & (col10 < 0)
drawdown_indices = np.where(mask)[0].tolist()

mask2 = ((col4 == 'Repayment') | (col4 == 'Repayment - Settle Capital') | (col4 == 'Repayment - Settle Interest')) & (col10 <0) & (col10 != 0)
repayment_indices = np.where(mask2)[0].tolist()


mask3 = (np.char.find(col4_lower, 'fee') >= 0) & (col4_lower != 'platform fee')
fee_indices = np.where(mask3)[0].tolist()


# In[7]:


dates = df_reindexed['Date']
date_first_drawdown = dates[drawdown_indices[0]].date() # time 00:00:00

last_date_index = max(max(drawdown_indices, default=0), max(repayment_indices, default = 0), max(fee_indices, default =0))
date_last = dates[last_date_index].date()


# In[8]:


def index_date_col(date1, drawdowns):
    dates = drawdowns['Date']
    indx = []
    for d in dates:
        indx.append((d.date()-date1).days)
    return indx
    


# In[9]:


drawdowns = df_reindexed.iloc[drawdown_indices]['Amount'].astype(int).values.tolist()
repayments = df_reindexed.iloc[repayment_indices]['Amount'].astype(int).values.tolist()
fees = df_reindexed.iloc[fee_indices]['Amount'].astype(int).values.tolist()


# In[10]:


date_range = pd.date_range(start=date_first_drawdown, end=date_last, freq='D')
df_return = pd.DataFrame({'Date': date_range})

df_return['Opening Balance'] = 0.0
df_return['Advances'] = 0.0
df_return['Repayments'] = 0.0
df_return['Interest Rate'] = 0.0
df_return['Interest'] = 0.0
df_return['Interest Charged']=0.0
df_return['Fees'] = 0.0
df_return['Closing Balance'] = 0.0


ix1 = index_date_col(date_first_drawdown, df_reindexed.iloc[drawdown_indices])
ix2 = index_date_col(date_first_drawdown, df_reindexed.iloc[repayment_indices])
ix3 = index_date_col(date_first_drawdown, df_reindexed.iloc[fee_indices])


# aggregate values per date
def agg_date(ix, vals):
    """Return a pandas Series: index = unique indices, values = summed values."""
    return pd.Series(vals, index=pd.Index(ix)).groupby(level=0).sum()
    
fees = agg_date(ix3, fees)

# 
df_return.loc[ix1, 'Advances'] = np.array(drawdowns)*-1
df_return.loc[ix2, 'Repayments'] = np.array(repayments)*-1
df_return.loc[ix3, 'Fees'] = fees

#print(drawdown_col(date_first_drawdown, df_reindexed.iloc[drawdown_indices]))


# In[11]:


len(date_range)


# In[12]:


# add variable interest rate

# add 0 for starting index
ix_int = []

for d in interest_dates:
    ix_int.append((d-date_first_drawdown).days)

ix_int.append(len(date_range)-1)


# In[13]:


ix_int


# In[14]:


prev_index = 0
counter = 0
for i in ix_int:
    df_return.loc[prev_index:i,'Interest Rate'] = interest_rates[counter]
    prev_index = i
    counter+=1


# In[15]:


df_return


# In[ ]:





# In[19]:


# add opening and closing balance
i = df_return.index[0]

df_return.loc[i, 'Closing Balance'] = (
    df_return.loc[i, ['Advances', 'Fees', 'Interest Charged']].sum()
    - df_return.loc[i, 'Repayments']
)

df_return.loc[i, 'Interest'] =     df_return.loc[i, 'Interest'] = (
    df_return.loc[i, ['Advances', 'Opening Balance', 'Fees']].sum()
    - df_return.loc[i, 'Repayments']
    )*df_return.loc[i,'Interest Rate']/365

start_month_index = 0

for i in df_return.index:
    if i ==0:
        continue

    df_return.loc[i, 'Opening Balance'] = df_return.loc[i-1,'Closing Balance']

    df_return.loc[i, 'Interest'] = (
    df_return.loc[i, ['Advances', 'Opening Balance', 'Fees']].sum()
    - df_return.loc[i, 'Repayments']
    )*df_return.loc[i,'Interest Rate']/365

    date = df_return.loc[i, 'Date']
    
    if date.is_month_end:
        interest_total = df_return['Interest'].iloc[start_month_index:i+1].sum()
        df_return.loc[i, 'Interest Charged'] = interest_total
        start_month_index = i+1


    df_return.loc[i, 'Closing Balance'] = (
    df_return.loc[i, ['Advances', 'Opening Balance', 'Fees', 'Interest Charged']].sum()
    - df_return.loc[i, 'Repayments']
    )

    


# In[20]:


df_return


# In[ ]:

print("Completed amort")

print("Generating Excel")

df_return.to_excel('_output_'+ file_path.lower(), sheet_name='amort') 

print("Complete, see folder.")


# In[ ]:




