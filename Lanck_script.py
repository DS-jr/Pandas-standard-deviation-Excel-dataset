import pandas as pd

k = 20

df = pd.read_excel(
    io='Lanck_input_data.xlsx',
#    usecols=['Attempts'], 
    nrows=k)

df['P'] = df['ASR (%)'] / 100
df['Successful Calls (X)'] = df['Attempts'] * df['ASR (%)'] / 100

P_mean = df['P'].mean()

df['Sigma Pi'] = (P_mean * (1 - P_mean) / df['Attempts']) ** 0.5 
df['Z'] = (df['P'] - P_mean) / df['Sigma Pi']

df['R'] = pd.Series(dtype='float64') 
for j in range(1, k):
    df['R'][j] = abs(df['Sigma Pi'][j] - df['Sigma Pi'][j-1])    

df['R_primed'] = pd.Series(dtype='float64')
for j in range(1, k):
    df['R_primed'][j] = abs(df['Z'][j] - df['Z'][j-1])

R_primed_mean = df['R_primed'].sum() / (k - 1)
Sigma_Z = R_primed_mean / 1.128

df['UCL (Laney)'] = P_mean + 3 * df['Sigma Pi'] * Sigma_Z
df['LCL (Laney)'] = P_mean - 3 * df['Sigma Pi'] * Sigma_Z

print(df)


