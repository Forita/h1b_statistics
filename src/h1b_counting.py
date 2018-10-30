import pandas as pd

df1 = pd.read_excel('./input/H-1B_FY14_Q4.xlsx')
df2 = pd.read_excel('./input/H-1B_Disclosure_Data_FY15_Q4.xlsx')
df3 = pd.read_excel('./input/H-1B_Disclosure_Data_FY16.xlsx')

df_2014 = df1[['LCA_CASE_NUMBER', 'STATUS', 'LCA_CASE_EMPLOYER_STATE', 'LCA_CASE_SOC_CODE', 'LCA_CASE_SOC_NAME']]
df_2015 = df2[['CASE_NUMBER', 'CASE_STATUS', 'EMPLOYER_STATE', 'SOC_CODE', 'SOC_NAME']]
df_2016 = df3[['CASE_NUMBER', 'CASE_STATUS', 'EMPLOYER_STATE', 'SOC_CODE', 'SOC_NAME']]

df_2014.columns = ['CASE_NUMBER', 'CASE_STATUS', 'EMPLOYER_STATE', 'SOC_CODE', 'SOC_NAME']
df = pd.concat([df_2014, df_2015, df_2016])

# total number of certified visa applications
totalCert = sum(df.CASE_STATUS == 'CERTIFIED')
print(totalCert)

# top 10 occupations
occup = df.groupby('SOC_NAME')
df_cert = occup.apply(lambda x: sum(x.CASE_STATUS == 'CERTIFIED'))
df_cert = df_cert.sort_values(ascending=False)

top_occup = pd.DataFrame()
top_occup['NUMBER_CERTIFIED_APPLICATIONS'] = df_cert[:10]
top_occup ['PERCENTAGE'] = df_cert[:10] / totalCert

top_occup.reset_index(level=0, inplace=True)
top_occup['PERCENTAGE'] = pd.Series(["{0:.2f}%".format(val * 100) for val in top_occup['PERCENTAGE']], index = top_occup.index)
top_occup.columns = ['TOP_OCCUPATIONS', 'NUMBER_CERTIFIED_APPLICATIONS', 'PERCENTAGE']

# top 10 states
state = df.groupby('EMPLOYER_STATE')
df_state = state.apply(lambda x: sum(x.CASE_STATUS == 'CERTIFIED'))
df_state = df_state.sort_values(ascending=False)

top_state = pd.DataFrame()
top_state['NUMBER_CERTIFIED_APPLICATIONS'] = df_state[:10]
top_state ['PERCENTAGE'] = df_state[:10] / totalCert

top_state.reset_index(level=0, inplace=True)
top_state['PERCENTAGE'] = pd.Series(["{0:.2f}%".format(val * 100) for val in top_state['PERCENTAGE']], index = top_state.index)
top_state.columns = ['TOP_STATES', 'NUMBER_CERTIFIED_APPLICATIONS', 'PERCENTAGE']

# output
top_occup.to_csv('./output/top_10_occupations.txt', index=None, sep=';', mode='a')
top_state.to_csv('./output/top_10_states.txt', index=None, sep=';', mode='a')
