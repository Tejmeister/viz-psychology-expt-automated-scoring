#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
import argparse
import json


parser = argparse.ArgumentParser(
		description='Inputs for automated scoring')
parser.add_argument('--input_responses_file_path', type=str,
										help='File path for input responses', required=True)
parser.add_argument('--output_responses_file_path', type=str,
										help='File path for output responses', required=True)


def fn_read_json(input_file):
	with open(input_file, encoding="utf-8") as f:
		data = json.load(f)
		f.close()
	return data


inputResponsesExcelFile = "C:/Users/tejparma/Personal/Projects/viz-psychology-expt-automated-scoring/Input_Responses.xlsx"
finalResponsesExcelFile = "C:/Users/tejparma/Personal/Projects/viz-psychology-expt-automated-scoring/Output_Test.xlsx"


# In[3]:

# Self Esteem response mapping
se_response_mapping = fn_read_json("./metadata/Self-Esteem-Response-Mapping.json")


# upMapping = {"strongly agree": 1, "agree": 2, "disagree": 3, "strongly disagree": 4}
# downMapping = {"strongly agree": 4, "agree": 3, "disagree": 2, "strongly disagree": 1}

# Self Esteem questions to response mapping
# upQuestions = ["Q1","Q3","Q4","Q7","Q10"]
# downQuestions = ["Q2","Q5","Q6","Q8","Q9"]

se_q_r_mapping = fn_read_json("./metadata/Self-Esteem-Question-to-Response-Mapping.json")

questions = fn_read_json("./metadata/Questions-List.json")

# rAMapping = ['Q1', 'Q2', 'Q3', 'Q6', 'Q8', 'Q11', 'Q12', 'Q13', 'Q14', 'Q16', 'Q21', 'Q24', 'Q25', 'Q27', 'Q29', 'Q30', 'Q31', 'Q33', 'Q34', 'Q36', 'Q37', 'Q38', 'Q39']
# rBMapping = ['Q4', 'Q5', 'Q7', 'Q9', 'Q10', 'Q15', 'Q17', 'Q18', 'Q19', 'Q20', 'Q22', 'Q23', 'Q26', 'Q28', 'Q32', 'Q35', 'Q40']

questions_mapping 

leadershipQuestions = ["Q1", "Q5", "Q10", "Q11", "Q12", "Q27", "Q32", "Q33", "Q34", "Q36", "Q40"]
grandioseQuestions = ["Q4", "Q7", "Q15", "Q19", "Q20", "Q26", "Q28", "Q29", "Q30", "Q38"]
entitlementQuestions = ["Q13", "Q14", "Q24", "Q25"]

# In[4]:

pd.set_option('max_columns', None)
df = pd.read_excel(inputResponsesExcelFile, engine='openpyxl') 

demographicHeaders = df.columns[2:9].tolist()

demographics = df.iloc[:, 2: 9].values.tolist()
dList = [[f'P{i+1}'] for i in range(0,len(demographics))]
demographicHeaders.append("")
demographics.append("")

selfEsteem = df.iloc[:, 9: 19].values.tolist()

selfEsteemHeader = ["SelfEsteem"]
for i in range(len(selfEsteem[0])):
		selfEsteemHeader.append(f'Q{i+1}')

res1 = [i + [""] + j for i, j in zip(demographics, dList)]

res = [i + j for i, j in zip(res1, selfEsteem)]

remainingQuestions = df.iloc[:, 19: ].values.tolist()
rowHeaders = demographicHeaders + selfEsteemHeader

df1 = pd.DataFrame(res, columns = rowHeaders)

newDf = df1.transpose()
newDf.columns = [f'P{i}' for i in range(1,len(demographics))]

for colName in newDf.columns:  
		for _, row in newDf.iterrows():    
				if _ in upQuestions:
						newDf.loc[_,colName] = upMapping[row.loc[colName].lower()]
				elif _ in downQuestions:
						newDf.loc[_,colName] = downMapping[row.loc[colName].lower()]
						
writer = pd.ExcelWriter(finalResponsesExcelFile, engine = 'openpyxl')

print("Writing self esteem scoring..")
newDf.to_excel(writer, sheet_name = "self esteem")

# In[5]:

remainingQuestions = df.iloc[:, 19: ].values.tolist()

for question in remainingQuestions:
		for i, item in enumerate(question):
				#print(f'Q{i+1}', item)
				item = item.replace('.', '')
				question[i] = questions[f'Q{i+1}'][item]


remainingQuestionsHeader = ["NARCISSISM"]
for i in range(len(remainingQuestions[0])):
		remainingQuestionsHeader.append(f'Q{i+1}')

rowHeaders = demographicHeaders + remainingQuestionsHeader
res1 = [i + [""] + j for i, j in zip(demographics, dList)]

res = [i + j for i, j in zip(res1, remainingQuestions)] 
df2 = pd.DataFrame(res, columns = rowHeaders)

nDf = df2.transpose()
nDf.columns = [f'P{i}' for i in range(1,len(demographics))]

for colName in nDf.columns:  
		for _, row in nDf.iterrows():    
				if _ in rAMapping:
						if row.loc[colName] == 'A':
								nDf.loc[_,colName] = 1
						else:
								nDf.loc[_,colName] = 0
				elif _ in rBMapping:
						if row.loc[colName] == 'B':
								nDf.loc[_,colName] = 1
						else:
								nDf.loc[_,colName] = 0
								
print("Writing narcissism scoring..")
nDf.to_excel(writer, sheet_name = "narcissism")


# In[8]:

remainingQuestions = df.iloc[:, 19: ].values.tolist()

leaderShipFiltered = []

for rq in remainingQuestions:
		l = [rq[int(i.replace("Q", ""))-1] for i in leadershipQuestions]
		leaderShipFiltered.append(l)

leadershipHeader = []
leadershipHeader += leadershipQuestions

for question in leaderShipFiltered:
		for i, item in enumerate(question):
				item = item.replace('.', '')
				question[i] = questions[leadershipQuestions[i]][item]   
				
rowHeaders = demographicHeaders + leadershipHeader

df2 = pd.DataFrame(leaderShipFiltered, columns = leadershipHeader)

lDf = df2.transpose()
lDf.columns = [f'P{i}' for i in range(1,len(demographics))]

for colName in lDf.columns:  
		for _, row in lDf.iterrows():    
				if _ in rAMapping:
						if row.loc[colName] == 'A':
								lDf.loc[_,colName] = 1
						else:
								lDf.loc[_,colName] = 0
				elif _ in rBMapping:
						if row.loc[colName] == 'B':
								lDf.loc[_,colName] = 1
						else:
								lDf.loc[_,colName] = 0
								
print("Writing leadership scoring..")
lDf.to_excel(writer, sheet_name = "leadership")


# In[10]:


remainingQuestions = df.iloc[:, 19: ].values.tolist()

entitlementFiltered = []

for rq in remainingQuestions:
		l = [rq[int(i.replace("Q", ""))-1] for i in entitlementQuestions]
		entitlementFiltered.append(l)

entitlementHeader = []
entitlementHeader += entitlementQuestions

for question in entitlementFiltered:
		for i, item in enumerate(question):
				item = item.replace('.', '')
				question[i] = questions[entitlementQuestions[i]][item]
				
df2 = pd.DataFrame(entitlementFiltered, columns = entitlementHeader)

lDf = df2.transpose()
lDf.columns = [f'P{i}' for i in range(1,len(demographics))]

for colName in lDf.columns:  
		for _, row in lDf.iterrows():    
				if _ in rAMapping:
						if row.loc[colName] == 'A':
								lDf.loc[_,colName] = 1
						else:
								lDf.loc[_,colName] = 0
				elif _ in rBMapping:
						if row.loc[colName] == 'B':
								lDf.loc[_,colName] = 1
						else:
								lDf.loc[_,colName] = 0

print("Writing entitlement scoring..")                
lDf.to_excel(writer, sheet_name = "entitlement")


# In[11]:


remainingQuestions = df.iloc[:, 19: ].values.tolist()

grandioseFiltered = []

for rq in remainingQuestions:
		l = [rq[int(i.replace("Q", ""))-1] for i in grandioseQuestions]
		grandioseFiltered.append(l)

grandioseHeader = []
grandioseHeader += grandioseQuestions

for question in grandioseFiltered:
		for i, item in enumerate(question):
				item = item.replace('.', '')
				question[i] = questions[grandioseQuestions[i]][item]
				
df2 = pd.DataFrame(grandioseFiltered, columns = grandioseHeader)

lDf = df2.transpose()
lDf.columns = [f'P{i}' for i in range(1,len(demographics))]

for colName in lDf.columns:  
		for _, row in lDf.iterrows():    
				if _ in rAMapping:
						if row.loc[colName] == 'A':
								lDf.loc[_,colName] = 1
						else:
								lDf.loc[_,colName] = 0
				elif _ in rBMapping:
						if row.loc[colName] == 'B':
								lDf.loc[_,colName] = 1
						else:
								lDf.loc[_,colName] = 0

print("Writing grandiose scoring..")
lDf.to_excel(writer, sheet_name = "grandiose")

writer.save()
writer.close()

