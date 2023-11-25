#Poject Prepared by CPA Abdulrahman Juma Bakari

#Import the required modules
import  requests as rq
import json
import matplotlib.pyplot as plt
import xlwings as xl
import pandas as pd
import seaborn as sns

#get the avilable keys for the data
api_key = open('bst.txt','r').read()

#Get the company for the analysis
comapany = "FB"

#Find the URL
url = f'https://financialmodelingprep.com/api/v3/income-statement/{comapany}'
payload = {
   'limit':5,
   'apikey':api_key
}

#Get the data
request = rq.get(url,params=payload).json()

print(json.dumps(request,indent=3))


expenses = ["researchAndDevelopmentExpenses",
      "generalAndAdministrativeExpenses",
      "sellingAndMarketingExpenses",
      "sellingGeneralAndAdministrativeExpenses",
      "otherExpenses",
      "operatingExpenses",
      "costAndExpenses",
      "interestExpense",
      ]

wb = xl.Book("Book1.xlsx")
sheet = wb.sheets['Sheet1']
year = 3

#Put in the headings
sheet.range('H5').value = 'Facebook Profit and Loss account'
sheet.range('H6').value = f'for the year ended {request[year]["date"]}'
sheet.range('H7').value = 'Accounts prapared by CPA Abdulrahman Juma Bakari'

#General update before loading the data
sheet.range('F9:L100').clear()

#Set the trading acocunts in the excel file
sheet.range('G9').value = 'Revenue'
sheet.range('K9').value = request[year]['revenue']

trading_acc = ['Cost of goods sold','Gross Profit','Other Income']
trading_acc_amount = ['costOfRevenue','grossProfit','interestIncome']

counter = 11
for acc in range(3):
    sheet.range(f'G{counter}').value = trading_acc[acc]
    if trading_acc_amount[acc] == 'grossProfit' or trading_acc_amount[acc] == 'interestIncome':
        sheet.range(f'K{counter}').value = request[year][trading_acc_amount[acc]]
    else:
        sheet.range(f'I{counter}').value = request[year][trading_acc_amount[acc]]
        counter += 1

#Setting the expenses accounts
sheet.range('G15').value = 'Expenses'
index = len(expenses)
total_expenses = 0
last_exp_cell = 16 + len(expenses)
for cell in range(index):

    sheet.range(f'G{15 + cell + 1  }').value = expenses[cell]
    sheet.range(f'I{15 + cell + 1}').value = request[year][expenses[cell]]
    total_expenses += request[year][expenses[cell]]

sheet.range(f'G{last_exp_cell}').value = "Total Expenses"
sheet.range(f'K{last_exp_cell}').value = total_expenses

#Get the Undisclosed Income
income_not_disclosed = (request[year]['ebitda'] + total_expenses) - (request[year]['grossProfit'] + request[year]["interestIncome"])

sheet.range(f'G{last_exp_cell + 1}').value = "Income Not Disclosed"
sheet.range(f'L{last_exp_cell + 1}').value = income_not_disclosed

last_exp_cell =  last_exp_cell + 2

#Set the last account fo the income statement
others = ["ebitda",
          "depreciationAndAmortization",
          "incomeBeforeTax",
          "incomeTaxExpense",
          "netIncome"]


iterator = len(others)
for rest in range(iterator):


    if others[rest] == "netIncome" :
        sheet.range(f'G{rest + 1 + last_exp_cell}').value = 'Net Income'
    else:
        sheet.range(f'G{rest + 1 + last_exp_cell}').value = others[rest]

    if others[rest] == "depreciationAndAmortization" or others[rest] == "incomeTaxExpense":
        sheet.range(f'I{rest + 1 + last_exp_cell}').value =  request[year][others[rest]]
    else:
        sheet.range(f'L{rest + 1 + last_exp_cell}').value = request[year][others[rest]]

revenues = list(reversed([request[i]['revenue'] for i in range(len(request))]))
profit = list(reversed([request[i]['netIncome'] for i in range(len(request))]))
research_D = list(reversed([request[i]["researchAndDevelopmentExpenses"] for i in range(len(request))]))
S_M = list(reversed([request[i]["sellingAndMarketingExpenses"] for i in range(len(request))]))
SG_A = list(reversed([request[i]["sellingGeneralAndAdministrativeExpenses"] for i in range(len(request))]))


df = pd.DataFrame({"reveue":revenues, "profit":profit,"Research_d":research_D,"Selling_m":S_M,'GenSelling_M':SG_A})

sns.heatmap(df.corr(),annot=True,cmap='Blues')

plt.plot(revenues,label='Revenue')
plt.plot(profit,label='Profit')
plt.legend(loc='upper left')
plt.show()

print(total_expenses)
print(last_exp_cell)

