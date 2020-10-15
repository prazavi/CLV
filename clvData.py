import xlsxwriter
import pandas as pd



# This block will create a list if time id's based on the i in CLV
df_time=pd.read_csv('time_data.csv')
timeList=[]
year=1997
month=1
timeList.append([])
for i in range(len(df_time)):
    if int(df_time['the_year'][i])==year and int(df_time['month_of_year'][i])==month:
        timeList[(month-1)+12*(year-1997)].append(int(df_time['time_id'][i]))
    elif int(df_time['the_year'][i])==year:
        month+=1
        timeList.append([])
        timeList[(month-1)+12*(year-1997)].append(int(df_time['time_id'][i]))
    else:
        year+=1
        month=1
        timeList.append([])
        timeList[(month-1)+12*(year-1997)].append(int(df_time['time_id'][i]))

# This block is for calculating CLV and customer's data
df_sales=pd.read_csv('sales_data.csv')
df_customer=pd.read_csv('custome_data.csv')


customerList=[]
customer_id=9
customer_id_list=[]
clv=0
tavan=1
coval=0
recency=[]

for i in range(len(df_customer)):
    customer_id_list.append(int(df_customer['customer_id'][i]))


for i in range(len(df_sales)):
    if customer_id==int(df_sales['customer_id'][i]):
        for j in range(len(timeList)):
            if int(df_sales['time_id'][i]) in timeList[j]:
                coval=24-j
        tavan=(-1)*(coval-0.5)
        clv+=(float(df_sales['store_sales'][i])-float(df_sales['store_cost'][i]))/(1.01**tavan)
    else:
        recency.append([customer_id,1096-int(df_sales['time_id'][i-1])])
        customerList.append([])
        customer=customer_id_list.index(int(customer_id))
        customerList[-1]=[customer_id,clv,df_customer['marital_status'][customer],df_customer['yearly_income'][customer],df_customer['gender'][customer],df_customer['num_children_at_home'][customer],df_customer['member_card'][customer]]
        customer_id=int(df_sales['customer_id'][i])
        clv=0
customerList.append([customer_id,clv,df_customer['marital_status'][customer],df_customer['yearly_income'][customer],df_customer['gender'][customer],df_customer['num_children_at_home'][customer],df_customer['member_card'][customer]])




# exporting in an excel

# workbook = xlsxwriter.Workbook('clv_data.xlsx')
# worksheet = workbook.add_worksheet()
# row = 0
# column = 0
# worksheet.write(row, 0, 'customer_id')
# worksheet.write(row, 1, 'clv')
# worksheet.write(row, 2, 'marital_status')
# worksheet.write(row, 3, 'yearly_income')
# worksheet.write(row, 4, 'gender')
# worksheet.write(row, 5, 'num_children_at_home')
# worksheet.write(row, 6, 'member_card')
# for i in customerList:
#     row+=1
#     worksheet.write(row, 0, i[0])
#     worksheet.write(row, 1, i[1])
#     worksheet.write(row, 2, i[2])
#     worksheet.write(row, 3, i[3])
#     worksheet.write(row, 4, i[4])
#     worksheet.write(row, 5, i[5])
#     worksheet.write(row, 6, i[6])
# workbook.close()


workbook = xlsxwriter.Workbook('recency.xlsx')
worksheet = workbook.add_worksheet()
row = 0
column = 0
worksheet.write(row, 0, 'customer_id')
worksheet.write(row, 1, 'Recency')
for i in recency:
    row+=1
    worksheet.write(row, 0, i[0])
    worksheet.write(row, 1, i[1])
workbook.close()