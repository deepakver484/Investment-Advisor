from http import client
from statistics import variance
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import matplotlib.pyplot as plt
import pandas as pd
import csv
import numpy as np

# setting up the connection with the google spread sheet
scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/drive'
]

# setting up the credentials
credentials1 = ServiceAccountCredentials.from_json_keyfile_name('service_account.json', scopes=scopes)
gc = gspread.authorize(credentials1)

# Opening all the file of spread sheets
bse_1 = gc.open("BSE500").sheet1
user_bse = gc.open("Income / Expense").sheet1
final_report = gc.open("Final Report")
Report = final_report.get_worksheet(0)
Report1 = final_report.get_worksheet(1)

# get all the data in list of dictionary and assigning to variable for further aggregation
data_bse = bse_1.get_all_records()
user_info = user_bse.get_all_records()


# this function working as a sumif function
def calculation(filter, filter_coulumn, sum_column, dataset):
    net_amount = 0
    for row in dataset:
        if row[filter_coulumn] == filter:
            net_amount += row[sum_column]
    return net_amount


#######################-----------------------------------------######################################
# ---------------------------------------SUB-TASK-1----------------------------------------------------
#######################-----------------------------------------######################################


# calling the function and try to find the output ask in the subtask_1
net_income = calculation("Income", "Income/Expense", "INR", user_info)
net_expense = calculation("Expense", "Income/Expense", "INR", user_info)
available_for_Investment = net_income - net_expense

# find output for each category
Food = calculation("Food", "Category", "INR", user_info)
Others = calculation("Other", "Category", "INR", user_info)
Transportation = calculation("Transportation", "Category", "INR", user_info)
Social_life = calculation("Social Life", "Category", "INR", user_info)
Household = calculation("Household", "Category", "INR", user_info)
Apparel = calculation("Apparel", "Category", "INR", user_info)
Education = calculation("Education", "Category", "INR", user_info)
Salary = calculation("Salary", "Category", "INR", user_info)
Allowance = calculation("Allowance", "Category", "INR", user_info)
Beauty = calculation("Beauty", "Category", "INR", user_info)
Gift = calculation("Gift", "Category", "INR", user_info)
Petty_cash = calculation("Petty cash", "Category", "INR", user_info)

# Populate the Final Report spreadsheet with the output of subtask 1
Report.update_acell("C7", net_income)
Report.update_acell("C8", net_expense)
Report.update_acell("C24", available_for_Investment)
Report.update_acell("C10", Food)
Report.update_acell("C11", Others)
Report.update_acell("C12", Transportation)
Report.update_acell("C13", Social_life)
Report.update_acell("C14", Household)
Report.update_acell("C15", Apparel)
Report.update_acell("C16", Education)
Report.update_acell("C17", Salary)
Report.update_acell("C18", Allowance)
Report.update_acell("C19", Beauty)
Report.update_acell("C20", Gift)
Report.update_acell("C21", Petty_cash)

#         //////////////////        INSIGHTS-SUB TASK - 1            //////////////////////

# Income and expenses done is calculated as net income and net expense
# After calculation we observed that net income was less than net expense
# Therefore we manipulated the data just by adding some rows for income to make the difference positive.
# We did that to find difference between net income and net expense as Available for investment.
# We find out the Cost incurred in different categories : Food, Other, Transportation, Social Life,
# Household, Apparel, Education, Salary, Allowance, Self-development, Beauty, Gift, Petty cash.

# ---------------------------------------------------------------------------------------------------------------------

#######################-----------------------------------------######################################
# ---------------------------------------SUB-TASK-2 ----------------------------------------------------
#######################-----------------------------------------######################################


# adding a new column name Delta and populate that coulmn with the result


# adding the header of the column
bse_1.update_acell("AP1", "Delta")

# calculate values for Delta column and hold these values in a list name called delta
delta = []
for row in data_bse:
    if row["Company"] != "":
        Weekhigh_52 = row["52 Week High"]
        Price = row["Price"]
        formula = [(Weekhigh_52 - Price) / Weekhigh_52]
        delta.append(formula)

# Populate the Delta column with their respective value   
bse_1.update(f"AP2:AP{len(delta) + 1}", delta)

# Making a data frame with the help of pandas and put all the data of BSE500 to that dataframe
df = pd.DataFrame(data_bse)

# try to replace all the blank column with the nan with the help of pandas
df2 = df.mask(df == '')

# Calculate the value according to filter given in the subtask-2 for highRiskTaking
var1 = df2[(df2['Delta'] > 0) & (df2['Market Cap(Cr)'] < 4000) & (df2['10-Year Return(%)'] < 8)] \
    .sort_values(by=['Dividend Per Share'], ascending=False)

# Calculate the value according to filter given in the subtask-2 for RiskTaking
var2 = df2[(df2['Delta'] > 0) & (df2['Market Cap(Cr)'] > 4000) & (df2['Market Cap(Cr)'] < 10000) \
           & (df2['10-Year Return(%)'] > 8) & (df2['10-Year Return(%)'] < 15)] \
    .sort_values(by=['Dividend Per Share'], ascending=False)

# Calculate the value according to filter given in the subtask-2 for ModerateRiskTaking
var3 = df2[(df2['Delta'] > 0) & (df2['Market Cap(Cr)'] > 10000) & (df2['Market Cap(Cr)'] < 18000) \
           & (df2['10-Year Return(%)'] > 15) & (df2['10-Year Return(%)'] < 20)] \
    .sort_values(by=['Dividend Per Share'], ascending=False)

# Calculate the value according to filter given in the subtask-2 for LowRiskTaking
var4 = df2[(df2['Delta'] > 0) & (df2['Market Cap(Cr)'] > 18000) & (df2['10-Year Return(%)'] > 20)] \
    .sort_values(by=['Dividend Per Share'], ascending=False)

# Calculate top 5 Company for all the fields
High_Risk_Taking = list(var1.head()['Company'])
Risk_Taking = list(var2.head()['Company'])
Moderate_Risk_Taking = list(var3.head()['Company'])
Low_Risk_Taking = list(var4.head()['Company'])


# defining a function to populate sheet 2 with the top 5 companyName and investment value
def Risk_taking_func(value):
    if value == 'High Risk Taking':
        data = High_Risk_Taking
    elif value == 'Risk Taking':
        data = Risk_Taking
    elif value == 'Moderate Risk Taking':
        data = Moderate_Risk_Taking
    else:
        data = Low_Risk_Taking
    object = []
    j = 2
    for i in data:
        object.append([i])
        Report1.update_acell(f"C{j}", available_for_Investment / 5)
        j += 1
    Report1.update("B2:B6", object)


# get the value of C27 from final Report sheet 1
Investment_profile = Report.acell('C27').value

# calling the Risk_taking_func and populate Final Report sheet with the data
Risk_taking_func(Investment_profile)


#         //////////////////        INSIGHTS-SUB TASK - 2            //////////////////////


# we collected the data for “Available for Investment” from SUB TASK - 1 and divided that amount by 5.
# This amount was to be invested.
# we observed that we need to find top 5 companies for each tag(high Risk Taking, Risk Taking, Moderate Risk Taking,
# low risk taking) by applying given filters.
# we applied filters and came to know  that there is no company whose  Market Cap(Cr) is lesser than 2000
# Therefore we applied filters as per our own. i.e
# high Risk Taking (< 4000 cr)
# Risk Taking between 4000 cr and 10000 cr
# Moderate Risk Taking between 10000 cr and 18000 cr
# low risk taking > 18000 cr

# ----------------------------------------------------------------------------------------------------------------------

#######################-----------------------------------------######################################
# ---------------------------------------SUB-TASK-3 ----------------------------------------------------
#######################-----------------------------------------######################################


"""
Read data from spreadsheet and convert it into dictionary
"""


def get_data_in_Df_Format(spreadsheet_path):
    sheet_data = gc.open_by_url(spreadsheet_path)
    data_in_dict = sheet_data.sheet1.get_all_records()
    return pd.DataFrame.from_dict(data_in_dict)


# ---------------------------------------- PART-1 ---------------------------------------------------

"""
Comparing the median of column Enterprise Value(Cr) across different Sectors.
"""


def get_median_of_each_sector():
    # Reading 'Sector' & 'Enterprise Value(Cr)' columns from BSE500 data table as we will use these column only.

    sector_Enterprise_df = BSE500_df[['Sector', 'Enterprise Value(Cr)']]

    # Replacing blank cells with 'NaN' to remove  the errors.

    sector_Enterprise_df['Enterprise Value(Cr)'] = sector_Enterprise_df['Enterprise Value(Cr)']. \
        replace('', np.nan).astype(float)

    # Grouping 'Sector' columns and finding median of 'Enterprise Value(Cr)' values

    res = sector_Enterprise_df.groupby(['Sector'])['Enterprise Value(Cr)'].median().reset_index()
    res.rename(columns={'Enterprise Value(Cr)': 'Enterprise Value(Cr) Median'}, inplace=True)

    # Storing the output into csv file

    res.to_csv("result/Task-3-1.csv")

    return res


#         //////////////////        INSIGHTS-SUB TASK -3 (part-1)            //////////////////////

# When we find median of 'Enterprise Value(Cr)', sector wise, we came to know that some companies in sectors are above
# the median value and some are below the median value.
# median  basically separates the higher half from a lower half
# Enterprise Value = market capitalization + total debt - C
# market capitalization = current stock * n/o outstanding stock share
# total debt = short term debt + long term debt
# C = cash and cash equivalents
#  we analyse that companies above median value are basically large cap companies also there are mid cap companies.



# ---------------------------------------- PART-2 ---------------------------------------------------


"""
finding a relation between Dividend Per Share with Market Cap(Cr)
"""


def get_correlation_between_given_columns(col_name1, col_name2):
    # Reading 'Market Cap(Cr)' & 'Dividend Per Share' columns from BSE500 data table.

    market_cap_divident_df = BSE500_df[['Market Cap(Cr)', 'Dividend Per Share']]

    # Replacing blank cells with 'NaN' to remove  the errors.

    market_cap_divident_df['Market Cap(Cr)'] = market_cap_divident_df['Market Cap(Cr)']. \
        replace('', np.nan).astype(float)
    market_cap_divident_df['Dividend Per Share'] = market_cap_divident_df['Dividend Per Share']. \
        replace('', np.nan).astype(float)

    # Finding Correlation between 'Market Cap(Cr)', 'Dividend Per Share' columns

    final_corr = market_cap_divident_df[col_name2].corr(market_cap_divident_df[col_name1])
    corr_df = pd.DataFrame([final_corr], columns=["Correlation"])

    # Storing the output into csv file

    corr_df.to_csv("result/Task-3-2.csv")

    return final_corr

#         //////////////////        INSIGHTS-SUB TASK - 3(part-2)           //////////////////////

# Market Cap(Cr) = stock price * shares outstanding
# share outstanding = Issued share - treasury share
# Dividend per share = dividend yield(%) * share price
# We try to find correlation between 'Market Cap(Cr)' & 'Dividend Per Share and found a correlation of 0.04
# which means we can say that they have nearly Zero correlation because they don't seem to be linked at all.



# ---------------------------------------- PART-3 ---------------------------------------------------
"""
Count the companies in different Industry with positive and negative 3-Year Return and
deciding which industry would be recommended to someone to invest if the same return is followed.
"""


def get_3year_return_analysis():
    # Reading 'Industry' & '3-Year Return' columns from BSE500 data table.

    industry_3year_return_df = BSE500_df[['Industry', '3-Year Return']]

    # Replacing blank cells with 'NaN' to remove  the errors.

    industry_3year_return_df['3-Year Return'] = industry_3year_return_df['3-Year Return'] \
        .replace('', np.nan).astype(float)

    # Finding negative & positive 3 years return from the column '3-Year Return'.

    industry_with_negative_3year = industry_3year_return_df[industry_3year_return_df['3-Year Return'] < 0]
    industry_with_positive_3year = industry_3year_return_df[industry_3year_return_df['3-Year Return'] >= 0]

    # Grouping 'Industry' columns and counting companies having negative '3-Year Return' value
    # & positive '3-Year Return' value

    negative_result = industry_with_negative_3year.groupby(['Industry'])['3-Year Return'].count().reset_index()
    positive_result = industry_with_positive_3year.groupby(['Industry'])['3-Year Return'].count().reset_index()

    # Converting dataframe into dictionary by taking industry column as key

    negative_result_dict = negative_result.set_index('Industry').T.to_dict('list')
    positive_result_dict = positive_result.set_index('Industry').T.to_dict('list')

    # fetching distinct value of industry

    distinct_industry_value = industry_3year_return_df.Industry.unique()

    # Reading above created dictionaries and storing result in the form of list of list
    # (industry,negative_count,positive_count)
    res = []
    for i in distinct_industry_value:
        temp = [i]
        if i in negative_result_dict:
            temp.append(negative_result_dict[i][0])
        else:
            temp.append(0)
        if i in positive_result_dict:
            temp.append(positive_result_dict[i][0])
        else:
            temp.append(0)
        res.append(temp)

    # we are sorting result on the basis of positive 3 year return

    res.sort(key=lambda x: x[2], reverse=True)
    df = pd.DataFrame(res, columns=['Industry', 'negative_count', 'positive_count'])

    # Storing the output into csv file

    df.to_csv("result/Task-3-3.csv")
    return df


#         //////////////////        INSIGHTS-SUB TASK - 3(part-3)           //////////////////////

# some companies have positive 3 year return
# some companies have negative 3 year return
# In this part after applying all the conditions we can see that Pharma industry have most number of company listed.
# and they also hve the most number of companies with positive 3 years return.
#  Since same return to be followed, Pharma industry will be recommended for investment.



# ---------------------------------------- PART-4 ---------------------------------------------------

"""
the best stock across different Sector considering one of the KPI
"""


def get_best_stock_per_industry():
    # Reading 'Sector', 'Company' & 'Price to Earnings' columns from BSE500 data table and
    # sorting the values of the column 'Sector', 'Price to Earnings'

    company_sector_PE_ratio_df = BSE500_df[['Sector', 'Company', 'Price to Earnings']].sort_values(
        by=['Sector', 'Price to Earnings'])

    # Replacing blank cells with 'NaN' to remove  the errors.

    company_sector_PE_ratio_df['Price to Earnings'] = company_sector_PE_ratio_df['Price to Earnings'] \
        .replace('', np.nan).astype(float)

    # grouping by sector and picking best stock from each sector.

    final_res = company_sector_PE_ratio_df.groupby(['Sector']).head(1)

    # Storing the output into csv file.

    final_res.to_csv("result/Task-3-4.csv")

    return final_res


#         //////////////////        INSIGHTS-SUB TASK - 3(part-4)           //////////////////////

# There were 18 distinct sectors
# In each sector there are many companies
# KPI is Key Performance Indicator
# There are many KPI's than can decide the best stock in each sector like EV / EBITDA,Price to Earnings,
# Dividend Per Share,Price earnings to growth,Price to Book etc.
#  We choose our KPI as Price to Earnings
# Price to Earnings = share price/earning per share
# We thought that if earning per share is more means companies are doing a good business. They have works,projects
# ,market in their control.
# So higher is earning per share, lesser will be Price to Earnings and hence
# company with least Price to Earnings will be the best stock of that sector.

# ---------------------------------------------------------------------------------------------------------------------


# For plotting bar graph

def make_bar_graph(dataFrame, x_axis, y_axis, title):
    dataFrame.plot(kind='bar', x=x_axis, y=y_axis, fontsize='7', color='Purple')
    plt.title(title)
    plt.show()


# For plotting scattered bar graph

def make_scattered_graph(x_axis_data, y_axis_data, x_axis, y_axis, title):
    plt.title(title)
    plt.scatter(x_axis_data.replace('', np.nan).astype(float), y_axis_data.replace('', np.nan).astype(float))

    # Below written codes are used for getting correlation.(positive,negative or zero correlation)

    plt.plot(np.unique(x_axis_data.replace('', np.nan).astype(float)), np.poly1d(
        np.polyfit(x_axis_data.replace('', np.nan).astype(float), y_axis_data.replace('', np.nan).astype(float), 1))
    (np.unique(x_axis_data.replace('', np.nan).astype(float))), color='red')
    plt.xlabel(x_axis)
    plt.ylabel(y_axis)
    plt.show()


# For clustered bar graph

def make_cluster_bar_graph(data, x_axis, title):
    data.plot(x=x_axis,
              kind='bar',
              stacked=False,
              title=title)
    plt.show()


if __name__ == '__main__':
    #  To get BSE500 data

    BSE500_df = get_data_in_Df_Format('https://docs.google.com/spreadsheets/d'
                                      '/1xRMYSr048KLJReIrLDQ8GPlYndUZppGmusixu_NZARE/edit#gid=1547461811')
    # Calling PART-1

    data = get_median_of_each_sector()
    make_bar_graph(data, 'Sector', 'Enterprise Value(Cr) Median',
                   'bar plot of Sector VS Enterprise Value(Cr) median Value')
    # Calling PART-2

    correlation = get_correlation_between_given_columns('Market Cap(Cr)', 'Dividend Per Share')
    make_scattered_graph(BSE500_df['Market Cap(Cr)'],
                         BSE500_df['Dividend Per Share'],
                         'Market Cap(Cr)', 'Dividend Per Share', 'Correlation')
    # Calling PART-3

    data = get_3year_return_analysis()
    make_cluster_bar_graph(data.head(20), 'Industry', 'Industry 3year return positive negative count')

    # Calling PART-4

    res = get_best_stock_per_industry()
    res['Sector_Company'] = res.apply(lambda x: '%s_%s' % (x['Sector'], x['Company']), axis=1)
    make_bar_graph(res[['Sector_Company', 'Price to Earnings']], 'Sector_Company',
                   'Price to Earnings', 'Best Stock across different Sector')
