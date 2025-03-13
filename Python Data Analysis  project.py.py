import pandas as pd 
import matplotlib.pyplot as plt

data = pd.read_excel('C:\Data.xlsx')

# Group by Region and calculate sum and average
region_summary = data.groupby('Region').agg({'Sales': ['sum', 'mean'], 'profit': ['sum', 'mean']})

# Plotting the data
# Bar Chart
region_summary['Sales']['sum'].plot(kind='bar', title='Total Sales by Region')
plt.show()

# Line Chart
region_summary['Sales']['mean'].plot(kind='line', title='Average Sales by Region')
plt.show()

# Pie Chart
region_summary['profit']['sum'].plot(kind='pie', title='Total Profit by Region', autopct='%1.1f%%')
plt.show()

# Histogram
data['Sales'].hist(by=data['Region'], bins=10)
plt.suptitle('Sales Distribution by Region')
plt.show()


#2-Sum, Average of Sales and Profit by Region, Province, and Product Category

# Group by Region, Province, and Product Category
detailed_summary = data.groupby(['Region', 'Province', 'Product Category']).agg({'Sales': ['sum', 'mean'], 'profit': ['sum', 'mean']})

# You can plot similarly as above for detailed_summary based on your needs



#3-Analyze Ship Mode with Product Category and Display Graphically
# Group by Ship Mode and Product Category
ship_mode_summary = data.groupby(['Ship Mode', 'Product Category']).size().unstack()

# Bar Chart
ship_mode_summary.plot(kind='bar', stacked=True, title='Product Category Distribution by Ship Mode')
plt.show()

# Line Chart (Not typical for categorical data but can be used if needed)
ship_mode_summary.plot(kind='line', title='Product Category by Ship Mode')
plt.show()

# Pie Chart (Not typical for this kind of comparison but can be used for a single category)
ship_mode_summary.sum().plot(kind='pie', title='Total Products by Category', autopct='%1.1f%%')
plt.show()

# Histogram
ship_mode_summary.plot(kind='hist', title='Ship Mode Distribution by Product Category', bins=10)
plt.show()