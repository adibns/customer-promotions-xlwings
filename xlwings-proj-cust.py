# Imports

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import psycopg2
import xlwings as xl
import plotly
import plotly.express as px
import json
from datetime import date
import openpyxl
import time
import schedule
import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning) 

# Set up a connection to the postgres DB server
con = psycopg2.connect("dbname=customer user=postgres password=pswd")

# Create a cursor object
cursor = con.cursor()

# SQL command to get region wise, country wise, province wise, city wise, department wise average salaies of each position
sql_command_ = """WITH customer_total_purchase AS (
	SELECT 
		cust.customer_id,
		SUM(orders.order_value) AS total_purchase
	FROM 
		customers cust JOIN orders 
		ON cust.customer_id = orders.customer_id
	GROUP BY
		cust.customer_id
), customer_geography AS (
	SELECT 
		customers.customer_id,
		regions.region_name,
		countries.country_name,
		locations.state_province,
		locations.zip_code
	FROM
		regions JOIN countries ON regions.region_id = countries.region_id
		JOIN locations ON countries.country_id = locations.country_id
		JOIN customers ON locations.zip_code = customers.zip_code
)
SELECT 
	customer_geography.region_name,
	customer_geography.country_name,
	customer_geography.state_province,
	customer_geography.zip_code,
	customer_total_purchase.customer_id,
	ROUND(customer_total_purchase.total_purchase,2) AS purchase_value
FROM 
	customer_total_purchase JOIN customer_geography 
		ON customer_total_purchase.customer_id = customer_geography.customer_id
ORDER BY 
	purchase_value DESC
LIMIT 200 OFFSET 100"""

# Job scheduling function
def job_scheduler(sql_command, conn):
	# Read SQL query output
	sql_result = pd.read_sql(sql_command, conn)

	# Save the summarised SQL output as excel sheet
	date_now = date.today()
	sql_result.to_excel(r"C:\Users\admin\Desktop\{}_high_val_cus_report.xlsx".format(str(date_now)), 
						index=False, 
						sheet_name='Reported Data')

	# Map plot
	fig = px.choropleth(sql_result, 
						locations='state_province',
						color='purchase_value',
						scope='north america',
						labels={'purchase_value': 'Province wise Total Purchase Value (in $)'})

	# Open saved workbook
	wb = xl.Book(r"C:\Users\admin\Desktop\{}_high_val_cus_report.xlsx".format(date_now))

	# Autofit all cells for readability
	wb.sheets['Reported Data'].autofit()

	# Plotly express plot
	wb.sheets.add(name='PX Plot')
	py_plot_sheet = wb.sheets['PX Plot']
	plot = py_plot_sheet.pictures.add(fig, 
									name='Total Spending', 
									update=True,
									left=py_plot_sheet.range('C2').left, 
									top=py_plot_sheet.range('C2').top)

	# Seaborn plot
	fig2 = plt.figure(figsize=(8,4))
	sns.distplot(sql_result['purchase_value'])
	wb.sheets.add(name='SNS Plot')
	sns_plot_sheet = wb.sheets['SNS Plot']
	plot = sns_plot_sheet.pictures.add(fig2, 
									name='Customer wise Total Purchase Value', 
									update=True, 
									left=sns_plot_sheet.range('C2').left, 
									top=sns_plot_sheet.range('C2').top)
	plot.height *= 0.9
	plot.width *= 0.9


	# Excel Plot
	wb.sheets.add(name='Excel Plot')
	excel_plot_sheet = wb.sheets['Excel Plot']
	excel_plot = excel_plot_sheet.charts.add(left=excel_plot_sheet.range('C2').left, top=excel_plot_sheet.range('C2').top)
	excel_plot.chart_type = 'line_markers_stacked'
	excel_plot.set_source_data(wb.sheets['Reported Data'].range('F2:G2').expand('down'))
	excel_plot.height *= 1.5
	excel_plot.width *= 1.5

	# Save the workbook
	wb.save()

	# Close all active excel applications
	xl.apps.active.quit()


# Scheduling 
schedule.every(90).day.at("05:00").do(job_scheduler,sql_command_, con)
while True: 
  schedule.run_pending()