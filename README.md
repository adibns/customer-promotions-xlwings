# customer-promotions-xlwings
An excel report is generated at 5 A.M. every 90 days which has customers that should be attracted by giving promotions along with customer total order value plots and geo plots.

- The script connects with a Postgres DB, fetches summarised data from DB by running an SQL query to python.
- An excel report is generated with this fetched data.
- A plotly express geo plot shows heatmap of province wise purchase values, a histogram of total purchase values per customer using seaborn is generated in new respective sheets.
- An excel plot is generated in a new sheet using the data in the excel sheet.
