import pandas as pd
import openpyxl


# import database
sells_table = pd.read_excel("Vendas.xlsx")

# view database
pd.set_option("display.max_columns", None)
#print(sells_table)

# invoicing by shopping mall

# filter and select the columns of database
invoicing_by_shopping_mall_table = sells_table[["ID Loja","Valor Final"]].groupby("ID Loja").sum()
# Add crescending values of invoicing by shopping mall
invoicing_by_shopping_mall_table = invoicing_by_shopping_mall_table[["Valor Final"]].sort_values(by="Valor Final", ascending=False)
#print(invoicing_by_shopping_mall_table)
print("-" * 50)

# how many products sells by store
amount_products_sells = sells_table[["ID Loja","Quantidade"]].groupby("ID Loja").sum()
amount_products_sells = amount_products_sells[["Quantidade"]].sort_values(by="Quantidade", ascending=False)
#print(amount_products_sells)
print("-" * 50)

# average value of product by store (amount / invoicing)
# create Ticket Medio column
average_product_value_by_shopping_mall = (invoicing_by_shopping_mall_table["Valor Final"] / amount_products_sells["Quantidade"]).to_frame()
#average_product_value_by_shopping_mall = sells_table[["ID Loja", "Ticket Médio"]].groupby("ID Loja").sum()
#average_product_value_by_shopping_mall = average_product_value_by_shopping_mall[["ID Loja"]].sort_values(by="ID Loja", ascending=False)
print(average_product_value_by_shopping_mall)

# send email as a report