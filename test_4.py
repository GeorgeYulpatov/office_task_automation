import pandas as pd
import openpyxl as op
import numpy as np
import os
from pandas import DataFrame

# df_ozon.to_excel("ozon.xlsx", sheet_name="Шаблон", index=False)

df_ozon: DataFrame = pd.read_excel("ozon.xlsx", sheet_name="Шаблон")

# df_ozon = op.load_workbook("ozon.xlsx")

df_ozon["balance_total"] = df_ozon["ost_a"] + df_ozon["ost_b"]
df_ozon["estimated_sales"] = df_ozon["orders_current_week"] * 2
df_ozon["проверка_остатка"] = df_ozon["balance_total"] - df_ozon["estimated_sales"]
df_ozon["participation"] = ["no" if x <= 0 else "" for x in df_ozon["проверка_остатка"]]
df_ozon.drop(columns=["balance_total", "estimated_sales", "проверка_остатка"], axis=1, inplace=True)
df_ozon.to_excel(f"parts/part_1.xlsx",sheet_name="Шаблон",index=False)
print(df_ozon)
# Первый этап
df_ozon_ost = df_ozon.query('participation == "no"')
print(df_ozon_ost)
df_ozon_ost.to_excel(f"parts/part_2.xlsx",sheet_name="Нет",index=False)

# Второй этап
df_ozon_xxx = df_ozon.query('deviation_from_opt >= 0.1 & XXX == "xxx" & margin_in_stock >= 0.01 & participation != "no"')
df_ozon_xxx.participation = "yes2"
print(df_ozon_xxx)
df_ozon_xxx.to_excel(f"parts/part_3.xlsx",sheet_name="На XXX",index=False)
# Третий этап
df_ozon_no_xxx = df_ozon.query('deviation_from_opt >= 0.15 & XXX == 0 & margin_in_stock >= 0.05 & '
                               'orders_current_week <= 1 & participation != "no" & participation != "yes2"')
df_ozon_no_xxx.participation = "yes3"
print(df_ozon_no_xxx)
df_ozon_no_xxx.to_excel(f"parts/part_4.xlsx",sheet_name="Не XXX",index=False)
# Нужно обьеденить фреймы df_ozon_ost, df_ozon_xxx, df_ozon_no_xxx и используя ВПР подтянуть данные
# в первоначальный вариант таблицы

# df_ozon_all = pd.DataFrame()
#
#
#
#
# Четвертый этап
df_ozon_rest = df_ozon.query('deviation_from_opt >= 0.15 & margin_in_stock >= 0.1 & participation != "no" & '
                             'participation != "yes2" & participation != "yes3"')
df_ozon_rest.participation = "yes4"
print(df_ozon_rest)
df_ozon_rest.to_excel(f"parts/part_5.xlsx",sheet_name="Остальные",index=False)

# # df_ozon_rest.to_excel("ozon2.xlsx",sheet_name="Остальные",index=False)

# frame = [df_ozon, df_ozon_rest, df_ozon_no_xxx, df_ozon_xxx]
# result = pd.concat(frame)
# # print(df_ozon)
# print(frame)
 # СОХРАНЕНИЕ PD

# df_ozon_no_xxx.to_excel("ozon2.xlsx",sheet_name="Не XXX",index=False)
# df_ozon_rest.to_excel("ozon2.xlsx",sheet_name="Остальные",index=False)