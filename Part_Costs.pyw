import _mssql
import os
import pandas as pd
import pandas.io.sql
import pymssql
import tkinter as tk
import xlsxwriter

master = tk.Tk()
tk.Label(master, text="                              Part Number").grid(row=2, column=1)

e = tk.Entry(master)
e.place(x=25, y=30, width=200)
e.focus_set()

conn = pymssql.connect(
    host=r"host",
    user=r"user",
    password="password",
    database="database"
)

def costsAndHours():
    part = str(e.get())
    testSql = "SELECT * FROM WORK_ORDER WHERE PART_ID = '" + part + "'"
    testDf = pd.io.sql.read_sql(testSql, conn)
    pdDf = pd.DataFrame(testDf)
    if pdDf.empty:
        top = tk.Toplevel(master)
        top.title('Error')
        msg = tk.Message(top, text="Error! Part number provided does not exist in the database.", width=750)
        msg.grid(row=0, column=1)
        return
    else:
      file = "K:/" + part + ".xlsx"
      writer = pd.ExcelWriter(file)

      #Average Costs
      avgCostSql = """SELECT AVG(EST_MATERIAL_COST) AS 'AVG EST MAT COST',
                     AVG(ACT_MATERIAL_COST) AS 'AVG MAT COST',
                     (AVG(EST_MATERIAL_COST) - AVG(ACT_MATERIAL_COST)) AS 'AVG MAT DIFFERENCE',
                     AVG(EST_LABOR_COST) AS 'AVG EST LAB COST',
                     AVG(ACT_LABOR_COST) AS 'AVG LAB COST',
                     (AVG(EST_LABOR_COST) - AVG(ACT_LABOR_COST)) AS 'AVG LAB DIFFERENCE',
                     AVG(EST_BURDEN_COST) AS 'AVG EST BUR COST',
                     AVG(ACT_BURDEN_COST) AS 'AVG BUR COST',
                     (AVG(EST_BURDEN_COST) - AVG(ACT_BURDEN_COST)) AS 'AVG BUR DIFFERENCE',
                     AVG(EST_SERVICE_COST) AS 'AVG EST SER COST',
                     AVG(ACT_SERVICE_COST) AS 'AVG SER COST',
                     (AVG(EST_SERVICE_COST) - AVG(ACT_SERVICE_COST)) AS 'AVG SER DIFFERENCE',
                     ((AVG(EST_MATERIAL_COST) - AVG(ACT_MATERIAL_COST)) + (AVG(EST_LABOR_COST) - AVG(ACT_LABOR_COST)) + (AVG(EST_BURDEN_COST) - AVG(ACT_BURDEN_COST)) + (AVG(EST_SERVICE_COST) - AVG(ACT_SERVICE_COST))) AS 'AVG TOTAL DIFF'
                      FROM WORK_ORDER
                      WHERE PART_ID = '""" + part + """'
                      AND TYPE = 'W'"""
      df = pd.io.sql.read_sql(avgCostSql, conn)
      df.to_excel(writer, 'Average Costs', index=False)

      #Total Costs
      totalCostSql = """SELECT BASE_ID,
                       LOT_ID,
                       SPLIT_ID,
                       SUM(EST_MATERIAL_COST) AS 'EST MAT COST',
                       SUM(ACT_MATERIAL_COST) AS 'ACT MAT COST',
                       (SUM(EST_MATERIAL_COST) - SUM(ACT_MATERIAL_COST)) AS 'MAT DIFF',
                       SUM(EST_LABOR_COST) AS 'EST LAB COST',
                       SUM(ACT_LABOR_COST) AS 'ACT LAB COST',
                       (SUM(EST_LABOR_COST) - SUM(ACT_LABOR_COST)) AS 'LABOR DIFF',
                       SUM(EST_BURDEN_COST) AS 'EST BUR COST',
                       SUM(ACT_BURDEN_COST) AS 'ACT BUR COST',
                       (SUM(EST_BURDEN_COST) - SUM(ACT_BURDEN_COST)) AS 'BUR DIFF',
                       SUM(EST_SERVICE_COST) AS 'EST SER COST',
                       SUM(ACT_SERVICE_COST) AS 'ACT SER COST',
                       (SUM(EST_SERVICE_COST) - SUM(ACT_SERVICE_COST)) AS 'SER DIFF',
                       ((SUM(EST_MATERIAL_COST) - SUM(ACT_MATERIAL_COST)) + (SUM(EST_LABOR_COST) - SUM(ACT_LABOR_COST)) + (SUM(EST_BURDEN_COST) - SUM(ACT_BURDEN_COST)) + (SUM(EST_SERVICE_COST) - SUM(ACT_SERVICE_COST))) AS 'TOTAL DIFF'
                        FROM WORK_ORDER
                        WHERE PART_ID = '""" + part + """'
                        AND TYPE = 'W'
                        GROUP BY BASE_ID,
                         LOT_ID,
                         SPLIT_ID"""
      df2 = pd.io.sql.read_sql(totalCostSql, conn)
      df2.to_excel(writer, 'Total Costs', index=False)

      #Average Hours
      avgHoursSql = """SELECT AVG(O.RUN_HRS) AS 'ACT AVG UNIT HOURS',
                       AVG(O.ACT_RUN_HRS) AS 'EST AVG UNIT HOURS',
                       (AVG(O.RUN_HRS) - AVG(O.ACT_RUN_HRS)) AS DIFFERENCE
                        FROM OPERATION O
                        INNER JOIN WORK_ORDER W ON O.WORKORDER_BASE_ID = W.BASE_ID
                        AND O.WORKORDER_LOT_ID = W.LOT_ID
                        AND O.WORKORDER_SPLIT_ID = W.SPLIT_ID
                        WHERE W.PART_ID = '""" + part + """'
                        AND WORKORDER_TYPE = 'W'"""
      df3 = pd.io.sql.read_sql(avgHoursSql, conn)
      df3.to_excel(writer, 'Average Hours', index=False)

      #Total Hours
      totalHoursSql = """SELECT O.WORKORDER_BASE_ID,
                         O.WORKORDER_LOT_ID,
                         O.WORKORDER_SPLIT_ID, O.RESOURCE_ID,
                         SUM(O.RUN_HRS) AS 'ACT TOTAL UNIT HOURS',
                         SUM(O.ACT_RUN_HRS) AS 'EST TOTAL UNIT HOURS',
                         (SUM(O.RUN_HRS) - SUM(O.ACT_RUN_HRS)) AS DIFFERENCE
                          FROM OPERATION O
                          INNER JOIN WORK_ORDER W ON O.WORKORDER_BASE_ID = W.BASE_ID
                          AND O.WORKORDER_LOT_ID = W.LOT_ID
                          AND O.WORKORDER_SPLIT_ID = W.SPLIT_ID
                          WHERE W.PART_ID = '""" + part + """'
                          AND WORKORDER_TYPE = 'W'
                          GROUP BY O.WORKORDER_BASE_ID,
                           O.WORKORDER_LOT_ID,
                           O.WORKORDER_SPLIT_ID, O.RESOURCE_ID"""
      df4 = pd.io.sql.read_sql(totalHoursSql, conn)
      df4.to_excel(writer, 'Total Hours', index=False)
      writer.save()
      os.startfile(file)

menubar = tk.Menu(master)
menubar.add_command(label="Run", command=costsAndHours)
menubar.add_command(label="Quit", command=master.quit)
master.config(menu=menubar)
master.title("Part Cost Report")
master.minsize(width=270, height=70)
master.mainloop()
