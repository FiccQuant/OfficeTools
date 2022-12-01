# -*- coding: utf-8 -*-
"""
Created on Tue Nov 29 08:52:08 2022

@author: admin
"""

import pandas as pd
import xlwings as xw

if __name__=="__main__":
    # 使用pandas加载交易明细
    df=pd.read_excel("./交易明细.xlsx")
    # 交易合并：按照日期、交易对手、交易员、合约、买卖方向、交易单位分组汇总成交量、成交金额
    df=df.groupby(["日期","交易对手","交易员","合约","买卖","交易单位"]).agg({"成交量":sum,"成交金额":sum})
    df["成交价"]=df["成交金额"]/df["成交量"]
    df.reset_index(inplace=True)
    # 生成汇总统计（可选）
    df.to_excel("./汇总统计.xlsx",index=None,float_format="%.4f")
    # 写入剪贴板。执行后可以通过黏贴（Ctrl-V）到Excel中（可选）
    df.to_clipboard(index=None)
    # 加载Excel
    app=xw.App(visible=False)
    # 加载审批单模板
    wb=app.books.open("./审批单模板.xlsx")
    # 选择审批单模板在Excel中的Sheet
    sht=wb.sheets[0]
    # 将交易明细写入模板
    sht.range('B5').value = df[['日期','合约','买卖','成交量','交易单位','成交金额','成交价']].values
    # 保存、生成处理后的审批单
    wb.save("./审批单.xlsx")
    wb.close()
    app.quit()