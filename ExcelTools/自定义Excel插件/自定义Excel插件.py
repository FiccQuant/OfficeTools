import time
import plotly.express as px
import pandas as pd
import xlwings as xw
from datetime import datetime

# 自定义的Excel函数show，生成时序图，按指定的间隔刷新
@xw.func(async_mode='threading')
@xw.arg("cell",doc="单元格编号")
@xw.arg("title",doc="标题")
def show(cell,title):
    wb=xw.books.active
    sht=wb.sheets.active
    df=pd.DataFrame(columns=["时间","价差"])
    # 缓存从Excel指定单元格获得的当前价差
    df=df.append({"时间":datetime.now(),"价差":sht.range(cell).value},ignore_index=True)
    fig=px.line(df,x="时间",y="价差",title=title)
    p=sht.pictures.add(fig,name=title,update=True,left=500,top=100)
    while True:
        # 设置间隔多久刷新一次
        time.sleep(2)
        # 缓存从Excel指定单元格获得的当前价差
        df=df.append({"时间":datetime.now(),"价差":sht.range(cell).value},ignore_index=True)
        fig=px.line(df,x="时间",y="价差",title=title)
        # 更新时序图
        p=p.update(fig)