import pandas as pd
import plotly.offline as py

if __name__=="__main__":
    # 指定pandas使用plotly绘图
    pd.options.plotting.backend="plotly"
    # 使用pandas加载交易明细
    trades=pd.read_excel("./交易明细.xlsx",dtype={"日期":"string"})
    # 加载每日的估值价格
    prices=pd.read_excel("./每日估值价格.xls",dtype={"日期":"string"})
    # 计算每笔交易成交金额
    trades["金额"]=-trades["成交量"]*trades["成交价"]
    # 按日期和合约汇总成交量和金额
    pos=trades.groupby(["日期","合约"])[["成交量","金额"]].sum().reset_index().sort_values("日期")
    # 计算持仓和累计成交金额
    pos[["累计成交量","累计金额"]]=pos.groupby(["合约"])[["成交量","金额"]].transform("cumsum")
    # 匹配交易中的合约对应日期的估值价格
    df=pd.merge(prices,pos[["日期","合约","累计成交量","累计金额"]],on=["日期","合约"],how="outer")
    # 填充未发生交易日期持仓
    df[["累计成交量","累计金额"]]=df.groupby(["合约"])[["累计成交量","累计金额"]].transform(lambda x:x.fillna(method="ffill"))
    df=df.dropna()
    # 计算每笔交易的损益
    df["损益"]=df["累计金额"]+df["累计成交量"]*df["价格"]
    # 以日期为基准，将合约及合约的损益展开至列
    res_df=df.pivot(index="日期",columns=["合约"],values=["损益"])
    res_df.columns=res_df.columns.droplevel(0)
    res_df=res_df.fillna(0)
    # 绘制规制结果曲线
    ax=res_df.plot(title="交易估值")
    ax.layout["yaxis"]["title"]["text"]="估值损益"
    py.plot(ax)
    res_df.reset_index(inplace=True)
    # 生成估值结果Excel
    res_df.to_excel("./估值结果.xlsx",index=None)