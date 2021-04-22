import pandas as pd
from pathlib import Path

xlsxs = ['初审20210406.xlsx', '电核20210406.xlsx', '合同明细20210406.xlsx']

cwd = Path("C:/code/py/code")

first_check_df = pd.read_excel(cwd.joinpath(xlsxs[0]))  # 初审20210406
phone_check_df = pd.read_excel(cwd.joinpath(xlsxs[1]))  # 电核20210406
contract_df = pd.read_excel(cwd.joinpath(xlsxs[2]))  # 合同

# union 初审和电核的表
check_df = pd.concat([first_check_df, phone_check_df], ignore_index=True)
# 按照订单编号升序和审批时间降序排列后去重订单编号,只保留最新状态的订单
sorted_check_df = check_df.sort_values(by=['订单编号', '审批时间'], ascending=[True, False]).drop_duplicates('订单编号',
                                                                                                     keep='first')
print(sorted_check_df.shape)
# 合并明细去重
sorted_contract_df = contract_df.sort_values(by=['订单编号'], ascending=[True]).drop_duplicates('订单编号', keep='first')
print(sorted_contract_df.shape)

# 审核表和合同表交集
merged_df = pd.merge(sorted_check_df, sorted_contract_df, how='inner', on='订单编号')
print(merged_df.shape)
stats_count = merged_df.groupby(by=['审批结果']).count()
print(stats_count.head(100))
pd.write
