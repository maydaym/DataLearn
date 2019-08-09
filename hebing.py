# import pandas as pd
# df=pd.DataFrame(['asda'],['asdas'])
# print(df)

import pandas as pd
import os
file=r'C:\Users\Administrator\Desktop\test'
total_table=pd.DataFrame(columns=['总后大类','总后子类','品牌','产品名称','关键参数','电商价格/合同平均价','电商链接','其他渠道证明','审核意见','发件人邮箱'])
for parents,dirnames,filenames in os.walk(file):
    for filename in filenames:
        df=pd.read_excel(os.path.join(parents,filename))
        df1=df.drop(index=0)
        total_table=total_table.append(df1,ignore_index=True)
total_table.to_excel(r'E:\star1\333.xlsx')
