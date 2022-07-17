import pandas as pd
name1='D:\\2021\\基坑监测\\高新安置\\日报\\2019.11.4 - 高兴安置房监测成果表（报告） .xlsm'
pd1 = pd.read_excel(name1)
cols = pd1.columns
numm=0
for i in range(len(cols)):
    if("桩顶竖向位移监测日报表" in cols[i]):#边坡顶部
        numm=numm+1
if(numm==2):
    gc_name1 = "桩顶竖向位移监测日报表"
    gc_name2 = gc_name1 + '.1'
    gc_name3 = "桩顶水平位移监测日报表"
    gc_name4 = gc_name3 + ".1"
    for i in range(0, len(cols)):
        if (cols[i] == gc_name1):
            index1 = i + 2
            break
    for i in range(0, len(cols)):
        if (cols[i] == gc_name2):
            index2 = i + 2
            break
    for i in range(0, len(cols)):
        if (cols[i] == gc_name3):
            index3 = i + 2
            break
    for i in range(0, len(cols)):
        if (cols[i] == gc_name4):
            index4 = i + 2
            break
    #########################################边坡顶部竖向位移监测点日报表
    for i in range(pd1.shape[0]):
        if(isinstance(pd1.iloc[i,index1-2],float)):
            pd1.iloc[i,index1-2]=str(pd1.iloc[i,index1-2])
        if 'CW' in pd1.iloc[i,index1-2]:
            start_index1=i
            break
        else:
            pass
    for i in range(start_index1,pd1.shape[0]):
        if(isinstance(pd1.iloc[i,index1-2],float)):
            pd1.iloc[i,index1-2]=str(pd1.iloc[i,index1-2])
    for i in range(start_index1, pd1.shape[0]):
        if('CW' in pd1.iloc[i,index1-2] and '注' in pd1.iloc[i+1,index1-2]):
            end_index1=i
            break
        if 'CW' not in pd1.iloc[i,index1-2]:
            end_index1=i-1
            break
    #######################################边坡顶部竖向位移监测点日报表.1
    for i in range(pd1.shape[0]):
        if(isinstance(pd1.iloc[i,index2-2],float)):
            pd1.iloc[i,index2-2]=str(pd1.iloc[i,index2-2])
        if 'CW' in pd1.iloc[i,index2-2]:
            start_index2=i
            break
        else:
            pass
    for i in range(start_index2,pd1.shape[0]):
        if(isinstance(pd1.iloc[i,index2-2],float)):
            pd1.iloc[i,index2-2]=str(pd1.iloc[i,index2-2])
    for i in range(start_index2, pd1.shape[0]):
        if('CW' in pd1.iloc[i,index2-2] and '注' in pd1.iloc[i+1,index2-2]):
            end_index2=i
            break
        if 'CW' not in pd1.iloc[i,index2-2]:
            end_index2=i-1
            break
    #######################################边坡顶部水平位移监测日报表
    for i in range(pd1.shape[0]):
        if(isinstance(pd1.iloc[i,index3-2],float)):
            pd1.iloc[i,index3-2]=str(pd1.iloc[i,index3-2])
        if 'CW' in pd1.iloc[i,index3-2]:
            start_index3=i
            break
        else:
            pass
    for i in range(start_index3,pd1.shape[0]):
        if(isinstance(pd1.iloc[i,index3-2],float)):
            pd1.iloc[i,index3-2]=str(pd1.iloc[i,index3-2])
    for i in range(start_index3, pd1.shape[0]):
        if('CW' in pd1.iloc[i,index3-2] and '注' in pd1.iloc[i+1,index3-2]):
            end_index3=i
            break
        if 'CW' not in pd1.iloc[i,index3-2]:
            end_index3=i-1
            break
    ###########################################  边坡顶部水平位移监测日报表.1
    for i in range(pd1.shape[0]):
        if(isinstance(pd1.iloc[i,index4-2],float)):
            pd1.iloc[i,index4-2]=str(pd1.iloc[i,index4-2])
        if 'CW' in pd1.iloc[i,index4-2]:
            start_index4=i
            break
        else:
            pass
    for i in range(start_index4,pd1.shape[0]):
        if(isinstance(pd1.iloc[i,index4-2],float)):
            pd1.iloc[i,index4-2]=str(pd1.iloc[i,index4-2])
    for i in range(start_index4, pd1.shape[0]):
        if('CW' in pd1.iloc[i,index4-2] and '注' in pd1.iloc[i+1,index4-2]):
            end_index4=i
            break
        if 'CW' not in pd1.iloc[i,index4-2]:
            end_index4=i-1
            break
    pd2 = pd1.iloc[start_index1:end_index1+1, index1+1]
    pd3 = pd1.iloc[start_index2:end_index2+1, index2+1]
    pd4 = pd1.iloc[start_index3:end_index3+1, index3+1]
    pd5 = pd1.iloc[start_index4:end_index4+1, index4+1]
    cw_name1=pd1.iloc[start_index1:end_index1+1, index1-2]
    cw_name2=pd1.iloc[start_index2:end_index2+1, index2-2]
    cw_name3=pd1.iloc[start_index3:end_index3+1, index3-2]
    cw_name4=pd1.iloc[start_index4:end_index4+1, index4-2]
    pd_concat1 = pd.concat([pd2, pd3], axis=0)
    pd_concat2 = pd.concat([pd4, pd5], axis=0)
    cw_concat1=pd.concat([cw_name1,cw_name2],axis=0)
    cw_concat2=pd.concat([cw_name3,cw_name4],axis=0)
elif(numm==3):
    gc_name1 = "桩顶竖向位移监测日报表"
    gc_name2 = gc_name1 + '.1'
    gc_name3 = gc_name1 + '.2'
    gc_name4 = "桩顶水平位移监测日报表"
    gc_name5 = gc_name4 + ".1"
    gc_name6 = gc_name4 + ".2"
    for i in range(0, len(cols)):
        if (cols[i] == gc_name1):
            index1 = i + 2
            break
    for i in range(0, len(cols)):
        if (cols[i] == gc_name2):
            index2 = i + 2
            break
    for i in range(0, len(cols)):
        if (cols[i] == gc_name3):
            index3 = i + 2
            break
    for i in range(0, len(cols)):
        if (cols[i] == gc_name4):
            index4 = i + 2
            break
    for i in range(0, len(cols)):
        if (cols[i] == gc_name5):
            index5 = i + 2
            break
    for i in range(0, len(cols)):
        if (cols[i] == gc_name6):
            index6 = i + 2
            break
    #########################################边坡顶部竖向位移监测点日报表
    for i in range(pd1.shape[0]):
        if(isinstance(pd1.iloc[i,index1-2],float)):
            pd1.iloc[i,index1-2]=str(pd1.iloc[i,index1-2])
        if 'CW' in pd1.iloc[i,index1-2]:
            start_index1=i
            break
        else:
            pass
    for i in range(start_index1,pd1.shape[0]):
        if(isinstance(pd1.iloc[i,index1-2],float)):
            pd1.iloc[i,index1-2]=str(pd1.iloc[i,index1-2])
    for i in range(start_index1, pd1.shape[0]):
        if('CW' in pd1.iloc[i,index1-2] and '注' in pd1.iloc[i+1,index1-2]):
            end_index1=i
            break
        if 'CW' not in pd1.iloc[i,index1-2]:
            end_index1=i-1
            break
    #######################################边坡顶部竖向位移监测点日报表.1
    for i in range(pd1.shape[0]):
        if(isinstance(pd1.iloc[i,index2-2],float)):
            pd1.iloc[i,index2-2]=str(pd1.iloc[i,index2-2])
        if 'CW' in pd1.iloc[i,index2-2]:
            start_index2=i
            break
        else:
            pass
    for i in range(start_index2,pd1.shape[0]):
        if(isinstance(pd1.iloc[i,index2-2],float)):
            pd1.iloc[i,index2-2]=str(pd1.iloc[i,index2-2])
    for i in range(start_index2, pd1.shape[0]):
        if('CW' in pd1.iloc[i,index2-2] and '注' in pd1.iloc[i+1,index2-2]):
            end_index2=i
            break
        if 'CW' not in pd1.iloc[i,index2-2]:
            end_index2=i-1
            break
    #######################################边坡顶部竖向位移监测点日报表.2
    for i in range(pd1.shape[0]):
        if(isinstance(pd1.iloc[i,index3-2],float)):
            pd1.iloc[i,index3-2]=str(pd1.iloc[i,index3-2])
        if 'CW' in pd1.iloc[i,index3-2]:
            start_index3=i
            break
        else:
            pass
    for i in range(start_index3,pd1.shape[0]):
        if(isinstance(pd1.iloc[i,index3-2],float)):
            pd1.iloc[i,index3-2]=str(pd1.iloc[i,index3-2])
    for i in range(start_index3, pd1.shape[0]):
        if('CW' in pd1.iloc[i,index3-2] and '注' in pd1.iloc[i+1,index3-2]):
            end_index3=i
            break
        if 'CW' not in pd1.iloc[i,index3-2]:
            end_index3=i-1
            break
    ###########################################边坡顶部水平位移监测日报表
    for i in range(pd1.shape[0]):
        if(isinstance(pd1.iloc[i,index4-2],float)):
            pd1.iloc[i,index4-2]=str(pd1.iloc[i,index4-2])
        if 'CW' in pd1.iloc[i,index4-2]:
            start_index4=i
            break
        else:
            pass
    for i in range(start_index4,pd1.shape[0]):
        if(isinstance(pd1.iloc[i,index4-2],float)):
            pd1.iloc[i,index4-2]=str(pd1.iloc[i,index4-2])
    for i in range(start_index4, pd1.shape[0]):
        if('CW' in pd1.iloc[i,index4-2] and '注' in pd1.iloc[i+1,index4-2]):
            end_index4=i
            break
        if 'CW' not in pd1.iloc[i,index4-2]:
            end_index4=i-1
            break
    ##########################################边坡顶部水平位移监测日报表.1
    for i in range(pd1.shape[0]):
        if(isinstance(pd1.iloc[i,index5-2],float)):
            pd1.iloc[i,index5-2]=str(pd1.iloc[i,index5-2])
        if 'CW' in pd1.iloc[i,index5-2]:
            start_index5=i
            break
        else:
            pass
    for i in range(start_index5,pd1.shape[0]):
        if(isinstance(pd1.iloc[i,index5-2],float)):
            pd1.iloc[i,index5-2]=str(pd1.iloc[i,index5-2])
    for i in range(start_index5, pd1.shape[0]):
        if('CW' in pd1.iloc[i,index5-2] and '注' in pd1.iloc[i+1,index5-2]):
            end_index5=i
            break
        if 'CW' not in pd1.iloc[i,index5-2]:
            end_index5=i-1
            break
    ##########################################边坡顶部水平位移监测日报表.2
    for i in range(pd1.shape[0]):
        if(isinstance(pd1.iloc[i,index6-2],float)):
            pd1.iloc[i,index6-2]=str(pd1.iloc[i,index6-2])
        if 'CW' in pd1.iloc[i,index6-2]:
            start_index6=i
            break
        else:
            pass
    for i in range(start_index6,pd1.shape[0]):
        if(isinstance(pd1.iloc[i,index6-2],float)):
            pd1.iloc[i,index6-2]=str(pd1.iloc[i,index6-2])
    for i in range(start_index6, pd1.shape[0]):
        if('CW' in pd1.iloc[i,index6-2] and '注' in pd1.iloc[i+1,index6-2]):
            end_index6=i
            break
        if 'CW' not in pd1.iloc[i,index6-2]:
            end_index6=i-1
            break
    ##########################################
    pd2 = pd1.iloc[start_index1:end_index1+1, index1+1]
    pd3 = pd1.iloc[start_index2:end_index2+1, index2+1]
    pd4 = pd1.iloc[start_index3:end_index3+1, index3+1]
    pd5 = pd1.iloc[start_index4:end_index4+1, index4+1]
    pd6 = pd1.iloc[start_index5:end_index5+1, index5+1]
    pd7 = pd1.iloc[start_index6:end_index6+1, index6+1]
    cw_name1=pd1.iloc[start_index1:end_index1+1, index1-2]
    cw_name2=pd1.iloc[start_index2:end_index2+1, index2-2]
    cw_name3=pd1.iloc[start_index3:end_index3+1, index3-2]
    cw_name4=pd1.iloc[start_index4:end_index4+1, index4-2]
    cw_name5=pd1.iloc[start_index5:end_index5+1, index5-2]
    cw_name6=pd1.iloc[start_index6:end_index6+1, index6-2]
    pd_concat1 = pd.concat([pd2, pd3,pd4], axis=0)
    pd_concat2 = pd.concat([pd4, pd5,pd6], axis=0)
    cw_concat1=pd.concat([cw_name1,cw_name2],axis=0)
    cw_concat2=pd.concat([cw_name3,cw_name4],axis=0)