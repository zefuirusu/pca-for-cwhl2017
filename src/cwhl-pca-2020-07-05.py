#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Jun  2 14:51:19 2019

@author: 一介貂蝉

面向过程开发。

"""

# 1.Import required libs
import xlrd
import xlwt
import numpy as np
import pandas as pd
import matplotlib as mpl
mpl.rcParams['font.sans-serif'] = ['FangSong']
# from sklearn import preprocessing

fdir=r'./data/GLofChewang2017.xlsx'
expdir=r'./data/ExpenseofCWHL2017.xlsx'

# 2.Read and transfer general ledgers into matrix of expense data.

gl=pd.read_excel(fdir,sheet_name='管理费用GL')
gl_good1=gl.loc[:,['具体科目','金额']]
gl_good2=gl.drop(labels=['具体科目','金额'],axis=1,inplace=False)

names=gl.loc[:,'具体科目'].drop_duplicates()
na_list=names.reset_index(drop=True)
na_list=list(na_list)
gl_good2=pd.DataFrame(gl_good2,columns=na_list)

for i in na_list:
    a=gl_good1[gl_good1['具体科目']==i]
    gl_good2.loc[a.index,i]=a.loc[:,'金额']
print('原始数据阵的形状和金额合计：',gl_good2.shape,gl_good2.sum().sum())

good=gl_good2.fillna(0)

# 3.Hierarchical Clustering

# # Standardize (optional)
# from sklearn import preprocessing
# good_standard=preprocessing.scale(good.values,axis=0)
# good_standard=pd.DataFrame(good_standard)
# good_standard.columns=good.columns

# import scipy
# from scipy import cluster
# from scipy.cluster import hierarchy
from scipy.cluster.hierarchy import linkage,dendrogram

# a=scipy.spatial.distance.pdist(good.T.values, metric='cosine')
# 采用其他距离也可以：mahalanobis马哈拉诺比斯距离，minkowski明可夫斯基距离。
z=linkage(good.T,method='average',metric='minkowski') #ward法最完美，但不适用这种对变量聚类的情况，首选single/average法。
f=dendrogram(z,orientation='right',show_leaf_counts=True)
f=pd.Series(f)
lf=f['leaves']

li=list(lf)
li1=li[0:7:1]
li2=li[7::1]
g1=good.iloc[:,li1]
g2=good.iloc[:,li2]
print('验证将原始数据阵拆分为未聚类的g1和聚类的g2之和与原始数据阵金额总和是否相等：',g1.sum().sum()+g2.sum().sum())
oth_na=list(g2.columns)
print('g2中的科目合并为科目others(oth_na)：',oth_na)
g2.loc[:,'others']=g2.sum(axis=1)
oth=g2.loc[:,'others']
oth=list(oth)

g1.loc[:,'others']=oth
print('将合并科目others再合并到g1，得到聚类整理之后的数据阵g1的金额合计:',g1.sum().sum())
print(g1.head(3))

# 4.Principal Component Analysis and Conclusion

# # Standardize (optional)
# na=list(g1.columns)
# from sklearn import preprocessing
# g1=preprocessing.scale(g1.values,axis=0)
# g1=pd.DataFrame(g1)
# g1.columns=na

sgm=g1.cov()
print('协方差矩阵：',sgm.shape)
tzz,tzxl=np.linalg.eigh(sgm)
tzxl=pd.DataFrame(tzxl)
tzz=pd.Series(tzz)
tzzsort=tzz.sort_values(ascending=False).round(6)

gxl=tzzsort/tzzsort.sum()
gxl_cu=gxl.cumsum()
gxl_cu=pd.DataFrame(gxl_cu,columns=['gxl_cu'])
tzz_se=tzzsort[gxl_cu[gxl_cu['gxl_cu']<0.92].index]
print('累计方差贡献率:',(tzz_se.cumsum()/tzzsort.sum()).round(3))
print('符合条件的特征值:',tzz_se.round(3))

tzxl_se=tzxl.iloc[:,tzz_se.index]
na=list(g1.columns) #前边定义过na了。
tzxl_se.index=na

# Conclusion
bdyy=tzxl_se.abs().idxmax(axis=0)
bdyy=pd.DataFrame(bdyy)
print('变动原因',bdyy)

expdf=pd.read_excel(expdir)
expdf.set_index(['月份'],inplace=True)
print('验证金额是否相等：',expdf.sum().sum(),gl['金额'].sum(),good.sum().sum())
print('验证科目列表是否一致')
def compareli(li1=[],li2=[]):
    '''li1 and li2 are both list-type.验证科目列表是否一致'''
    shared=[]
    li1priv=[]
    li2priv=[]
    for i in li1:
        if i in li2:
            shared.append(i)
        else:
            li1priv.append(i)
    for i in li2:
        if i in li1:
            if i in shared:
                pass
            else:
                shared.append(i)
        else:
            li2priv.append(i)
    resu=[shared,li1priv,li2priv]
    return resu
compData=compareli(list(expdf.columns),oth_na)
print('in common:',compData[0])
print('expdf_columns_unique:',compData[1])
print('oth_na_unique:',compData[2])
expdf_oth=expdf.loc[:,compData[0]]
# print('expdf_oth:',type(expdf_oth),'\n',expdf_oth)
expdf_oth_sum=pd.DataFrame(expdf_oth.sum(axis=1),columns=['others'],index=expdf_oth.index)
print(expdf_oth_sum)

# 5.Visualization

bdyylist=list(bdyy[0])
print('展示如下变量的折线图',bdyylist)

compbdyy=compareli(list(expdf.columns),bdyylist)
print(compbdyy[0])
print(compbdyy[1])
print(compbdyy[2])
if len(compbdyy[2])==0:
    expdf_bdyylist=expdf.loc[:,bdyylist]
else:
    expdf_bdyylist=expdf.loc[:,compbdyy[0]]
    if 'others' in bdyylist:
        expdf_bdyylist.loc[:,'others']=expdf_oth_sum.loc[:,'others']
    else:
        pass
fi=expdf_bdyylist

fi.plot.line(legend=True,table=True,figsize=(14,6),use_index=True,grid=True,sort_columns=True)
fi.plot.box(legend=True,figsize=(14,6),use_index=True,grid=True,sort_columns=True)
fi.plot.hist(legend=True,figsize=(14,6),use_index=True,grid=True,sort_columns=True)

# ————————————————
# 版权声明：本文为CSDN博主「一介貂蝉Phantaska」的原创文章，遵循 CC 4.0 BY-SA 版权协议，转载请附上原文出处链接及本声明。
# 原文链接：https://blog.csdn.net/weixin_44588870/article/details/89462568
