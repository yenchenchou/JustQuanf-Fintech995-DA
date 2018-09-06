#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon May 28 13:48:15 2018

@author: yc
"""
import pandas as pd
import numpy as np
from pandas import Series, DataFrame
import matplotlib.pyplot as plt
import seaborn as sns
from statsmodels.stats.proportion import proportions_ztest
from bs4 import BeautifulSoup
import re
import xlsxwriter
###################################################
############ read the registered data #############
data = pd.ExcelFile('/Users/yc/Desktop/AppWorks_StartupHackers/芬特克歷史報名資料/180611_芬特克整理LIST.xlsx')
print(data.sheet_names)
excel_list = ['428高雄場', '411台北場', '41台北場',
              '0114台中場', '113新竹場', '17新竹場', 
              '0106台中場', '1227台北場', '1217台北場', 
              '1126桃園場', '1105高雄場', '1025台北加開場',
              '1015台北場', '1008台中場'
              ]

def get_excel_data(excel_list):
    
    df=[]
    for sheet in excel_list:
        df.append(data.parse(sheet))
    return df

df = get_excel_data(excel_list)
register_df = pd.concat(df,axis=0)
#register_df.to_csv(path_or_buf='/Users/yc/Desktop/register_df.csv')

############ basic information from the data #############
register_df.info()
register_df.isnull().sum()
register_df.head()
register_df.columns
############ check if the value of each column does corresponded to what the column itself ############
def val_in_col():
    col_list = register_df.columns
    col_list = ['Variation', '報名免費講座場次',
                '從哪裡得知講座消息？', '是否到場', 
                '職業 / 職稱', '薪資水準區間（月薪）']
    
    for check_column_value in col_list:
        print(register_df[check_column_value].value_counts() , '\n\n')
        
val_in_col()

register_df['報名免費講座場次'].str.contains('5/5').sum()
register_df = register_df[register_df['報名免費講座場次']!='高雄加開場｜ 5/5 （六）高雄商務會議中心三樓']
register_df = register_df[register_df['報名免費講座場次']!='高雄加開場｜ 5/5 （六）英格迪酒店']
register_df = register_df[register_df['報名免費講座場次']!='高雄加開場｜ 5/5 （六）英迪格酒店']
############ 重複報名的人到場率是不是比較高？ ############
def double_register():
    duplcate_data = register_df[register_df.duplicated(subset='Email')]
    """
    報名到場率36.20 %，有重複報名的名單到場率僅為13.63 %
    """
    print('報名總到場率: ' , '{:.2f}'.format(len(register_df[register_df['是否到場']=='是'])/len(register_df)*100),'%') #報名到場率36.20 %
    print('重複報名到場率: ' , '{:.2f}'.format(len(duplcate_data['是否到場']=='是')/len(register_df['是否到場']=='是')*100),'%') #重複報名到場率13.63 %
double_register()

register_df = register_df.drop_duplicates()#register_df.to_csv(path_or_buf='/Users/yc/Desktop/registered_df.csv')

############ 哪個報名來源有較高到場率？ ＋處理講座消息來源資料 ############

def unique_valuein_register():   
    """觀察誰要當報名資料跟上課學員資料的primary key"""
    
    a = ['姓名','Email','手機號碼']
    for i in a:
        print(len(register_df[i].unique()),'\n')
        
unique_valuein_register()
register_df = register_df.drop_duplicates(subset=['Email'], keep="first")
register_df = register_df.drop_duplicates(subset=['手機號碼'], keep="first")

register_df['從哪裡得知講座消息？'].unique()
register_df.at[register_df['報名免費講座場次']=='名單型','從哪裡得知講座消息？'] = '名單型'
register_df.at[register_df['報名免費講座場次']=='活動通','從哪裡得知講座消息？'] = 'Accupass 活動通'
register_df['從哪裡得知講座消息？'] = register_df['從哪裡得知講座消息？'].fillna('Accupass 活動通')
register_df['從哪裡得知講座消息？'] = register_df['從哪裡得知講座消息？'].replace('Facebook  社團','Facebook 社團')

def each_pct():    
    """算出各個源到場率 vs 總報到場率"""
    
    source = register_df['從哪裡得知講座消息？'].unique()
    numerator = ["a%d" %i for i in np.arange(len(source))]   
    denominator = ["b%d" %i for i in np.arange(len(source))]  
    dictionary = dict(zip(numerator, denominator))
    
    for source_name in source:
        for numerator, denominator in dictionary:            
            numerator = (register_df['從哪裡得知講座消息？'] == source_name) & (register_df['是否到場'] == '是')
            denominator = register_df['從哪裡得知講座消息？'] == source_name            
        print(source_name ,'到場率: ', '{:.2f}'.format(numerator.sum() / denominator.sum()*100), '%')
    print('報名總到場率: ' , '{:.2f}'.format(len(register_df[register_df['是否到場']=='是'])/len(register_df)*100),'%') #報名到場率36.20 %
    
each_pct()


############ 不留lineID的人到場率會比較差？ ############
def People_no_line():
    
    """ 
    不留lINE的報名狀況：
    先計算沒有line總數，
    後計算是因為accupass或是名單型廣告造成沒有line
    """
    
    noline = register_df['Line ID '].isnull().sum()
    missingline_from_accupass = register_df[register_df['從哪裡得知講座消息？']=='Accupass 活動通']['Line ID '].isnull().sum()
    missingline_from_namelist = register_df[register_df['從哪裡得知講座消息？']=='名單型']['Line ID '].isnull().sum()
    LINE_USAGE_PCT = (noline-missingline_from_accupass-missingline_from_namelist)/len(register_df)
    return '{:.2f}'.format(LINE_USAGE_PCT*100)

print("不用/不留Line的比例:", People_no_line() , '%')

def P_ztest():
    
    """ 計算有無提供是否影響報名狀況 """
    
    a=register_df[register_df['Line ID '].isnull()]
    line1 = a['是否到場'].value_counts()/len(a['是否到場'])
    line2 = register_df['是否到場'].value_counts()/len(register_df['是否到場'])
    line1 = np.array(line1)
    line2 = np.array(line2)
    stat, pval = proportions_ztest(line1, line2)##有問題
    return pval

print(P_ztest(), ' => 不給line到場比率低')



############ 清理register_df剩下資料 ############
register_df.info()
#register_df = register_df.drop(['Line ID ', 'ip'], axis=1)
register_df.isnull().sum()

"""
清理資料順序：
時間 & 報名免費講座場次 向前補上(ffill)
報名免費講座場次去掉‘名單型’、‘活動通’等不正確字詞＋報名免費講座場次標準化
職稱標準化
Variation ＆ 標準化職稱 ＆ 薪資水準區間（月薪）用fancy impute KNN解決

"""
def fill_accuapass_listname_Date():
    register_df[register_df['報名免費講座場次']=='名單型'] = register_df[register_df['報名免費講座場次']=='名單型'].replace('名單型', np.NaN)
    register_df[register_df['報名免費講座場次']=='活動通'] = register_df[register_df['報名免費講座場次']=='活動通'].replace('活動通', np.NaN)
    register_df[['Date','報名免費講座場次']] = register_df[['Date','報名免費講座場次']].fillna(method='ffill')
    any(register_df['報名免費講座場次'].isnull())
fill_accuapass_listname_Date()    
#c = pd.Series(register_df['報名免費講座場次'].unique())


def split_col_seminar():
    
    global register_df
    split1 = register_df['報名免費講座場次'].str.split('｜', expand=True)
    split2 = split1[1].str.split('（', expand=True)
    split3 = split2[1].str.split('）', expand=True)
    register_df = pd.concat([register_df, split1[0], split2[0], split3], axis=1)
    register_df.columns = ['Date',  'Variation', '姓名', '手機號碼', 'Email', '報名免費講座場次',
                           '薪資水準區間（月薪）', '職業', '從哪裡得知講座消息？', 
                           '是否到場', '場次', '日期', '星期', '場地']

    word_df=[]
    for word in register_df['場次']:
        word_df.append(re.findall('^..', word))
    register_df['場次'] = np.array(word_df)

    register_df['星期'].unique()
    register_df['場地'].unique()
    register_df['場地'] = register_df['場地'].replace('益品書店','益品書屋')
    register_df['場地'] = register_df['場地'].replace('大倉久和大飯店','大倉久和')
    register_df['場地'] = register_df['場地'].replace('英格迪酒店','英迪格酒店')

split_col_seminar()


def salary():
    d = list(register_df['薪資水準區間（月薪）'].unique())
    register_df['薪資水準區間（月薪）'] = register_df['薪資水準區間（月薪）'].replace('$30,000~50,000','$30,000～50,000')
    register_df['薪資水準區間（月薪）'] = register_df['薪資水準區間（月薪）'].replace('$50,000~100,000','$50,000～100,000')
salary()
#register_df.to_csv(path_or_buf='/Users/yc/Desktop/register_df.csv')
#register_df2 = pd.read_csv('/Users/yc/Desktop/芬特克歷史報名資料/register_df 2.csv')

def attendence_situation():
    """ 報名率與報退率趨勢 """
    a = register_df.columns
    attendence_situation_list = ['Variation', '薪資水準區間（月薪）',
                                 '從哪裡得知講座消息？','日期',
                                 '場次', '星期', '場地']
    
    for attendence_situation in register_df[attendence_situation_list]:
        print(register_df.pivot_table(values = '職業', 
                                     index=attendence_situation, 
                                     columns='是否到場', 
                                     aggfunc='count',
                                     margins=True), '\n\n')
attendence_situation()

register_df = register_df.drop(['Date','報名免費講座場次','職業'],axis=1)
register_df[register_df['是否到場']=='是'].count()
register_df.isnull().sum()

###################################################
############ read the course data #################
course_df = pd.read_csv('/Users/yc/Desktop/芬特克歷史報名資料/180606-芬特克學員資料-v1r03.csv')
course_df = course_df.drop_duplicates()
course_df = course_df.drop_duplicates(subset=['Email'],keep='first')
course_df = course_df.rename(columns={'聯絡電話':'手機號碼'})
course_df = course_df.drop_duplicates(subset=['手機號碼'],keep='first')

def unsucribe_course_ratio():      
    """ 課程總報退率 """    
    
    return '課程總報退率： ' + '{:.2f}'.format(len(course_df[course_df['是否報退']=="是"])/course_df['是否報退'].notnull().sum()*100) + '%'
    
unsucribe_course_ratio()


###################################################################
############ COMBINE DATA register_df & course_df #################
result = pd.merge(register_df,course_df, on=['手機號碼'], how='outer')
result['是否報退'].notnull().sum()
result.isnull().sum()
#result.to_csv(path_or_buf='/Users/yc/Desktop/180613-result-v1r10.csv')

df = pd.read_csv('/Users/yc/Desktop/AppWorks_StartupHackers/芬特克歷史報名資料/180613-result-v1r10.csv')
df.info()
df = df.drop_duplicates()#register_df.to_csv(path_or_buf='/Users/yc/Desktop/registered_df.csv')
df.isnull().sum()
df['是否報退'].notnull().sum()
df = df.drop_duplicates()
df = df.drop_duplicates(subset=['手機號碼','Email_x'],keep='first')

#填補因為重複筆數被刪掉的資料
df = df.drop(['Unnamed: 0','姓名_x','Email_x','學員編號','姓名_y','Email_y','Line ID','實際報名日期'],axis=1)
[df[x].unique() for x in df.columns]


#Variation:
df['Variation'].isnull().sum()
df['Variation'].value_counts()
df['Variation'] = df['Variation'].fillna('Unknown')

#薪資水準
df['薪資水準區間（月薪）'].isnull().sum()
df['薪資水準區間（月薪）'].value_counts()
df['薪資水準區間（月薪）'] = df['薪資水準區間（月薪）'].fillna('Unknown')

#從哪裡得知
df['從哪裡得知講座消息？'].isnull().sum()
df['從哪裡得知講座消息？'].value_counts()
df['從哪裡得知講座消息？'] = df['從哪裡得知講座消息？'].fillna('Accupass 活動通')

#課程方案
df['課程方案'].isnull().sum()
df['課程方案'].value_counts()
df['課程方案'] = df['課程方案'].replace('白金升級VIP','VIP鑽石尊爵')

#手機號碼
df[df['手機號碼'].isnull()]
df['手機號碼'].isnull().sum()
df['手機號碼'] = df['手機號碼'].fillna('ya')
len(df['手機號碼'].unique())

#是否報退
df['是否報退'].notnull().sum()
df['是否報退'].value_counts()
df['確定會買課程'] =  np.where(df['是否報退']=="否", '有買', '沒買')
df[df['確定會買課程']=='有買'].count()

df.info()
df.isnull().sum()
################################################################
############ READ THE FINAL REVISED DATA & DATA EDA ############

"""
任務：
報名率與報退率，與其趨勢 done
會買的買哪一種課，與其趨勢done
用機器學習看誰會買
用機器學習看誰買哪一種課
"""
def buy_course_ratio():   
    """ 課程總購買率 """    
    a = '課程總購買率： ' + '{:.2f}'.format(len(df[df['確定會買課程']=='有買'])/len(df[df['是否到場']=='是'])*100) + '%'
    return a
buy_course_ratio()


def buy_course_situation():
    """ 報名率與報退率趨勢 """
    df_register_and_buy = df[df['是否到場']=='否']
    course_situation_list = ['Variation','薪資水準區間（月薪）', '從哪裡得知講座消息？',
                             '場次', '日期', '星期', '場地']
    
    for buy_course_situation in df_register_and_buy[course_situation_list]:
        print(df_register_and_buy.pivot_table(values = '手機號碼', 
                                             index=buy_course_situation, 
                                             columns='確定會買課程', 
                                             aggfunc='count',
                                             margins=True), '\n\n')
buy_course_situation()
#register_df.to_csv(path_or_buf='/Users/yc/Desktop/芬特克註冊資料(無職業)-v1r00.csv')

def course_type_situation():
    """ 報名率與報退率趨勢 """

    df_purchase_type = df[df['確定會買課程']=='有買']
    course_situation_list = ['Variation', '薪資水準區間（月薪）', '從哪裡得知講座消息？',
                             '場次', '日期', '星期', '場地']
    
    for buy_course_situation in df_purchase_type[course_situation_list]:
        print(df_purchase_type.pivot_table(values = '手機號碼', 
                                             index=buy_course_situation, 
                                             columns='課程方案', 
                                             aggfunc='count',
                                             margins=True), '\n\n')
course_type_situation()
   

############ COMBINE DATA ML for BUYING COURSE  ############
'''Import usual classification method'''
df.columns
df_buy_course = df[['Variation', '薪資水準區間（月薪）', 
                    '從哪裡得知講座消息？', '場次', '場地',
                    '星期', '確定會買課程']]

#Train_Test_Split
df_buy_course.info()
df_buy_course['確定會買課程'] = np.where(df_buy_course['確定會買課程']=="有買",1, 0)
Y = df_buy_course['確定會買課程']
Y = df_buy_course['確定會買課程'].values
X = df_buy_course.drop('確定會買課程',axis=1)

train=[]
for dummy in X.columns:
    train.append(pd.get_dummies(X[dummy]))
X = pd.concat([train[0],train[1],train[2],train[3],train[4],train[5]], axis=1)
X = X.values

#Z = pd.concat([X,Y],axis=1)# this is for feature_importances_Z(your ML method)


##############Split data into train & test
from sklearn.model_selection import train_test_split
X_train, X_test, Y_train, Y_test = train_test_split(X, Y, test_size=0.3) 

############## Feature Scaling
#from sklearn.preprocessing import StandardScaler
#scaler = StandardScaler()
#X_train[:,0:4] = scaler.fit_transform(X_train[:,0:4])
#X_test[:,0:4] = scaler.transform(X_test[:,0:4])
from sklearn import metrics
from sklearn.metrics import confusion_matrix

""" 基本準確率 """
'基本準確率高於： ' + '{:.2f}'.format((len(Y)-Y.sum())/len(Y)*100) + '%'

#1.
### KNN ###
from sklearn.neighbors import KNeighborsClassifier
knn = KNeighborsClassifier(n_neighbors=3) 
knn.fit(X_train, Y_train)
y_pred = knn.predict(X_test)
knn.score(X_train,Y_train)
print(metrics.accuracy_score(Y_test, y_pred))
#confusion_matrix
confusion_matrix(Y_test, y_pred)

#2.
### Logistoc Regression ###
from sklearn.linear_model import LogisticRegression
L_reg = LogisticRegression(C=1)
L_reg.fit(X_train, Y_train)
L_reg.score(X_test,Y_test)
y_lg_pred = L_reg.predict(X_test)
L_reg.score(X_train, Y_train)
print(metrics.accuracy_score(Y_test, y_lg_pred))
#confusion_matrix
confusion_matrix(Y_test, y_lg_pred)

#3.
### LinearSVM ###
from sklearn.svm import LinearSVC
lsvc = LinearSVC(C=1)
lsvc.fit(X_train, Y_train)
lsvc.score(X_train, Y_train)
lsvc_pred = lsvc.predict(X_test)
lsvc.score(X_test,Y_test)
# from sklearn import metrics
print(metrics.accuracy_score(Y_test, lsvc_pred))
#confusion_matrix
confusion_matrix(Y_test, y_lg_pred)

#4.
### SVM with kernel trick under rbf(Gaussian kernel)###
from sklearn.svm import SVC
svc = SVC(kernel='rbf', C=100, gamma='auto')
svc.fit(X_train, Y_train)
svc.score(X_train, Y_train)
svc_pred = svc.predict(X_test)
svc.score(X_test,Y_test)
print("The test score is: {}".format(metrics.accuracy_score(Y_test, svc_pred)))
confusion_matrix(Y_test, svc_pred)

#4.
### Decisiom Tree ###
from sklearn.tree import DecisionTreeClassifier
from IPython.display import Image  
from sklearn import tree
import pydotplus

tree = DecisionTreeClassifier(max_depth=None,max_features=None, criterion="gini")
tree.fit(X_train, Y_train)
tree.score(X_train, Y_train)
tree_pred = tree.predict(X_test)
print(metrics.accuracy_score(Y_test, tree_pred)) #tree.score(X_test,Y_test)
confusion_matrix(Y_test, tree_pred)


from sklearn import tree
tree = DecisionTreeClassifier(max_depth=5,max_features=None,criterion="gini")
#tree.fit(X_train, Y_train)
clf = tree.fit(X_train, Y_train)
import graphviz 
dot_data = tree.export_graphviz(clf, out_file=None) 
graph = graphviz.Source(dot_data) 
graph.render("iris") 
dot_data = tree.export_graphviz(clf, out_file=None, 
                         #feature_names=a,  
                         class_names='確定會買課程',  
                         filled=True, rounded=True,  
                         special_characters=True)  
graph = graphviz.Source(dot_data)  
graph 


#4.
### Decisiom Tree Advanced: Random Forest ###
from sklearn.ensemble import RandomForestClassifier
rdn_forest = RandomForestClassifier(n_estimators=100, max_depth=5, max_features='auto')
rdn_forest.fit(X_train,Y_train)
rdn_forest.score(X_train,Y_train)
#from sklean import metrics => metrcis.accuracy_score
rdn_forest.score(X_test, Y_test)
rdn_forest_pred = rdn_forest.predict(X_test)
confusion_matrix(Y_test, rdn_forest_pred)

#5.
### Gradient Boosting machines ###
from sklearn.ensemble import GradientBoostingClassifier
bgc = GradientBoostingClassifier(learning_rate=0.05, max_depth=5, n_estimators=200)
bgc.fit(X_train,Y_train)
bgc.score(X_train,Y_train)
bgc_pred = bgc.predict(X_test)
metrics.accuracy_score(Y_test,bgc_pred)
bgc.score(X_test,Y_test)


#6. 
### Naive Bayes ###
from sklearn.naive_bayes import GaussianNB
nb = GaussianNB(priors=None)
nb.fit(X_train,Y_train)
nb.score(X_train,Y_train)
nb.score(X_test,Y_test)
#or confusion_matrix
from sklearn import metrics
from sklearn.metrics import confusion_matrix
nb_pred = nb.predict(X_test)
confusion_matrix(Y_test, nb_pred)


knn = knn.score(X_test,Y_test)
logistic_reg = L_reg.score(X_test,Y_test)
linear_svm = lsvc.score(X_test,Y_test)
kernel_svm = svc.score(X_test,Y_test)
decision_tree = tree.score(X_test, Y_test)
random_forest = rdn_forest.score(X_test, Y_test)
gradient_boosting = bgc.score(X_test,Y_test)
naive_bayes = nb.score(X_test,Y_test)
    
def accuracy_plot():    
    accuracy_score = pd.Series([knn, logistic_reg, linear_svm,
                               kernel_svm, decision_tree, random_forest,
                               gradient_boosting ,naive_bayes])
    accuracy_score = round(accuracy_score*100, 2)
    
    ML_name = pd.Series(['knn', 'logistic_reg', 'linear_svm',
                               'kernel_svm', 'decision_tree', 'random_forest',
                               'gradient_boosting' ,'naive_bayes'])
    
    df_accuracy = pd.DataFrame({'準確率':accuracy_score, '方法':ML_name})
    return sns.factorplot(x="方法", y="準確率", data=df_accuracy,
                          kind="bar", palette="muted", aspect=1.5)
accuracy_plot()


def accuracy_table():    
    accuracy_score = pd.Series([knn, logistic_reg, linear_svm,
                               kernel_svm, decision_tree, random_forest,
                               gradient_boosting ,naive_bayes])
    accuracy_score = round(accuracy_score*100, 2)
    
    ML_name = pd.Series(['knn', 'logistic_reg', 'linear_svm',
                               'kernel_svm', 'decision_tree', 'random_forest',
                               'gradient_boosting' ,'naive_bayes'])
    
    df_accuracy = pd.DataFrame({'準確率':accuracy_score, '方法':ML_name})
    return df_accuracy

accuracy_table()

