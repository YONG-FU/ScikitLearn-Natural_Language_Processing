from sklearn.datasets import load_files
from sklearn.feature_extraction.text import CountVectorizer, TfidfTransformer
from sklearn.naive_bayes import MultinomialNB
import openpyxl
import re


# *** Process 01 - Input Provided Documents ***

# 读入需要分析预测的文件
# file_name = "Abengoa - Facility Agreement"
# file_name = "Apcoa - Senior Facilities Agreement"
# file_name = "Basin - Note Purchase Agreement"
# file_name = "Cory - Junior Facility Agreement"
# file_name = "Noble - Financing Agreement"
file_name = "Northwest - Senior Loan Agreement"
# file_name = "Trumbull - Note Purchase Agreement"

file_type = "Loan Agreement"
fo = open("input-documents\\txt\\" + file_name + ".txt", encoding="utf8")

# *** Process 02 - Extract Keywords Related Contents ***

maturity_date_pattern = "Maturity Date+|Termination Date+|Repayment Date+|Scheduled Termination+|Final Repyament Date+"
maturity_date_list = []


# *** Process 03 - Classify Extracted Contents ***

# 选取参与分析的文本类别
datasets_categories = ['target-maturity-date', 'non-target-maturity-date']

# 从硬盘获取训练数据
datasets_train=load_files('datasets-train',
    categories=datasets_categories,
    load_content = True,
    encoding='latin1',
    decode_error='strict',
    shuffle=True, random_state=42)

# 统计词语出现次数
count_vectorizer = CountVectorizer()
X_train_counts = count_vectorizer.fit_transform(datasets_train.data)

# 使用tf-idf方法提取文本特征
tfidf_transformer = TfidfTransformer()
X_train_tfidf = tfidf_transformer.fit_transform(X_train_counts)

# 使用多项式朴素贝叶斯方法训练分类器
clf = MultinomialNB().fit(X_train_tfidf, datasets_train.target)

for line in fo.readlines():
    if len(re.findall(maturity_date_pattern, line)) > 0:
        # 预测用的新字符串
        docs_new = [line]

        # 字符串向量化处理
        X_new_counts = count_vectorizer.transform(docs_new)
        X_new_tfidf = tfidf_transformer.transform(X_new_counts)

        # 进行机器学习预测
        predicted = clf.predict(X_new_tfidf)

        if predicted[0] == 1:
            maturity_date_list.append(line)


# *** Process 04 - Understand Specific Classified Contents ***
date_value_pattern = "\d{1,2} January \d\d\d\d|" \
                     "\d{1,2} February \d\d\d\d|" \
                     "\d{1,2} March \d\d\d\d|" \
                     "\d{1,2} April \d\d\d\d|" \
                     "\d{1,2} May \d\d\d\d|" \
                     "\d{1,2} June \d\d\d\d|" \
                     "\d{1,2} July \d\d\d\d|" \
                     "\d{1,2} August \d\d\d\d|" \
                     "\d{1,2} September \d\d\d\d|" \
                     "\d{1,2} October \d\d\d\d|" \
                     "\d{1,2} November \d\d\d\d|" \
                     "\d{1,2} December \d\d\d\d|" \
                     "January \d{1,2}, \d\d\d\d|" \
                     "February \d{1,2}, \d\d\d\d|" \
                     "March \d{1,2}, \d\d\d\d|" \
                     "April \d{1,2}, \d\d\d\d|" \
                     "May \d{1,2}, \d\d\d\d|" \
                     "June \d{1,2}, \d\d\d\d|" \
                     "July \d{1,2}, \d\d\d\d|" \
                     "August \d{1,2}, \d\d\d\d|" \
                     "September \d{1,2}, \d\d\d\d|" \
                     "October \d{1,2}, \d\d\d\d|" \
                     "November \d{1,2}, \d\d\d\d|" \
                     "December \d{1,2}, \d\d\d\d"


# *** Process 05 - Output Required Information Results ***
wb = openpyxl.Workbook()
sheet = wb.get_sheet_by_name('Sheet')
sheet["A1"] = "File Name"
sheet["A2"] = file_name
sheet["B1"] = "File Type"
sheet["B2"] = file_type
sheet["C1"] = "Machine Learning Training Model"
sheet["C2"] = str(clf)
sheet["D1"] = "Machine Learning Training Datasets"
sheet["D2"] = str(datasets_train)
sheet["E1"] = "Machine Learning Training Datasets Categories"
sheet["E2"] = str(datasets_categories)
sheet["F1"] = "Feature Extractor"
sheet["F2"] = str(tfidf_transformer)
sheet["G1"] = "Matrix of Feature Extractor"
sheet["G2"] = str(X_train_tfidf)
sheet["H1"] = "Count Vectorizer"
sheet["H2"] = str(count_vectorizer)
sheet["I1"] = "Matrix of Count Vectorizer"
sheet["I2"] = str(X_train_counts)
sheet["J1"] = "Pattern of Maturity Date"
sheet["J2"] = str(maturity_date_pattern)
sheet["K1"] = "Pattern of Date Value"
sheet["K2"] = str(date_value_pattern)
sheet["M1"] = "Maturity Date Text"
sheet["N1"] = "Maturity Date Value"
possible_maturity_date_number = 1 + 1

for maturity_date in maturity_date_list:
    results = re.findall(date_value_pattern, maturity_date)
    if len(results) > 0:
        # Maturity Date Text
        sheet["M" + str(possible_maturity_date_number)] = maturity_date
        print(maturity_date)
        # Maturity Date Value
        cell_value = ""
        if len(results) == 1:
            cell_value = results[0]
        else:
            for result in results:
                cell_value += result + "\n"
        sheet["N" + str(possible_maturity_date_number)] = cell_value
        print(cell_value)
        # Update the number of Maturity Date
        possible_maturity_date_number += 1

wb.save("Maturity Date - " + file_name + ".xlsx")