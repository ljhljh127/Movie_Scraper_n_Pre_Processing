import os,sys
import pandas as pd
from openpyxl import load_workbook
path="C:/Users/이정현/PycharmProjects/pythonProject/dataprocessing/영화"
count=0
file_list = os.listdir(path)
initialize = False
data_dict = {
    "이름" : []
}
max_list = ["스크린수","누적관객수"]
mean_list = ["상영횟수","상영점유율","좌석수","관객수"]
name_cal = lambda x: data_dict["이름"].append(x)
max_cal = lambda x: data_dict[x].append(movie_date_list[x].max())
mean_cal = lambda x: data_dict[x].append(movie_date_list[x].mean())

for file_name_raw in file_list:
    for file_name_raw in file_list:
        # Get file name
        file_real_name = file_name_raw.replace('.xlsx', "")
        print(f"Processing : {file_real_name}")
    file_name="C:/Users/이정현/PycharmProjects/pythonProject/dataprocessing/영화/"+file_name_raw
    file_date="C:/Users/이정현/PycharmProjects/pythonProject/dataprocessing/영화 날짜/영화 날짜 파일.xlsx"
    df_file_date = pd.read_excel(file_date,engine='openpyxl')
    df = pd.read_excel(file_name,engine='openpyxl')
    date = df_file_date.iloc[count,1]
    print(file_name_raw)
    movie_date_list = df[df["날짜"]>=date].iloc[0:10,:]
    if not initialize:
        initialize = True
        columns = list(movie_date_list.columns.values)
        columns.remove("날짜")
        for i in columns:
            data_dict[i] = list()

    name_cal(file_real_name)
    for x in max_list: max_cal(x)
    for x in mean_list: mean_cal(x)
    movie_date_list.set_index('날짜',inplace=True)
    count += 1
    total_data = pd.DataFrame(data_dict)
    total_data.to_excel("Result.xlsx")

