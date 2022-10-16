import os,sys
import pandas as pd
file_path = os.path.join(os.getcwd(),"영화")
date_path = os.path.join(os.getcwd(),"영화 날짜")

count=0
file_list = os.listdir(file_path)

# 데이터 프레임의 리스트생성
initialize = False
total_result = {
    "이름" : []
}
# 계산 종류 람다함수 사용
max_list = ["스크린수","누적관객수"]
mean_list = ["상영횟수","상영점유율","좌석수","관객수"]
name_cal = lambda x: total_result["이름"].append(x)# 데이터 리스트에 이름 추가
max_cal = lambda x: total_result[x].append(movie_list[x].max()) # 데이터 리스트에 최대값으로 계산하여 추가
mean_cal = lambda x: total_result[x].append(movie_list[x].mean()) # 데이터 리스트에 평균으로 계산하여 추가

for file_name_raw in file_list:
    print("현재 진행중인 파일 : "+file_name_raw)

    #  엑셀 확장자를 제거하고 이름으로 사용하기 위함
    Remove_extension = file_name_raw.replace('.xlsx',"")

    file_name= os.path.join(file_path,file_name_raw)# 영화 모든 파일 이름으로 가져오기
    file_date= os.path.join(date_path,"영화 날짜 파일.xlsx")
    df_file_date = pd.read_excel(file_date,engine='openpyxl')
    df = pd.read_excel(file_name,engine='openpyxl')
    date = df_file_date.iloc[count,1]
    movie_list = df[df["날짜"]>=date].iloc[0:10,:] #모든 열 가져오기 단 날짜가 엑셀상의 영화 개봉일 부터 10일

    # 초기화 구문 칼럼명들을 전부 추가해주고 필요없는 날짜는 제거함 이 구문은 한번만 사용되면 되나
    # 코드에서  movie_date_list를 for문상에서 읽어오기 때문에 for문상 한번만 돌수 있게 구현
    # 초기화 변수를 for문 밖에서 false로 두고 한번 수행하면 true로 바뀌게 하여  더이상 실행  되지 않음
    # 리스트에 날짜 칼럼 제외 모든 칼럼 추가
    if not initialize:
        initialize = True
        columns = list(movie_list.columns.values)
        columns.remove("날짜")
        for i in columns:
            total_result[i] = list()



    name_cal(Remove_extension)
    for x in max_list: max_cal(x)
    for x in mean_list: mean_cal(x)
    movie_list.set_index('날짜',inplace=True)
    count += 1
total_data = pd.DataFrame(total_result)
total_data.to_excel("Result.xlsx")