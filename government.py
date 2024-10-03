import pandas as pd


# 엑셀 파일 로드
file_path = "./price_monthly_usd.xlsx"
df = pd.read_excel(file_path, sheet_name='Sheet1')

# 파산 파일 로드
file_path_quarterly = "./data/historical_country_united_states_indicator_bankruptcies_.csv"
quarterly_data = pd.read_csv(file_path_quarterly)

# 'Date' 열을 datetime 형식으로 변환 후 연도와 월만 남김
df['Date'] = pd.to_datetime(df['Date']).dt.to_period('M')

# 결과 출력
print(df.head())


# DateTime 열을 datetime 형식으로 변환
quarterly_data['DateTime'] = pd.to_datetime(quarterly_data['DateTime'])

# 연도와 분기 정보 추출
quarterly_data['Year'] = quarterly_data['DateTime'].dt.year
quarterly_data['Month'] = quarterly_data['DateTime'].dt.month

# 필요한 열만 추출
quarterly_data = quarterly_data[['Year', 'Month', 'Close']]

# 데이터 프레임을 연도별로 그룹화하고 각 연도마다 데이터를 월별로 확장
monthly_data_list = []
for year, group in quarterly_data.groupby('Year'):
    # 연도 내의 각 월을 1월부터 12월까지 채우기
    year_monthly_data = pd.DataFrame({'Year': [year] * 12, 'Month': range(1, 13)})

    # 현재 그룹의 월별 데이터를 year_monthly_data에 병합하여 누락된 월을 NaN으로 채우기
    year_monthly_data = year_monthly_data.merge(group, on=['Year', 'Month'], how='left')
    
    # 연도 내에 분기별 데이터가 하나만 있을 경우: 해당 값을 12개월로 나눔
    if year_monthly_data['Close'].notna().sum() == 1:
        single_quarter_value = year_monthly_data['Close'].dropna().values[0]
        year_monthly_data['Close'] = single_quarter_value / 12  # 12개월로 나눠서 모든 월에 동일 값 할당
    
    # 분기별 데이터가 3개일 경우: 각각 3개월씩 동일하게 배분
    elif year_monthly_data['Close'].notna().sum() == 3:
        for i in range(len(year_monthly_data)):
            if pd.notna(year_monthly_data.loc[i, 'Close']):
                # 현재 분기 값을 다음 세 달에 동일하게 나눔
                year_monthly_data.loc[i:i+2, 'Close'] = year_monthly_data.loc[i, 'Close'] / 3

    # 분기별 데이터가 4개 이상일 경우: 각 분기 값을 나머지 월에 채움
    else:
        year_monthly_data['Close'] = year_monthly_data['Close'].interpolate(method='linear', limit_direction='both')
    
    # 새로운 Date 컬럼 추가
    year_monthly_data['Date'] = pd.to_datetime(year_monthly_data[['Year', 'Month']].assign(Day=1))
    year_monthly_data['Date'] = year_monthly_data['Date'].dt.to_period('M')
    
    # Frequency 컬럼 추가
    year_monthly_data['Frequency'] = 'Monthly'
    
    # 필요한 컬럼만 추출하여 리스트에 추가
    monthly_data_list.append(year_monthly_data[['Date', 'Close', 'Frequency']])

# 모든 연도의 데이터를 병합하여 최종 월별 데이터프레임 생성
final_monthly_df = pd.concat(monthly_data_list, ignore_index=True)

# Frequency 컬럼 삭제
final_monthly_df = final_monthly_df.drop(columns=['Frequency'])

# 'Close' 컬럼을 'Bankruptcies'로 변경
final_monthly_df = final_monthly_df.rename(columns={'Close': 'Bankruptcies'})

# 결과 출력
print(final_monthly_df.head(50))


# 'Date' 컬럼을 기준으로 두 데이터프레임을 병합 (df를 기준으로 하고, 없는 날짜는 비워둠)
merged_df = pd.merge(df, final_monthly_df, on='Date', how='left')

# 'Date' 컬럼을 'YYYY-MM' 형식의 문자열로 변환
merged_df['Date'] = merged_df['Date'].dt.strftime('%Y-%m')

# 결과 확인
print(merged_df.head(50))

# 병합된 데이터프레임을 CSV로 저장
merged_df.to_csv("./merged_data.csv", index=False)

