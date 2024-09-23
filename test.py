import pandas as pd
import matplotlib.pyplot as plt
import platform
from matplotlib import font_manager, rc

# 한글 폰트 설정
path = "C:/Windows/Fonts/malgun.ttf"  # Windows에서 맑은 고딕 폰트 경로
if platform.system() == "Darwin":  # macOS
    rc("font", family="AppleGothic")
elif platform.system() == "Windows":  # Windows
    font_name = font_manager.FontProperties(fname=path).get_name()
    rc("font", family=font_name)
else:
    print("Unknown platform - default font used")

# 음수 기호 깨짐 방지
plt.rcParams["axes.unicode_minus"] = False

# 엑셀 파일 경로
file_path = './data/Changes_latest_as_of_Sep2024_IFS.xlsx'

# 1. Monthly 시트 데이터 처리
monthly_df = pd.read_excel(file_path, sheet_name='Monthly')

# 2. 분석할 국가 설정 (China, P.R.: Mainland는 China로 대체)
selected_countries = ['United States', 'United Kingdom', 'Russia', 'China', 'India', 'France', 'Germany']

# 3. 해당 국가들의 데이터 필터링 및 불필요한 열 삭제
monthly_df['Unnamed: 0'] = monthly_df['Unnamed: 0'].replace('China, P.R.: Mainland', 'China')
monthly_selected_df = monthly_df[monthly_df['Unnamed: 0'].isin(selected_countries)]

# 날짜 관련 열만 추출 (숫자형 또는 날짜형 열만 필터링)
date_columns = monthly_selected_df.columns[monthly_selected_df.columns.str.contains(r'\d{4}', na=False)]

# 'Unnamed: 0'을 '국가'로 열 이름 변경 후, 'China, P.R.: Mainland'를 'China'로 변경
monthly_selected_df = monthly_selected_df.rename(columns={'Unnamed: 0': '국가'}).set_index('국가')

# 4. 월별 변화량만 추출하고 null 값을 0으로 대체
monthly_changes_df = monthly_selected_df[date_columns].apply(pd.to_numeric, errors='coerce').fillna(0)

# 5. 국가별 금 보유 트렌드 분석 (누적 변화량)
monthly_cumsum_df = monthly_changes_df.cumsum(axis=1)

# 6. 전체 국가 평균 추가 (숫자형 데이터만 대상으로 함)
monthly_avg_df = monthly_df[date_columns].apply(pd.to_numeric, errors='coerce').fillna(0).mean(axis=0)
monthly_cumsum_df.loc['전체 국가 평균'] = monthly_avg_df.cumsum()

# 월별 데이터를 연도별로 집계
monthly_cumsum_df_transposed = monthly_cumsum_df.T
monthly_cumsum_df_transposed.index = pd.to_datetime(monthly_cumsum_df_transposed.index, format='%Y.%m')

# 연도 단위로 월 데이터를 합산하여 연단위 데이터 생성
yearly_from_monthly = monthly_cumsum_df_transposed.resample('Y').mean()

# 누적 변화량 시각화 (전체 국가 평균 포함, 검정색 점선)
plt.figure(figsize=(10, 6))
ax = monthly_cumsum_df.T.plot()
ax.lines[-1].set_linestyle("--")  # 전체 국가 평균을 점선으로 설정
ax.lines[-1].set_color("black")   # 전체 국가 평균을 검정색으로 설정

plt.title('금 보유량 누적 변화 추이')
plt.xlabel('월')
plt.ylabel('누적 변화량 (톤)')
plt.legend(title='국가', bbox_to_anchor=(1.05, 1), loc='upper left')  # 범례를 국가명으로 설정
plt.tight_layout()
plt.show()

### 2. Annual 시트 데이터 처리
annual_df = pd.read_excel(file_path, sheet_name='Annual')

# Comments 열 삭제 및 'Country' 사용 (Annual 시트에서만 Comments가 존재함)
annual_df = annual_df.drop(columns=['Comments'])  # Comments 열 삭제
annual_df['Country'] = annual_df['Country'].replace('China, P.R.: Mainland', 'China')

# 시트에 있는 국가만 필터링
available_countries = [country for country in selected_countries if country in annual_df['Country'].values]

# 날짜에 해당하는 연도 컬럼만 선택
years = [col for col in annual_df.columns if isinstance(col, int)]

# 필터링된 국가들로 데이터 필터링
yearly_changes_df = annual_df[annual_df['Country'].isin(available_countries)].set_index('Country')[years]

# 7. 전체 국가 평균 연도별 금 보유량 변화 추가
yearly_avg_df = annual_df[years].apply(pd.to_numeric, errors='coerce').fillna(0).mean(axis=0)
yearly_changes_df.loc['전체 국가 평균'] = yearly_avg_df

# 월 데이터를 집계한 연단위 데이터를 연 시트와 비교
yearly_from_monthly.loc['전체 국가 평균'] = yearly_avg_df  # 평균값을 연 시트에 반영

# 연도별 변화량 시각화 (전체 국가 평균 포함, 검정색 점선)
plt.figure(figsize=(10, 6))
ax = yearly_changes_df.T.plot()
ax.lines[-1].set_linestyle("--")  # 전체 국가 평균을 점선으로 설정
ax.lines[-1].set_color("black")   # 전체 국가 평균을 검정색으로 설정

plt.title('국가별 연도별 금 보유량 변화')
plt.xlabel('연도')
plt.ylabel('금 보유량 변화 (톤)')
plt.legend(title='국가', bbox_to_anchor=(1.05, 1), loc='upper left')
plt.tight_layout()
plt.show()

# 8. 국가 간 금 보유 전략 비교 (최종 누적 변화량 비교)
final_cumulative_changes = monthly_cumsum_df.iloc[:, -1]

# 최종 누적 변화량 시각화 (양수/음수 색상 다르게 적용)
colors = ['red' if value < 0 else 'blue' for value in final_cumulative_changes]

plt.figure(figsize=(10, 6))
final_cumulative_changes.plot(kind='bar', color=colors)
plt.title('국가별 최종 누적 금 보유량 변화')
plt.ylabel('누적 변화량 (톤)')
plt.show()
