import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import Rectangle

# ====================================================================
# 🚨🚨🚨 중요: 사용할 파일 및 경로 설정 🚨🚨🚨
# 엑셀 파일을 직접 읽어옵니다.
# ====================================================================
file_path = 'EST_TEST_ENDTIME2.xlsx'

# Sheet1의 실제 컬럼 이름 ('AVG_MIN', 'MED_MIN')을 사용합니다.
columns_to_plot = ['AVG_MIN', 'MED_MIN']
# 차트에 표시할 간결한 이름
legend_labels = ['AVG_MIN', 'MED_MIN']

# 한글 폰트 설정 (Mac 환경)
import matplotlib.font_manager as fm

# Mac에서 사용 가능한 한글 폰트 설정
available_fonts = [f.name for f in fm.fontManager.ttflist]
korean_fonts = ['AppleGothic', 'AppleSDGothicNeo', 'NanumGothic', 'Arial Unicode MS']

font_set = False
for font in korean_fonts:
    if font in available_fonts:
        plt.rcParams['font.family'] = font
        font_set = True
        print(f"한글 폰트 '{font}' 설정 완료")
        break

if not font_set:
    print("Warning: 한글 폰트를 찾을 수 없습니다. 기본 폰트를 사용합니다.")
    plt.rcParams['font.family'] = 'DejaVu Sans'

plt.rcParams['axes.unicode_minus'] = False  # 마이너스 폰트 깨짐 방지


# 1. 데이터 불러오기
try:
    # 엑셀 파일을 직접 읽어옵니다. 첫 번째 시트를 읽고, header=0으로 설정하여 첫 번째 행을 헤더로 사용
    # LOT_NUMBER도 함께 불러옵니다
    df = pd.read_excel(file_path, sheet_name=0, header=0, engine='openpyxl')
    print(f"엑셀 파일 '{file_path}'을 성공적으로 로드했습니다.")
except FileNotFoundError:
    print(f"오류: 파일 '{file_path}'을 현재 디렉토리에서 찾을 수 없습니다.")
    print("-> 해결 방법: 엑셀 파일을 END_TIME_CHART.py 파일과 같은 폴더에 넣어주세요.")
    exit()
except Exception as e:
    print(f"파일 로드 실패 (추가 정보: {e})")
    exit()


# 2. 데이터 전처리 및 NULL 값 제거
try:
    # 필요한 컬럼만 선택하고, 결측치(NULL/NaN)를 포함하는 행 제거
    # LOT_NUMBER도 함께 유지
    required_columns = columns_to_plot + ['LOT_NUMBER']
    data_with_lot = df[required_columns].dropna(subset=columns_to_plot)

    if data_with_lot.empty:
        print("선택된 컬럼에 유효한 데이터가 없습니다. 차트 생성을 중단합니다.")
        exit()

    # 데이터 형식 변환 (숫자가 아닌 값은 NaN으로 변환 후 다시 제거)
    for col in columns_to_plot:
        data_with_lot[col] = pd.to_numeric(data_with_lot[col], errors='coerce')

    # 변환 과정에서 숫자가 아니었던 행(NaN이 된 행)도 제거
    data_with_lot = data_with_lot.dropna(subset=columns_to_plot)

    # data_for_plot는 숫자 컬럼만 (기존 로직과 호환성 유지)
    data_for_plot = data_with_lot[columns_to_plot].copy()

    if data_for_plot.empty:
        print("데이터를 숫자형으로 변환한 후 유효한 데이터가 남아있지 않습니다. 차트 생성을 중단합니다.")
        exit()

except KeyError:
    print(f"오류: 데이터 파일에 '{columns_to_plot[0]}' 또는 '{columns_to_plot[1]}' 컬럼이 없습니다. 컬럼 이름을 확인해주세요.")
    exit()
except Exception as e:
    print(f"데이터 처리 중 예상치 못한 오류 발생: {e}")
    exit()


# 3. 데이터셋 재구성 (Long-form으로 변환)
# 두 컬럼을 시각화 라이브러리에 적합하도록 하나의 'value' 컬럼으로 합칩니다.
df_long = data_for_plot.melt(var_name='지표', value_name='값 (분)')

# 이상치 필터링을 위한 임계값 설정 (100분 = 약 1.7시간)
outlier_threshold = 100
df_long_filtered = df_long[df_long['값 (분)'] <= outlier_threshold].copy()

# 전체 데이터 통계
total_count = len(data_for_plot)
normal_count = len(data_for_plot[(data_for_plot['AVG_MIN'] <= outlier_threshold) &
                                   (data_for_plot['MED_MIN'] <= outlier_threshold)])
outlier_count = total_count - normal_count

# 정상 범위 데이터만 추출
normal_data = data_for_plot[(data_for_plot['AVG_MIN'] <= outlier_threshold) &
                             (data_for_plot['MED_MIN'] <= outlier_threshold)]

# 통계 계산
avg_mean = normal_data['AVG_MIN'].mean()
avg_median = normal_data['AVG_MIN'].median()
med_mean = normal_data['MED_MIN'].mean()
med_median = normal_data['MED_MIN'].median()
diff = normal_data['AVG_MIN'] - normal_data['MED_MIN']
avg_greater = (diff > 0).sum()
med_greater = (diff < 0).sum()


# 4. 이상치 탐지 (IQR 방식 - Q3 + 1.5×IQR)
def detect_outliers_iqr(data, column):
    """IQR 방식으로 이상치 탐지"""
    Q1 = data[column].quantile(0.25)
    Q3 = data[column].quantile(0.75)
    IQR = Q3 - Q1
    upper_bound = Q3 + 1.5 * IQR
    median = data[column].median()

    outliers = data[data[column] > upper_bound].copy()
    outliers['deviation_from_median'] = outliers[column] - median
    outliers['deviation_percent'] = (outliers['deviation_from_median'] / median * 100)

    return outliers, upper_bound, median

# AVG_MIN 이상치 분석
normal_data_with_lot = data_with_lot[(data_with_lot['AVG_MIN'] <= outlier_threshold) &
                                      (data_with_lot['MED_MIN'] <= outlier_threshold)]
avg_outliers, avg_upper_bound, avg_median_val = detect_outliers_iqr(normal_data_with_lot, 'AVG_MIN')
avg_outliers_sorted = avg_outliers.nlargest(5, 'AVG_MIN')[['LOT_NUMBER', 'AVG_MIN', 'deviation_from_median', 'deviation_percent']]

# MED_MIN 이상치 분석
med_outliers, med_upper_bound, med_median_val = detect_outliers_iqr(normal_data_with_lot, 'MED_MIN')
med_outliers_sorted = med_outliers.nlargest(5, 'MED_MIN')[['LOT_NUMBER', 'MED_MIN', 'deviation_from_median', 'deviation_percent']]

# 이상치 비율 계산
avg_outlier_ratio = len(avg_outliers) / len(normal_data_with_lot) * 100
med_outlier_ratio = len(med_outliers) / len(normal_data_with_lot) * 100


# 5. 정규분포도 차트 (KDE Plot) 생성 및 시각화
# 레이아웃: 상단(그래프 + 분석박스) + 하단(테이블 전체 너비)
fig = plt.figure(figsize=(18, 11))
gs = fig.add_gridspec(2, 2, height_ratios=[3.5, 1.2], width_ratios=[2.2, 1], hspace=0.3, wspace=0.15)

# 왼쪽 상단: 정규분포 그래프
ax = fig.add_subplot(gs[0, 0])

# KDE Plot 생성
sns.kdeplot(
    data=df_long_filtered,
    x='값 (분)',
    hue='지표',
    fill=True,
    common_norm=False,
    palette={'AVG_MIN': 'darkorange', 'MED_MIN': 'royalblue'},
    linewidth=2.5,
    ax=ax
)

# 차트 제목 및 레이블 설정
ax.set_title(f"AVG_MIN vs MED_MIN: 정규분포도 (0~{outlier_threshold}분)", fontsize=19, pad=20, weight='bold')
ax.set_xlabel("값 (분)", fontsize=15)
ax.set_ylabel("밀도", fontsize=15)
ax.grid(True, linestyle='--', alpha=0.6, axis='y')
ax.legend(title='지표', labels=legend_labels, loc='upper left', fontsize=13)

# 오른쪽 상단: 분석 정보 박스 (그래프 높이와 동일)
ax_analysis = fig.add_subplot(gs[0, 1])
ax_analysis.set_xlim(0, 1)
ax_analysis.set_ylim(0, 1)
ax_analysis.axis('off')

# 분석 정보 텍스트
analysis_text = f"""• 정상 데이터: {normal_count}개
  ({normal_count/total_count*100:.1f}%)
• 이상치 제외: {outlier_count}개
  ({outlier_count/total_count*100:.1f}%)

【AVG_MIN (주황색)】
• 평균: {avg_mean:.2f}분
• 중앙값: {avg_median:.2f}분

【MED_MIN (파란색)】
• 평균: {med_mean:.2f}분
• 중앙값: {med_median:.2f}분
"""

# 분석 박스를 꽉 채우기 위해 배경 사각형 먼저 그리기
rect = Rectangle((0, 0), 1, 1, transform=ax_analysis.transAxes,
                facecolor='#FF8C00', edgecolor='#CC7000', linewidth=3, alpha=0.85)
ax_analysis.add_patch(rect)

# 텍스트 박스 추가 (배경 없이 텍스트만)
ax_analysis.text(0.5, 0.5, analysis_text,
                 transform=ax_analysis.transAxes,
                 fontsize=10.5,
                 verticalalignment='center',
                 horizontalalignment='center',
                 linespacing=1.6,
                 color='white',
                 weight='normal')


# 6. 하단에 이상치 테이블 추가 (전체 너비 사용)
# 하단 영역을 2개로 나눠서 AVG_MIN과 MED_MIN 테이블 배치
gs_bottom = gs[1, :].subgridspec(2, 1, hspace=0.15)

# AVG_MIN 이상치 테이블
ax_table1 = fig.add_subplot(gs_bottom[0])
ax_table1.axis('tight')
ax_table1.axis('off')

if len(avg_outliers_sorted) > 0:
    table_data_avg = []
    for idx, row in avg_outliers_sorted.iterrows():
        table_data_avg.append([
            row['LOT_NUMBER'],
            f"{row['AVG_MIN']:.2f}분",
            f"+{row['deviation_from_median']:.2f}분"
        ])

    table1 = ax_table1.table(
        cellText=table_data_avg,
        colLabels=['LOT_NUMBER', 'AVG_MIN 값', '중앙값 대비 차이'],
        cellLoc='center',
        loc='center',
        colWidths=[0.35, 0.3, 0.35]
    )
    table1.auto_set_font_size(False)
    table1.set_fontsize(9)
    table1.scale(1, 1.5)

    # 헤더 스타일
    for i in range(3):
        table1[(0, i)].set_facecolor('#FF8C00')
        table1[(0, i)].set_text_props(weight='bold', color='white')


# MED_MIN 이상치 테이블
ax_table2 = fig.add_subplot(gs_bottom[1])
ax_table2.axis('tight')
ax_table2.axis('off')

if len(med_outliers_sorted) > 0:
    table_data_med = []
    for idx, row in med_outliers_sorted.iterrows():
        table_data_med.append([
            row['LOT_NUMBER'],
            f"{row['MED_MIN']:.2f}분",
            f"+{row['deviation_from_median']:.2f}분"
        ])

    table2 = ax_table2.table(
        cellText=table_data_med,
        colLabels=['LOT_NUMBER', 'MED_MIN 값', '중앙값 대비 차이'],
        cellLoc='center',
        loc='center',
        colWidths=[0.35, 0.3, 0.35]
    )
    table2.auto_set_font_size(False)
    table2.set_fontsize(9)
    table2.scale(1, 1.5)

    # 헤더 스타일
    for i in range(3):
        table2[(0, i)].set_facecolor('#4169E1')
        table2[(0, i)].set_text_props(weight='bold', color='white')


plt.tight_layout()

# 차트 표시
plt.show()

print("\n차트 생성이 완료되었습니다.")
print(f"사용된 'AVG_MIN' 데이터 개수: {data_for_plot[columns_to_plot[0]].count()}")
print(f"사용된 'MED_MIN' 데이터 개수: {data_for_plot[columns_to_plot[1]].count()}")
