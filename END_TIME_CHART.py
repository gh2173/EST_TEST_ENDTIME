import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import numpy as np

# ====================================================================
# 🚨🚨🚨 중요: 사용할 파일 및 경로 설정 🚨🚨🚨
# 엑셀 파일을 직접 읽어옵니다.
# ====================================================================
file_path = 'EST_TEST_ENDTIME2.xlsx' 

# Sheet1의 실제 컬럼 이름 ('AVG_MIN', 'MED_MIN')을 사용합니다.
columns_to_plot = ['AVG_MIN', 'MED_MIN']
# 차트에 표시할 간결한 이름 ㅅ 
legend_labels = ['AVG_MIN', 'MED_MIN']

# 한글 폰트 설정 (Mac 환경 기준)
try:
    plt.rcParams['font.family'] = 'AppleGothic'  # Mac
    plt.rcParams['axes.unicode_minus'] = False # 마이너스 폰트 깨짐 방지
except:
    print("Warning: 한글 폰트 설정을 찾을 수 없습니다. 그래프에 한글이 깨져 보일 수 있습니다.")


# 1. 데이터 불러오기
try:
    # 엑셀 파일을 직접 읽어옵니다. 첫 번째 시트를 읽고, header=0으로 설정하여 첫 번째 행을 헤더로 사용
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
    data_for_plot = df[columns_to_plot].dropna()
    
    if data_for_plot.empty:
        print("선택된 컬럼에 유효한 데이터가 없습니다. 차트 생성을 중단합니다.")
        exit()

    # 데이터 형식 변환 (숫자가 아닌 값은 NaN으로 변환 후 다시 제거)
    for col in columns_to_plot:
        data_for_plot[col] = pd.to_numeric(data_for_plot[col], errors='coerce')
    
    # 변환 과정에서 숫자가 아니었던 행(NaN이 된 행)도 제거
    data_for_plot = data_for_plot.dropna()

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


# 4. 정규분포도 차트 (KDE Plot) 생성 및 시각화
# 두 개의 차트를 생성: (1) 전체 데이터 (2) 정상 범위 확대
fig, axes = plt.subplots(1, 2, figsize=(18, 6))

# ============ 차트 1: 정상 범위 (0~100분) - 메인 차트 ============
sns.kdeplot(
    data=df_long_filtered,
    x='값 (분)',
    hue='지표',
    fill=True,
    common_norm=False,
    palette={'AVG_MIN': 'darkorange', 'MED_MIN': 'royalblue'},
    linewidth=2.5,
    ax=axes[0]
)

axes[0].set_title(f"AVG_MIN vs MED_MIN: 정규분포 (0~{outlier_threshold}분)", fontsize=14, pad=15)
axes[0].set_xlabel("값 (분)", fontsize=12)
axes[0].set_ylabel("밀도", fontsize=12)
axes[0].grid(True, linestyle='--', alpha=0.6, axis='y')
axes[0].legend(title='지표', labels=legend_labels, loc='upper right')
axes[0].text(0.02, 0.98, f'정상 데이터: {normal_count}개',
             transform=axes[0].transAxes, fontsize=10, verticalalignment='top',
             bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))

# ============ 차트 2: 전체 데이터 (로그 스케일) ============
sns.kdeplot(
    data=df_long,
    x='값 (분)',
    hue='지표',
    fill=True,
    common_norm=False,
    palette={'AVG_MIN': 'darkorange', 'MED_MIN': 'royalblue'},
    linewidth=2.5,
    ax=axes[1],
    log_scale=True  # 로그 스케일 적용
)

axes[1].set_title("AVG_MIN vs MED_MIN: 전체 데이터 분포 (로그 스케일)", fontsize=14, pad=15)
axes[1].set_xlabel("값 (분, 로그 스케일)", fontsize=12)
axes[1].set_ylabel("밀도", fontsize=12)
axes[1].grid(True, linestyle='--', alpha=0.6, axis='y')
axes[1].legend(title='지표', labels=legend_labels, loc='upper right')
axes[1].text(0.02, 0.98, f'전체 데이터: {total_count}개\n이상치: {outlier_count}개',
             transform=axes[1].transAxes, fontsize=10, verticalalignment='top',
             bbox=dict(boxstyle='round', facecolor='lightcoral', alpha=0.5))

plt.tight_layout()

# 차트 표시
plt.show()

print("\n차트 생성이 완료되었습니다.")
print(f"사용된 'AVG_MIN' 데이터 개수: {data_for_plot[columns_to_plot[0]].count()}")
print(f"사용된 'MED_MIN' 데이터 개수: {data_for_plot[columns_to_plot[1]].count()}")