from flask import Flask, render_template, request, send_file # type: ignore
import pandas as pd # type: ignore
import os
import re
import warnings
import numpy as np
warnings.filterwarnings('ignore')

app = Flask(__name__)

# 업로드 폴더 설정
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/')
def index():
    return render_template('index.html')

"""
=============== 관내여비 담당자 서포터 기능 ===============
"""
@app.route('/trip/upload', methods=['POST'])
def upload_and_process_trip_files():
    if 'trip_file' not in request.files or 'tag_file' not in request.files:
        return 'No file part'
    
    trip_file = request.files['trip_file']
    tag_file = request.files['tag_file']
    
    if trip_file.filename == '' or tag_file.filename == '':
        return 'No selected file'
    
    # 경로 설정 및 저장
    trip_path = os.path.join(app.config['UPLOAD_FOLDER'], 'trip_all.xlsx')
    tag_path = os.path.join(app.config['UPLOAD_FOLDER'], 'tag_all.xlsx')
    trip_file.save(trip_path)
    tag_file.save(tag_path)
    
    try:
        # 1. 데이터 로드 및 전처리
        df_trip = pd.read_excel(trip_path, header=1)
        col_tmp = pd.read_excel(trip_path, nrows=0).columns
        col_tmp = col_tmp[:col_tmp.get_loc('출장기간') - 1]
        df_trip.columns.values[:len(col_tmp)+1] = pd.read_excel(trip_path, nrows=0).columns[:8]

        # 필터링
        df_trip = df_trip[
            (df_trip['근태항목'] == '관내출장') &
            (df_trip['결재상태'].str.startswith('결재완료'))
        ]

        # 부서 검증
        bepa = ['경영기획실', '청년사업단', '산업인력지원단', '소상공인지원단', '기업지원단', '글로벌사업추진단', '부원장', '기업옴부즈맨실', '임원']
        if not all(dept in bepa for dept in df_trip['부서'].unique()):
            raise ValueError("오류가 발생하였습니다. 개발팀 연락 바랍니다.")

        # 필요한 컬럼 추출
        cols_needed = ['부서', '사원코드', '사원', '직급', '신청일', '시작일', '종료일', '시작시간', '종료시간',
                       '일수', '신청시간', '교통수단', '운전자', '출발지', '도착지', '경유지', '방문처', '목적', '내용']
        df_trip = df_trip[cols_needed].copy()
        
        # 2. 태그 데이터 처리
        df_tag = pd.read_excel(tag_path)
        df_tag = df_tag[['태깅일자', '사원코드', '근태구분', '근무시간']]
        
        tags_out = df_tag[df_tag['근태구분'] == '외출'].sort_values('근무시간').groupby(['태깅일자', '사원코드'])['근무시간'].last().reset_index(name='외출태그')
        tags_in = df_tag[df_tag['근태구분'] == '복귀'].sort_values('근무시간').groupby(['태깅일자', '사원코드'])['근무시간'].first().reset_index(name='복귀태그')

        df_trip = pd.merge(df_trip, tags_out, how='left', left_on=['시작일', '사원코드'], right_on=['태깅일자', '사원코드'])
        df_trip = pd.merge(df_trip, tags_in, how='left', left_on=['시작일', '사원코드'], right_on=['태깅일자', '사원코드'])

        # 3. 로직 적용 함수
        def apply_logic(row):
            str_time = row['시작시간']
            end_time = row['종료시간']
            out_time = row['외출태그']
            in_time = row['복귀태그']
            
            # 기본값: 태그가 있으면 일단 가져옴 (시간 포맷팅)
            if pd.notna(out_time): 
                out_time = str(out_time)[:5]
                out_use = out_time
            else:
                out_use = None
                
            if pd.notna(in_time): 
                in_time = str(in_time)[:5]
                in_use = in_time
            else:
                in_use = None

            # [수정] 로직 1: 신청시간과 태그시간 불일치 -> None (빈칸) 반환
            if pd.notna(out_time) and pd.notna(in_time):
                if (out_time > end_time) or (in_time < str_time):
                    return pd.Series([out_time, in_time, None, None]) 

            # 로직 2: 자동 설정 (없는 경우 인정)
            if (str_time <= '09:00') and pd.isna(out_time):
                out_use = str_time
            elif pd.notna(out_time) and (str_time > out_time):
                out_use = str_time
            
            if (end_time >= '18:00') and pd.isna(in_time):
                in_use = end_time
            elif pd.notna(in_time) and (end_time < in_time):
                in_use = end_time

            return pd.Series([out_time, in_time, out_use, in_use])

        # 로직 적용
        df_trip[['외출태그', '복귀태그', '외출태그(인정)', '복귀태그(인정)']] = df_trip.apply(apply_logic, axis=1)
        
        # [삭제됨] fillna 코드를 삭제했습니다. 
        # 이유: 불일치라서 None이 된 경우, fillna가 있으면 다시 원래 태그값으로 채워져서 돈이 지급되어 버립니다.
        
        # 4. 시간 및 여비 계산
        # [수정] calc_out 대신 '외출태그(인정)'을 바로 사용합니다.
        out_dt = pd.to_datetime(df_trip['외출태그(인정)'], format='%H:%M', errors='coerce')
        in_dt = pd.to_datetime(df_trip['복귀태그(인정)'], format='%H:%M', errors='coerce')
        
        diff = in_dt - out_dt
        df_trip['출장시간(산출)/분'] = diff.dt.total_seconds() // 60
        
        def format_duration(x):
            if pd.isna(x): return None
            total_seconds = int(x.total_seconds())
            hours, remainder = divmod(total_seconds, 3600)
            minutes = remainder // 60
            return f'{hours}:{minutes:02d}'
        
        df_trip['출장시간'] = diff.apply(format_duration)

        # 여비 계산
        conditions = [
            df_trip['출장시간(산출)/분'].isna(), # 인정시간이 None이면 0원
            df_trip['출장시간(산출)/분'] < 240
        ]
        choices = [0, 10000]
        df_trip['여비'] = np.select(conditions, choices, default=20000)
        
        df_trip.loc[df_trip['교통수단'] == '관용차량', '여비'] -= 10000
        df_trip['여비'] = df_trip['여비'].clip(lower=0) 

        # 불필요 컬럼 제거 (calc_out, calc_in은 애초에 안 만들었으니 제거 리스트에서도 뺌)
        df_trip.drop(columns=['태깅일자_x', '사원코드_x', '태깅일자_y', '사원코드_y'], inplace=True, errors='ignore')

    except Exception as e:
        return f"파일 처리 중 오류 발생: {str(e)}"
    
    # 5. 부서별 저장 및 서식 적용
    department_files = []
    final_cols = ['부서', '사원', '직급', '신청일', '시작일', '종료일', '시작시간', 
                  '종료시간', '일수', '신청시간', '외출태그', '복귀태그', '외출태그(인정)', '복귀태그(인정)',
                  '출장시간', '여비', '교통수단', '운전자', '출발지', '도착지', '경유지', '방문처', '목적', '내용']

    for dept, group in df_trip.groupby('부서'):
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{dept}_관내여비.xlsx')
        group = group.sort_values(by=['사원', '시작일'] if '시작일' in group.columns else ['사원'])
        group = group[final_cols]
        
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            group.to_excel(writer, index=False, sheet_name='Sheet1')
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            max_row = len(group) + 1
            
            # 서식
            red_format = workbook.add_format({'font_color': 'red', 'bold': True})
            gray_bg_format = workbook.add_format({'bg_color': '#D3D3D3'})
            
            # 서식 적용
            worksheet.conditional_format(f'T2:T{max_row}', {'type': 'cell', 'criteria': 'equal to', 'value': '"시청"', 'format': red_format})
            worksheet.conditional_format(f'K2:K{max_row}', {'type': 'blanks', 'format': gray_bg_format})
            worksheet.conditional_format(f'L2:L{max_row}', {'type': 'blanks', 'format': gray_bg_format})
            
            worksheet.set_column('A:X', 12)

        department_files.append(file_path)

    return render_template('trip_result.html', department_files=department_files)

"""
=============== 숫자 한글 변환기 ===============
"""
def number_to_korean(num):
    num = int(re.sub(r'[,]', '', num))

    units = ['', '만', '억', '조', '경']
    small_units = ['', '십', '백', '천']  
    digits = [''] + list('일이삼사오육칠팔구')  

    if num == 0:
        return '영'
    
    result = []
    unit_index = 0
    
    while num > 0:
        part = num % 10000  
        num //= 10000
        
        if part > 0:
            part_str = ''
            for i in range(4):  
                digit = (part // (10 ** i)) % 10
                if digit != 0:
                    part_str = digits[digit] + small_units[i] + part_str
            
            result.append(part_str + units[unit_index])
        
        unit_index += 1
    
    return ''.join(result[::-1])

@app.route('/money', methods=['GET', 'POST'])
def money_converter():
    input_value = ""
    converted_value = ""
    
    if request.method == 'POST':
        num = request.form.get('number', '')
        try:
            num = re.sub(r'[^0-9]', '', num)  # 숫자만 남기기
            if num:
                input_value = "{:,}".format(int(num))  # 콤마 추가된 입력값
                converted_value = number_to_korean(num)  # 변환값
            else:
                converted_value = "올바른 숫자를 입력하세요."
        except ValueError:
            converted_value = "올바른 숫자를 입력하세요."
    
    return render_template('money.html', input_value=input_value, converted_value=converted_value)

if __name__ == '__main__':
    app.run(debug=True)

"""
=============== 관내여비 담당자 서포터 기능 ===============
"""

@app.route('/system', methods=['GET', 'POST'])
def system_index():
    if request.method == 'POST':
        # 신규 직원 정보 처리 로직 추가 예정
        return '신규 직원 계정 생성 기능은 곧 제공될 예정입니다.'
    return render_template('system_index.html')