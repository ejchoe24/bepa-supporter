from flask import Flask, render_template, request, send_file # type: ignore
import pandas as pd # type: ignore
import os
import re
import warnings
warnings.filterwarnings('ignore')

app = Flask(__name__)

# 업로드 및 처리 폴더 설정
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER

@app.route('/')
def index():
    return render_template('index.html')

"""
=============== 교육 담당자 서포터 기능 ===============
"""
@app.route('/edu')
def edu_index():
    return render_template('edu_index.html')

@app.route('/edu/upload', methods=['POST'])
def upload_edu_file():
    if 'file' not in request.files:
        return "파일이 없습니다."
    
    file = request.files['file']
    if file.filename == '':
        return "파일 이름이 없습니다."
    
    # CSV 파일 저장
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)
    
    # CSV 파일 처리
    try:
        df = pd.read_csv(file_path, encoding='CP949', skiprows=1)
    except Exception as e:
        return f"CSV 파일을 처리하는 중 오류가 발생했습니다: {e}"
    
    df = df.dropna(subset=['연번', '이름']).drop(columns=['비고1', '비고2', 'Unnamed: 14'])
    df = df.rename(columns={'교육\n일시': '교육일시', '교육\n시간': '교육시간', '구분2\n(법정의무/자율)' : '구분2(법정의무/자율)'})
    #df[['연번', '교육시간']] = df[['연번', '교육시간']].astype('int')
    
    # 체크박스 선택 여부 확인
    include_date = request.form.get('include_date') == 'yes'

    # 중복된 이름이 있는 데이터 추출 (체크박스 선택 시 '교육일시' 포함)
    subset_columns = ['이름', '구분1(외부/내부)', '구분2(법정의무/자율)', '법정 과정', '과정명']
    if include_date:
        subset_columns.append('교육일시')
 
    duplicated_names = df[df.duplicated(subset=subset_columns, keep=False)]
    duplicated_names = duplicated_names.sort_values(by=['이름', '과정명'])

    # 이름 개수 불일치 확인
    name_counts = df.groupby('과정명')['이름'].agg(고유개수='nunique', 이름개수='count').reset_index()
    name_mismatch = name_counts[name_counts['이름개수'] != name_counts['고유개수']]

    # 공백과 괄호 제거 전 후 비교
    space_yes = df.groupby('과정명').size().reset_index(name='띄어쓰기 제거 전')
    space_yes['과정명'] = space_yes['과정명'].str.replace(' ', '', regex=False)

    bracket_yes = space_yes.groupby('과정명').sum().reset_index()
    bracket_yes.columns = ['과정명', '괄호 제거 전']
    bracket_yes['과정명'] = bracket_yes['과정명'].apply(lambda x:re.sub(r'\(\d+(차시|시간)\)', '', x))

    space_no = df.copy()
    space_no['과정명'] = space_no['과정명'].str.replace(' ', '', regex=False)
    space_no = space_no.groupby('과정명').size().reset_index(name='띄어쓰기 제거 후')

    bracket_no = space_no.copy()
    bracket_no['과정명'] = bracket_no['과정명'].apply(lambda x: re.sub(r'\(\d+(차시|시간)\)', '', x))
    bracket_no = bracket_no.groupby('과정명').sum().reset_index()
    bracket_no.columns = ['과정명', '괄호 제거 후']

    space_df = pd.merge(space_yes, space_no, on='과정명', how='outer')
    space_df['일치 여부'] = space_df['띄어쓰기 제거 전'] == space_df['띄어쓰기 제거 후']
    space_df = space_df[space_df['일치 여부'] == False]
    space_df = space_df.groupby('과정명').agg({
        "띄어쓰기 제거 전" : lambda x: ", ".join(map(str, x)),
        "띄어쓰기 제거 후" : "first"
    }).reset_index()
    space_df.columns = ['과정명', '구분 별 개수', '전체 개수']

    bracket_df = pd.merge(bracket_yes, bracket_no, on='과정명', how='outer')
    bracket_df['일치 여부'] = bracket_df['괄호 제거 전'] == bracket_df['괄호 제거 후']
    bracket_df = bracket_df[bracket_df['일치 여부'] == False]
    bracket_df = bracket_df.groupby('과정명').agg({
        "괄호 제거 전" : lambda x: ", ".join(map(str, x)),
        "괄호 제거 후" : "first"
    }).reset_index()
    bracket_df.columns = ['과정명', '구분 별 개수', '전체 개수']

    # 파일 저장 경로 설정
    duplicated_file = os.path.join(app.config['PROCESSED_FOLDER'], 'duplicated_names.csv')
    mismatch_file = os.path.join(app.config['PROCESSED_FOLDER'], 'name_mismatch.csv')
    space_comparison_file = os.path.join(app.config['PROCESSED_FOLDER'], 'space_comparison.csv')
    bracket_comparison_file = os.path.join(app.config['PROCESSED_FOLDER'], 'bracket_comparison.csv')
    
    # 처리된 데이터 저장
    duplicated_names.to_csv(duplicated_file, index=False, encoding='CP949')
    name_mismatch[['과정명', '고유개수', '이름개수']].to_csv(mismatch_file, index=False, encoding='CP949')
    space_df.to_csv(space_comparison_file, index=False, encoding='CP949')
    bracket_df.to_csv(bracket_comparison_file, index=False, encoding='CP949')

    # 처리된 데이터프레임을 HTML로 변환하여 보여주기
    return render_template('edu_result.html', 
                           duplicated_names=duplicated_names.to_html(index=False, escape=False),
                           name_mismatch=name_mismatch[['과정명', '고유개수', '이름개수']].to_html(index=False, escape=False),
                           space_comparison=space_df.to_html(index=False, escape=False),
                           bracket_comparison=bracket_df.to_html(index=False, escape=False),
                           duplicated_file='duplicated_names.csv',
                           mismatch_file='name_mismatch.csv',
                           space_comparison_file='space_comparison.csv',
                           bracket_comparison_file='bracket_comparison.csv')

# 파일 다운로드 처리 (교육 관련)
@app.route('/edu/download/<filename>')
def download_edu_file(filename):
    file_path = os.path.join(app.config['PROCESSED_FOLDER'], filename)
    return send_file(file_path, as_attachment=True)

"""
=============== 관내여비 담당자 서포터 기능 ===============
"""
@app.route('/trip')
def trip_index():
    return render_template('trip_index.html')

@app.route('/trip/upload', methods=['POST'])
def upload_and_process_trip_files():
    if 'trip_file' not in request.files or 'tag_file' not in request.files:
        return 'No file part'
    
    trip_file = request.files['trip_file']
    tag_file = request.files['tag_file']
    
    if trip_file.filename == '' or tag_file.filename == '':
        return 'No selected file'
    
    # 업로드된 파일을 저장할 경로 설정
    trip_path = os.path.join(app.config['UPLOAD_FOLDER'], 'trip_all.xlsx')
    tag_path = os.path.join(app.config['UPLOAD_FOLDER'], 'tag_all.xlsx')
    
    # 파일 저장
    trip_file.save(trip_path)
    tag_file.save(tag_path)
    
    # 엑셀 파일 처리
    try:
        # 출장 신청 데이터 불러오기
        df_trip = pd.read_excel(trip_path, header=1)
        col_tmp = pd.read_excel(trip_path, nrows=0).columns
        col_tmp = col_tmp[:col_tmp.get_loc('출장기간') - 1]
        df_trip.columns.values[:len(col_tmp)+1] = pd.read_excel(trip_path, nrows=0).columns[:8]
        # 관내출장 추출
        df_trip = df_trip[df_trip['근태항목'] == '관내출장']
        # 결재완료 추출
        df_trip = df_trip[df_trip['결재상태'].str.startswith('결재완료')]

        # 부산경제진흥원 부서 검증
        depts = df_trip['부서'].unique()
        bepa = ['경영기획실', '청년사업단', '산업인력지원단', '소상공인지원단', '기업지원단', '글로벌사업추진단', '부원장', '기업옴부즈맨실', '임원']
        for dept in depts:
            if dept not in bepa:
                raise ValueError("오류가 발생하였습니다. 데이터전략TF팀으로 연락 바랍니다.")

        # 필요한 컬럼만 추출
        df_trip = df_trip[['부서', '사원코드', '사원', '직급', '신청일', '시작일', '종료일', '시작시간', '종료시간',
                   '일수', '신청시간', '교통수단', '운전자', '출발지', '도착지', '경유지', 
                   '방문처', '목적', '내용']]
        
        # 태그 데이터 불러오기
        df_tag = pd.read_excel(tag_path)
        df_tag = df_tag[['태깅일자', '사원코드', '근태구분', '근무시간']]
        
        # 외출/복귀 시간 태깅
        df_trip[['외출태그', '복귀태그', '외출태그(인정)', '복귀태그(인정)']] = [None] * 4
        # 변수 정의
        cols = ['사원코드', '부서', '시작일', '시작시간', '종료시간']
        
        for i in range(len(df_trip)):
            # 변수 정의
            id, dept, date, str_time, end_time = df_trip.iloc[i, df_trip.columns.get_indexer(cols)]
            out_time, in_time, out_time_use, in_time_use = [None] * 4
            
            # 태그 이력 추출
            cond_date = df_tag['태깅일자'] == date
            cond_id = df_tag['사원코드'] == id
            df_cond = df_tag[cond_date & cond_id]
            
            # 외출 : 가장 늦게 찍은 기록
            try:
                out_time = df_cond[df_cond['근태구분'] == '외출']['근무시간'].iloc[-1]
            except IndexError:
                pass

            # 복귀 : 가장 먼저 찍은 기록 
            try:
                in_time = df_cond[df_cond['근태구분'] == '복귀']['근무시간'].iloc[0]
            except IndexError:
                pass

            # 신청시간과 태그시간이 겹치지 않는 경우
            if out_time and in_time:
                if (out_time > end_time) or (in_time < str_time):
                    out_time_use, in_time_use = ['불인정'] * 2
                    df_trip.iloc[i, df_trip.columns.get_indexer(['외출태그', '복귀태그', '외출태그(인정)', '복귀태그(인정)'])] = out_time, in_time, out_time_use, in_time_use

            # 출장 시작 9시// 출장 종료 18시 : 자동 설정
            if (str_time <= '09:00')&(pd.isna(out_time)):
                out_time_use = str_time
            if (end_time >= '18:00')&(pd.isna(in_time)):
                in_time_use = end_time
            
            # 출장 시작보다 빨리 나간 경우 : 출장 시작 시간으로 설정
            if pd.isna(out_time):
                pass
            else:
                if str_time > out_time:
                    out_time_use = str_time
                
            # 출장 종료보다 늦게 들어온 경우 : 출장 종료 시간으로 설정
            if pd.isna(in_time):
                pass
            else:
                if end_time < in_time:
                    in_time_use = end_time

            df_trip.iloc[i, df_trip.columns.get_indexer(['외출태그', '복귀태그', '외출태그(인정)', '복귀태그(인정)'])] = out_time, in_time, out_time_use, in_time_use
        
        df_trip['외출태그(인정)'] = df_trip['외출태그(인정)'].fillna(df_trip['외출태그'])
        df_trip['복귀태그(인정)'] = df_trip['복귀태그(인정)'].fillna(df_trip['복귀태그'])

        df_trip['외출태그(인정)'] = df_trip['외출태그(인정)'].apply(lambda x : None if x=='불인정' else x)
        df_trip['복귀태그(인정)'] = df_trip['복귀태그(인정)'].apply(lambda x : None if x=='불인정' else x)

        # 출장시간 계산
        df_trip['외출태그(산출)'] = pd.to_datetime(df_trip['외출태그(인정)'], format='%H:%M')
        df_trip['복귀태그(산출)'] = pd.to_datetime(df_trip['복귀태그(인정)'], format='%H:%M')
        
        total_time = (df_trip['복귀태그(산출)'] - df_trip['외출태그(산출)'])
        df_trip['출장시간(산출)/분'] = total_time.dt.total_seconds() // 60
        df_trip['출장시간'] = total_time.apply(lambda x: None if pd.isna(x) else f'{x.components.hours}:{x.components.minutes:02d}')
        
        # 여비 계산
        df_trip['여비'] = 0
        for i in range(len(df_trip)):
            car, time = df_trip.iloc[i, df_trip.columns.get_indexer(['교통수단', '출장시간(산출)/분'])]

            if pd.isna(time): m = 0
            elif time < 240: m = 10000
            else: m = 20000

            if car == '관용차량': m -= 10000
            if m < 0: m = 0

            df_trip.iloc[i, df_trip.columns.get_loc('여비')] = m
        
    except Exception as e:
        return f"파일 처리 중 오류 발생: {str(e)}"
    
    # 부서별 엑셀 파일 저장
    department_files = []
    for dept, group in df_trip.groupby('부서'):
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{dept}_관내여비.xlsx')
        try:
            group.sort_values(by=['사원', '시작일'], inplace=True)
        except:
            group.sort_values(by=['사원'], inplace=True)
        group = group[['부서', '사원', '직급', '신청일', '시작일', '종료일', '시작시간', 
        '종료시간', '일수', '신청시간', '외출태그', '복귀태그', '외출태그(인정)', '복귀태그(인정)',
        '출장시간', '여비', '교통수단', '운전자', '출발지', '도착지', '경유지', '방문처', '목적', '내용']]
        group.to_excel(file_path, index=False)
        department_files.append(file_path)

    return render_template('trip_result.html', department_files=department_files)


# 파일 다운로드 처리 (관내여비 관련)
@app.route('/trip/download/<file_name>')
def download_trip_file(file_name):
    # 업로드 폴더 경로 설정
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file_name)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return f'파일 {file_name}을 찾을 수 없습니다.'

# 압축 안해용
# @app.route('/trip/download_zip/<zip_file_name>')
# def download_zip(zip_file_name):
#     # 압축된 파일 다운로드
#     zip_file_path = os.path.join(app.config['UPLOAD_FOLDER'], zip_file_name)
#     if os.path.exists(zip_file_path):
#         return send_file(zip_file_path, as_attachment=True)
#     else:
#         return f'압축 파일 {zip_file_name}을 찾을 수 없습니다.'

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