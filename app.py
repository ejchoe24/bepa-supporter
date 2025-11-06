from flask import Flask, render_template, request, send_file, session, redirect, url_for # type: ignore
import pandas as pd # type: ignore
import os
import re
import warnings
warnings.filterwarnings('ignore')

app = Flask(__name__)
app.secret_key = '6001312'

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
================== 관내여비 ========================
"""

@app.route('/trip')
def trip_index():
    return render_template('trip_index.html')


@app.route('/trip/upload', methods=['POST'])
def upload_trip():
    # 파일 유효성 확인
    if 'trip_file' not in request.files or 'tag_file' not in request.files:
        return 'No file part'
    
    trip_file = request.files['trip_file']
    tag_file = request.files['tag_file']

    if trip_file.filename == '' or tag_file.filename == '':
        return 'No selected file'
    
    # 저장 경로 설정
    trip_path = os.path.join(app.config['UPLOAD_FOLDER'], 'trip.xlsx')
    tag_path = os.path.join(app.config['UPLOAD_FOLDER'], 'tag.xlsx')

    trip_file.save(trip_path)
    tag_file.save(tag_path)

    # 세션에 경로 저장
    session['trip_path'] = trip_path
    session['tag_path'] = tag_path

    # ===== 엑셀에서 명단 추출 =====
    try:
        try:
            df_trip = pd.read_excel(trip_path, header=[0, 1])
            df_trip.columns = df_trip.columns.map(lambda x: x[0] if 'Unnamed' in str(x[1]) else x[1])
        except Exception:
            df_trip = pd.read_excel(trip_path)
    except Exception as e:
        return f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}"

    # 사원코드 / 사원 컬럼 확인
    if not set(['사원코드', '사원']).issubset(df_trip.columns):
        return "엑셀 파일에 '사원코드' 또는 '사원' 컬럼이 없습니다."

    unique_emp = df_trip[['사원코드', '사원']].dropna().drop_duplicates()

    emps = []
    for _, row in unique_emp.iterrows():
        emps.append({
            'employee_code': str(row['사원코드']),
            'name': str(row['사원']),
            'start_time': '09:00',
            'end_time': '18:00'
        })

    # 세션 저장
    session['trip_emp'] = emps

    # 다음 페이지로 이동
    return redirect(url_for('emp_times'))


@app.route('/trip/times')
def emp_times():
    emps = session.get('trip_emp', [])
    return render_template('trip_times.html', emps=emps)


@app.route("/trip/times/save", methods=["POST"])
def times_save():
    emps = session.get("trip_emp", [])
    new_emps = []

    for e in emps:
        code = e["employee_code"]
        start = request.form.get(f"start_{code}", "09:00")
        end = request.form.get(f"end_{code}", "18:00")

        new_emps.append({
            "employee_code": code,
            "name": e["name"],
            "start_time": start,
            "end_time": end
        })

    # 세션 업데이트
    session["trip_emp"] = new_emps

    return redirect(url_for("process_trip"))


@app.route('/trip/process', methods=['GET'])
def process_trip():
    trip_path = session.get('trip_path')
    tag_path = session.get('tag_path')
    emps = session.get('trip_emp')

    if not trip_path or not tag_path or not emps:
        return redirect(url_for('trip_index'))

    # 1) 사원별 출퇴근 시간 맵
    time_map = {e['employee_code']: (e['start_time'], e['end_time']) for e in emps}

    # 2) 출장 신청 데이터 로드
    try:
        # 출장 신청 데이터 불러오기
        df_trip = pd.read_excel(trip_path, heder=[0, 1])
        df_trip.columns = df_trip.columns.map(lambda x : x[0] if 'Unnamed' in str(x[1]) else x[1])

        for col in ['근태항목', '결재상태', '부서']:
            if col not in df_trip.columns:
                return f"출장 신청 내역에 '{col}' 열이 없습니다."
        
        # 우리 회사 부서 검증
        depts = df_trip['부서'].unique()
        bepa = ['경영기획실', '청년사업단', '산업인력지원단', '소상공인지원단', '기업지원단', '글로벌사업추진단', '부원장', '기업옴부즈맨실', '임원']
        for dept in depts:
            if dept not in bepa:
                return "오류가 발생하였습니다. 데이터전략TF팀으로 문의 부탁드립니다."

            
        df_trip = df_trip[df_trip['근태항목']=='관내출장']
        df_trip = df_trip[df_trip['결재상태'].str.startswith('결재완료')]

        # 필요한 컬럼만 추출
        cols = df_trip[['부서', '사원코드', '사원', '직급', '신청일', '시작일', '종료일', '시작시간', '종료시간', '일수',
                        '신청시간', '교통수단', '운전자', '출발지', '도착지', '경유지', '방문처', '목적', '내용']]
        cols = [c for c in cols if c in df_trip.columns]
        df_trip = df_trip[cols]

        # 3) 태깅 데이터 로드
        try:
            df_tag = pd.read_excel(tag_path)
            cols = ['태깅일자', '사원코드', '근태구분', '근무시간']
            for col in cols:
                if col not in df_tag.columns:
                    return f"태깅 파일에 '{col}' 열이 없습니다."

            df_tag = df_tag[cols]

            # 4) 외출/복귀 시간 계산 
            tag_cols = ['외출태그', '복귀태그', '외출태그(인정)', '복귀태그(인정)']
            df_trip[cols] = [None] * 4
            cols = ['사원코드', '부서', '시작일', '시작시간', '종료시간']

            for i in range(len(df_trip)):
                # 변수 정의 및 초기화
                id, dept, date, str_time, end_time = df_trip.iloc[i, df_trip.columns.get_indexer(cols)]
                out_time, in_time, valid_out, valid_in = [None] * 4

                # 태깅 이력 추출
                cond_date = df_tag['태깅일자'] == date
                cond_id = df_tag['사원코드'] == id
                df_cond = df_tag[cond_date & cond_id]

                # 외출 : 가장 늦게 찍은 기록
                try: out_time = df_cond[df_cond['근태구분']=='외출']['근무시간'].iloc[-1]
                except IndexError: pass

                # 복귀 : 가장 빨리 찍은 기록
                try: in_time = df_cond[df_cond['근태구분']=='복귀']['근무시간'].iloc[0]
                except IndexError: pass

                # 신청시간과 태그시간이 겹치지 않는 경우
                if out_time and in_time:
                    if (out_time > end_time) or (in_time < str_time):
                        valid_out, valid_in = ['불인정'] * 2
                        df_trip.iloc[i, df_trip.columns.get_indexer(tag_cols)] = out_time, in_time, valid_out, valid_in

                # 출장 시작 = 출근 시간 // 출장 종료 = 퇴근 시간 : 자동 설정
                str_work, end_work = time_map.get(id, ("09:00", "18:00"))

                if (str_time <= str_work) & (pd.isna(out_time)):
                    valid_out = str_time
                if (end_time >= end_work) & (pd.isna(in_time)):
                    valid_in = end_time

                # 출장 시작보다 빨리 나간 경우 : 출장 시작 시간으로 설정
                if pd.isna(out_time): pass
                else:
                    if str_time > out_time:
                        valid_out = str_time

                # 출장 종료보다 늦게 들어온 경우 : 출장 종료 시간으로 설정
                if pd.isna(in_time): pass
                else:
                    if end_time < in_time:
                        valid_in = end_time

                df_trip.iloc[i, df_trip.columns.get_indexer(tag_cols)] = out_time, in_time, valid_out, valid_in


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
        
        # 엑셀 파일 저장
        files = []
        for dept, group in df_trip.groupby('부서'):
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], f'{dept}_관내여비.xlsx')
            try: group.sort_values(by=['사원', '시작일'], inplace=True)
            except: group.sort_values(by=['사원'], inplace=True)
            
            cols = ['부서', '사원', '직급', '신청일', '시작일', '종료일', '시작시간',
                    '종료시간', '일수', '신청시간', '외출태그', '복귀태그', '외출태그(인정)', '복귀태그(인정)',
                    '출장시간', '여비', '교통수단', '운전자', '출발지', '도착지', '경유지', '방문처', '목적', '내용']
            cols = [c for c in cols if c in group]
            group = group[cols]
            
            group.to_excel(file_path, index=False)
            files.append(file_path)

    except Exception as e:
            return f"파일 처리 중 오류 발생: {str(e)}"

    return render_template('trip_result.html', files=files)

@app.route('/trip/download/<file_name>')
def download_trip_file(file_name):
    # 업로드 폴더 경로 설정
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file_name)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else: return f'파일 {file_name}을 찾을 수 없습니다.'