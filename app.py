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
    if 'trip_file' not in request.files or 'tag_file' not in request.files:
        return 'No file part'
    
    trip_file = request.files['trip_file']
    tag_file = request.files['tag_file']

    if trip_file.filename == '' or tag_file.filename == '':
        return 'No selected file'
    
    trip_path = os.path.join(app.config['UPLOAD_FOLDER'], 'trip.xlsx')
    tag_path = os.path.join(app.config['UPLOAD_FOLDER'], 'tag.xlsx')

    trip_file.save(trip_path)
    tag_file.save(tag_path)

    session['trip_path'] = trip_path
    session['tag_path'] = tag_path

    # 명단 추출
    df_trip = pd.read_excel(trip_path, heder=[0, 1])
    df_trip.columns = df_trip.columns.map(lambda x : x[0] if 'Unnamed' in str(x[1]) else x[1])
    unique_emp = df_trip[['사원코드', '사원']].dropna().drop_duplicates()

    emps = []
    for _, row in unique_emp.iterrows():
        emps.append({
            'employee_code': str(row['사원코드']),
            'name': str(row['사원']),
            'start_time': '09:00',
            'end_time': '18:00'
        })
    session['trip_emp'] = emps

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
    session["trip_users"] = new_emps
    return redirect(url_for("process_trip_with_user_times"))


    try:
        # 출장 신청 데이터 불러오기
        df_trip = pd.read_excel(trip_path, heder=[0, 1])
        df_trip.columns = df_trip.columns.map(lambda x : x[0] if 'Unnamed' in str(x[1]) else x[1])

        # 필요한 컬럼만 추출
        df_trip = df_trip[['부서', '사원코드', '사원', '직급', '신청일', '시작일', '종료일', '시작시간', '종료시간', '일수',
                           '신청시간', '교통수단', '운전자', '출발지', '도착지', '경유지', '방문처', '목적', '내용']]
        # 관내출장 추출
        df_trip = df_trip[df_trip['근태항목']=='관내출장']
        # 결재완료 추출
        df_trip = df_trip[df_trip['결재상태'].str.startwith('결재완료')]

        # 부산경제진흥원 부서 검증
        depts = df_trip['부서'].unique()
        bepa = ['경영기획실', '청년사업단', '산업인력지원단', '소상공인지원단', '기업지원단', '글로벌사업추진단', '부원장', '기업옴부즈맨실', '임원']
        for dept in depts:
            if dept not in bepa:
                raise ValueError("오류가 발생하였습니다. 데이터전략TF팀으로 문의 부탁드립니다.")
        

        # 태그 데이터 불러오기
        df_tag = pd.read_excel(tag_path)
        df_tag = df_tag[['태깅일자', '사원코드', '근태구분', '근무시간']]

        # 외출/복귀 시간 계산
        tag_cols = ['외출태그', '복귀태그', '외출태그(인정)', '복귀태그(인정)']
        df_trip[tag_cols] = [None] * 4
        cols = ['사원코드', '부서', '시작일', '시작시간', '종료시간']

        for i in range(len(df_trip)):
            # 변수 정의 및 초기화
            id, dept, date, str_time, end_time = df_trip.iloc[i, df_trip.columns.get_indexer(cols)]
            out_time, in_time, valid_out, valid_in = [None] * 4

            # 태그 이력 추출
            cond_date = df_tag['태깅일자'] == date
            cond_id = df_tag['사원코드'] == id
            df_cond = df_tag[cond_date & cond_id]

            # 외출 : 가장 늦게 찍은 기록
            try:
                out_time = df_cond[df_cond['근태구분']=='외출']['근무시간'].iloc[-1]
            except IndexError:
                pass

            # 복귀 : 가장 빨리 찍은 기록
            try:
                in_time = df_cond[df_cond['근태구분'] == '복귀']['근무시간'].iloc[0]
            except IndexError:
                pass

            # 신청시간과 태그시간이 겹치지 않는 경우
            if out_time and in_time:
                if (out_time > end_time) or (in_time < str_time):
                    valid_out, valid_in = ['불인정'] * 2
                    df_trip.iloc[i, df_trip.columns.get_indexer(tag_cols)] = out_time, in_time, valid_out, valid_in

            # 출장 시작 9시 // 출장 종료 18시 