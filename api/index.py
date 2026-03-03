from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import msoffcrypto
import pandas as pd
import io
import logging
import re

app = Flask(__name__)
CORS(app)

logging.basicConfig(level=logging.INFO)

# 데이터 추출 및 정제 보조 함수
def get_col_safe(df, name, n):
    target_clean = str(name).replace(" ", "")
    for c in df.columns:
        if str(c).replace(" ", "") == target_clean:
            return df[c].fillna("").astype(str).str.strip()
    return pd.Series([""] * n)

# 특수기호 삭제 함수
def remove_special_chars(text):
    if not text: return ""
    # 한글, 숫자, 영어, 공백만 남기고 모두 삭제
    return re.sub(r'[^가-힣a-zA-Z0-9\s]', '', str(text))

@app.route('/api', methods=['POST'])
def convert():
    try:
        if 'file' not in request.files:
            return jsonify({"error": "파일이 없습니다."}), 400
        
        file = request.files['file']
        password = request.form.get('password', '').strip()
        
        # 1. 네이버 엑셀 암호 해제
        input_buffer = io.BytesIO(file.read())
        decrypted_buffer = io.BytesIO()
        try:
            office_file = msoffcrypto.OfficeFile(input_buffer)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted_buffer)
            decrypted_buffer.seek(0)
        except:
            return jsonify({"error": "비밀번호가 틀렸습니다."}), 403

        # 2. 데이터 읽기
        df_raw = pd.read_excel(decrypted_buffer, engine='openpyxl', header=None, dtype=str)
        header_row_idx = -1
        for i, row in df_raw.head(30).iterrows():
            if "수취인명" in [str(v).replace(" ", "") for v in row.values if pd.notna(v)]:
                header_row_idx = i
                break
        
        if header_row_idx == -1:
            return jsonify({"error": "네이버 엑셀 양식이 아닙니다."}), 400

        df = df_raw.iloc[header_row_idx + 1:].copy()
        df.columns = df_raw.iloc[header_row_idx]
        df = df.reset_index(drop=True)
        num_rows = len(df)

        # 3. 데이터 추출 및 정제
        names = get_col_safe(df, "수취인명", num_rows)
        zips = get_col_safe(df, "우편번호", num_rows)
        addr1 = get_col_safe(df, "기본배송지", num_rows)
        addr2 = get_col_safe(df, "상세배송지", num_rows).apply(lambda x: "." if not x or x.strip() == "" else x)
        
        # [수정] 5행 일반전화: 010으로 시작하면 강제로 비움
        tel_home = get_col_safe(df, "수취인연락처2", num_rows)
        tel_home = tel_home.apply(lambda x: "" if x.startswith("010") else x)
        
        tel_mobile = get_col_safe(df, "수취인연락처1", num_rows) # 휴대전화
        product = get_col_safe(df, "상품명", num_rows).str.slice(0, 20)
        
        # [수정] 12행 배송메세지: 특수기호 삭제
        memo = get_col_safe(df, "배송메세지", num_rows)
        memo = memo.apply(remove_special_chars)

        # 4. 우체국 표준 17개 항목 파이프(|) 결합
        lines = []
        for i in range(num_rows):
            row = [
                names[i],               # 1: 받는 분
                zips[i],                # 2: 우편번호
                addr1[i],               # 3: 주소
                addr2[i],               # 4: 상세주소
                tel_home[i],            # 5: 일반전화 (010 제거됨)
                tel_mobile[i],          # 6: 휴대전화
                "3",                    # 7: 중량
                "80",                   # 8: 부피
                "농/수/축산물(일반)",   # 9: 내용품코드
                product[i],             # 10: 내용물 (상품명)
                "미신청",               # 11: 배달방식
                memo[i],                # 12: 배송시요청사항 (특수기호 제거됨)
                "N",                    # 13: 분할접수 여부
                "", "", "", ""          # 14~17: 빈칸
            ]
            lines.append("|".join(row))
        
        content = "\n".join(lines)

        # 5. TXT 파일 전송
        output = io.BytesIO(content.encode('cp949', errors='replace'))
        output.seek(0)

        return send_file(
            output,
            mimetype='text/plain',
            as_attachment=True,
            download_name=f"post_upload_fixed.txt"
        )

    except Exception as e:
        return jsonify({"error": f"서버 오류: {str(e)}"}), 500
