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

def get_col_safe(df, name, n):
    target_clean = str(name).replace(" ", "")
    for c in df.columns:
        if str(c).replace(" ", "") == target_clean:
            return df[c].fillna("").astype(str).str.strip()
    return pd.Series([""] * n)

def remove_special_chars(text):
    if not text: return ""
    return re.sub(r'[^가-힣a-zA-Z0-9\s]', '', str(text))

@app.route('/api', methods=['POST'])
def convert():
    try:
        if 'file' not in request.files:
            return jsonify({"error": "파일이 없습니다."}), 400
        
        file = request.files['file']
        password = request.form.get('password', '').strip()
        # 프론트에서 선택한 품목 카테고리 수신 (기본값 설정)
        item_type = request.form.get('itemType', '농/수/축산물(일반)')
        
        input_buffer = io.BytesIO(file.read())
        decrypted_buffer = io.BytesIO()
        try:
            office_file = msoffcrypto.OfficeFile(input_buffer)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted_buffer)
            decrypted_buffer.seek(0)
        except:
            return jsonify({"error": "비밀번호가 틀렸습니다."}), 403

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

        names = get_col_safe(df, "수취인명", num_rows)
        zips = get_col_safe(df, "우편번호", num_rows)
        addr1 = get_col_safe(df, "기본배송지", num_rows)
        addr2 = get_col_safe(df, "상세배송지", num_rows).apply(lambda x: "." if not x or x.strip() == "" else x)
        
        raw_tel1 = get_col_safe(df, "수취인연락처1", num_rows)
        raw_tel2 = get_col_safe(df, "수취인연락처2", num_rows)

        final_tel_home = []
        final_tel_mobile = []
        mobile_prefixes = ("010", "050")

        for i in range(num_rows):
            t1 = raw_tel1[i].replace("-", "").strip()
            t2 = raw_tel2[i].replace("-", "").strip()
            mobile = ""; home = ""

            if t1.startswith(mobile_prefixes):
                mobile = t1
                if t2 and not t2.startswith(mobile_prefixes): home = t2
            elif t2.startswith(mobile_prefixes):
                mobile = t2
                if t1 and not t1.startswith(mobile_prefixes): home = t1
            else:
                home = t1
            
            final_tel_home.append(home)
            final_tel_mobile.append(mobile)
        
        memo = get_col_safe(df, "배송메세지", num_rows).apply(remove_special_chars)

        lines = []
        for i in range(num_rows):
            row = [
                names[i],               # 1: 받는 분
                zips[i],                # 2: 우편번호
                addr1[i],               # 3: 주소
                addr2[i],               # 4: 상세주소
                final_tel_home[i],      # 5: 일반전화
                final_tel_mobile[i],    # 6: 휴대전화
                "3",                    # 7: 중량
                "80",                   # 8: 부피
                item_type,              # 9: 내용품코드 (선택 품목)
                item_type,              # 10: 내용물 (상품명 대신 품목 입력)
                "미신청",               # 11: 배달방식
                memo[i],                # 12: 배송시요청사항
                "N",                    # 13: 분할접수 여부
                "", "", "", ""          # 14~17: 빈칸
            ]
            lines.append("|".join(row))
        
        content = "\n".join(lines)
        output = io.BytesIO(content.encode('cp949', errors='replace'))
        output.seek(0)

        return send_file(
            output,
            mimetype='text/plain',
            as_attachment=True,
            download_name=f"post_upload_{pd.Timestamp.now().strftime('%m%d')}.txt"
        )

    except Exception as e:
        return jsonify({"error": f"서버 오류: {str(e)}"}), 500
