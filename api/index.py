from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import msoffcrypto
import pandas as pd
import io
import logging

app = Flask(__name__)
CORS(app)

logging.basicConfig(level=logging.INFO)

# 데이터 추출 보조 함수
def get_col_safe(df, name, n):
    target_clean = str(name).replace(" ", "")
    for c in df.columns:
        if str(c).replace(" ", "") == target_clean:
            return df[c].fillna("").astype(str).str.strip()
    return pd.Series([""] * n)

@app.route('/api', methods=['POST'])
def convert():
    try:
        # 파일 확인
        if 'file' not in request.files:
            return jsonify({"error": "파일이 업로드되지 않았습니다."}), 400
        
        file = request.files['file']
        password = request.form.get('password', '').strip()
        
        if not password:
            return jsonify({"error": "비밀번호를 입력해주세요."}), 400

        # 파일 읽기 및 암호 해제
        file_data = file.read()
        input_buffer = io.BytesIO(file_data)
        decrypted_buffer = io.BytesIO()

        try:
            office_file = msoffcrypto.OfficeFile(input_buffer)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted_buffer)
            decrypted_buffer.seek(0)
        except Exception:
            return jsonify({"error": "비밀번호가 틀렸거나 파일이 손상되었습니다."}), 403

        # 엑셀 로드
        try:
            df_raw = pd.read_excel(decrypted_buffer, engine='openpyxl', header=None, dtype=str)
        except Exception as e:
            return jsonify({"error": f"엑셀 파일을 읽을 수 없습니다: {str(e)}"}), 400

        # 네이버 헤더 찾기
        header_row_idx = -1
        for i, row in df_raw.head(30).iterrows():
            row_values = [str(v).replace(" ", "") for v in row.values if pd.notna(v)]
            if "수취인명" in row_values:
                header_row_idx = i
                break
        
        if header_row_idx == -1:
            return jsonify({"error": "네이버 엑셀 양식이 아닙니다."}), 400

        df = df_raw.iloc[header_row_idx + 1:].copy()
        df.columns = df_raw.iloc[header_row_idx]
        df = df.reset_index(drop=True)
        num_rows = len(df)

        # 데이터 정제 및 파이프(|) 파일 생성
        names = get_col_safe(df, "수취인명", num_rows)
        zips = get_col_safe(df, "우편번호", num_rows)
        addr1 = get_col_safe(df, "기본배송지", num_rows)
        addr2 = get_col_safe(df, "상세배송지", num_rows).apply(lambda x: "." if not x or x.strip() == "" else x)
        tel1 = get_col_safe(df, "수취인연락처2", num_rows)
        tel2 = get_col_safe(df, "수취인연락처1", num_rows)
        product = get_col_safe(df, "상품명", num_rows).str.slice(0, 15)

        lines = []
        for i in range(num_rows):
            # 순서: 성명|우편번호|주소|상세주소|일반전화|휴대전화|참조번호(상품명)
            line = f"{names[i]}|{zips[i]}|{addr1[i]}|{addr2[i]}|{tel1[i]}|{tel2[i]}|{product[i]}"
            lines.append(line)
        
        content = "\n".join(lines)

        # 결과 파일 전송 (cp949는 우체국 등 윈도우 시스템 표준)
        output_buffer = io.BytesIO(content.encode('cp949', errors='replace'))
        output_buffer.seek(0)

        return send_file(
            output_buffer,
            mimetype='text/plain',
            as_attachment=True,
            download_name=f"post_upload_{pd.Timestamp.now().strftime('%m%d')}.txt"
        )

    except Exception as e:
        logging.exception("Conversion error")
        return jsonify({"error": f"서버 내부 오류: {str(e)}"}), 500
