from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import msoffcrypto
import pandas as pd
import io
import logging

app = Flask(__name__)
CORS(app)

logging.basicConfig(level=logging.INFO)

# 보조 함수: 컬럼 데이터 추출
def get_col_safe(df, name, n):
    target_clean = str(name).replace(" ", "")
    for c in df.columns:
        if str(c).replace(" ", "") == target_clean:
            # nan 값은 빈 문자열로, 모든 값은 문자열로 변환 후 공백 제거
            return df[c].fillna("").astype(str).str.strip()
    return pd.Series([""] * n)

@app.route('/api', methods=['POST'])
def convert():
    try:
        # 1. 파일 및 비밀번호 수신 확인
        if 'file' not in request.files:
            return jsonify({"error": "파일이 없습니다."}), 400
        
        file = request.files['file']
        password = request.form.get('password', '').strip()
        
        if not password:
            return jsonify({"error": "비밀번호를 입력해주세요."}), 400

        # 2. 파일 읽기 및 복호화
        file_data = file.read()
        input_buffer = io.BytesIO(file_data)
        decrypted_buffer = io.BytesIO()

        try:
            office_file = msoffcrypto.OfficeFile(input_buffer)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted_buffer)
            decrypted_buffer.seek(0)
        except Exception:
            return jsonify({"error": "비밀번호가 틀렸거나 파일 해독에 실패했습니다."}), 403

        # 3. 엑셀 로드
        try:
            df_raw = pd.read_excel(decrypted_buffer, engine='openpyxl', header=None, dtype=str)
        except Exception as e:
            return jsonify({"error": f"엑셀 읽기 실패: {str(e)}"}), 400

        # 4. 네이버 헤더(수취인명) 찾기
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

        # 5. 데이터 정제 (우체국 텍스트 양식 7개 항목 순서)
        names = get_col_safe(df, "수취인명", num_rows)
        zips = get_col_safe(df, "우편번호", num_rows)
        addr1 = get_col_safe(df, "기본배송지", num_rows)
        
        # 상세주소 필수 처리 (비어있으면 마침표 입력)
        addr2 = get_col_safe(df, "상세배송지", num_rows)
        addr2 = addr2.apply(lambda x: "." if not x or x.strip() == "" else x)
        
        tel1 = get_col_safe(df, "수취인연락처2", num_rows) # 일반전화
        tel2 = get_col_safe(df, "수취인연락처1", num_rows) # 휴대전화
        
        # 7번째 항목: 참조번호 혹은 상품명 (20자 제한)
        product = get_col_safe(df, "상품명", num_rows).str.slice(0, 20)

        # 6. 파이프(|)로 구분된 텍스트 내용 생성
        lines = []
        for i in range(num_rows):
            # 형식: 성명|우편번호|주소|상세주소|일반전화|휴대전화|상품명
            line = f"{names[i]}|{zips[i]}|{addr1[i]}|{addr2[i]}|{tel1[i]}|{tel2[i]}|{product[i]}"
            lines.append(line)
        
        content = "\n".join(lines)

        # 7. 텍스트 파일(TXT)로 전송
        # cp949 인코딩은 윈도우 우체국 시스템 호환용입니다.
        output_buffer = io.BytesIO(content.encode('cp949', errors='replace'))
        output_buffer.seek(0)

        return send_file(
            output_buffer,
            mimetype='text/plain',
            as_attachment=True,
            download_name=f"post_office_upload_{pd.Timestamp.now().strftime('%m%d')}.txt"
        )

    except Exception as e:
        logging.exception("Conversion failed")
        return jsonify({"error": f"서버 오류: {str(e)}"}), 500
