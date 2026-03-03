from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import msoffcrypto
import pandas as pd
import io
import logging

app = Flask(__name__)
CORS(app)

logging.basicConfig(level=logging.INFO)

@app.route('/api', methods=['POST'])
def convert():
    try:
        if 'file' not in request.files:
            return jsonify({"error": "파일이 없습니다."}), 400
        
        file = request.files['file']
        password = request.form.get('password', '').strip()
        
        file_data = file.read()
        input_buffer = io.BytesIO(file_data)
        decrypted_buffer = io.BytesIO()

        # 1. 암호 해제
        try:
            office_file = msoffcrypto.OfficeFile(input_buffer)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted_buffer)
            decrypted_buffer.seek(0)
        except Exception as e:
            return jsonify({"error": "비밀번호가 일치하지 않습니다.", "details": str(e)}), 403

        # 2. 엑셀 로드 (헤더를 일단 무시하고 데이터만 가져옴)
        try:
            # header=None으로 읽어서 모든 데이터를 일단 가져옵니다.
            df_raw = pd.read_excel(decrypted_buffer, engine='openpyxl', header=None)
        except Exception as e:
            return jsonify({"error": "엑셀 읽기 실패", "details": str(e)}), 400

        # 3. 진짜 데이터 시작점 찾기 (네이버는 보통 2번째나 3번째 줄부터 진짜 데이터임)
        # '수취인명'이라는 글자가 포함된 행을 찾아서 그 아래부터 데이터로 인식합니다.
        header_row_idx = 0
        for i, row in df_raw.head(10).iterrows():
            if "수취인명" in row.values:
                header_row_idx = i
                break
        
        # 헤더 아래부터 데이터로 슬라이싱
        df = df_raw.iloc[header_row_idx + 1:].copy()
        df.columns = df_raw.iloc[header_row_idx] # 헤더 행 설정
        df = df.reset_index(drop=True)

        num_rows = len(df)
        if num_rows == 0:
            return jsonify({"error": "변환할 데이터가 없습니다."}), 400

        # 4. 열 이름을 기반으로 찾되, 못 찾으면 순서(Index)로 백업 시도
        def safe_get(col_name, fallback_idx):
            if col_name in df.columns:
                return df[col_name]
            # 컬럼명이 없으면 지정된 순서(번호)로 가져옴
            if fallback_idx < len(df.columns):
                return df.iloc[:, fallback_idx]
            return [""] * num_rows

        # [네이버 표준 양식 기준 열 순서 (0부터 시작)]
        # 수취인명(10번), 우편번호(44번), 기본배송지(42번), 상세배송지(43번), 연락처1(40번), 연락처2(41번), 상품명(16번), 배송메세지(45번)
        # ※ 양식마다 다를 수 있어 가장 안전한 '이름 찾기'와 '번호'를 병행합니다.

        mapping = {
            "받는 분": safe_get("수취인명", 10),
            "우편번호": safe_get("우편번호", 44),
            "주소(시도+시군구+도로명+건물번호)": safe_get("기본배송지", 42),
            "상세주소(동, 호수, 洞명칭, 아파트, 건물명 등)": safe_get("상세배송지", 43),
            "일반전화(02-1234-5678)": safe_get("수취인연락처2", 41),
            "휴대전화(010-1234-5678)": safe_get("수취인연락처1", 40),
            "중량(kg)": [3] * num_rows,
            "부피(cm)=가로+세로+높이": [80] * num_rows,
            "내용품코드": ["농/수/축산물(일반)"] * num_rows,
            "내용물": safe_get("상품명", 16),
            "배달방식": ["일반소포"] * num_rows,
            "배송시요청사항": safe_get("배송메세지", 45),
            "분할접수 여부(Y/N)": ["N"] * num_rows
        }
        
        post_df = pd.DataFrame(mapping)

        # 5. 결과 파일 생성
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            post_df.to_excel(writer, index=False)
        output_buffer.seek(0)

        return send_file(
            output_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"우체국_양식_{pd.Timestamp.now().strftime('%m%d')}.xlsx"
        )

    except Exception as e:
        logging.error(f"Error: {str(e)}")
        return jsonify({"error": f"변환 실패: {str(e)}"}), 500
