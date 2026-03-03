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
        
        # 1. 비밀번호를 받을 때 인코딩 문제를 방지하기 위해 명시적으로 처리
        password = request.form.get('password', '')
        if not password:
            return jsonify({"error": "비밀번호를 입력해주세요."}), 400
        
        # 문자열 양 끝의 공백만 제거 (영문 대소문자는 유지)
        password = password.strip()
        
        file_data = file.read()
        input_buffer = io.BytesIO(file_data)
        decrypted_buffer = io.BytesIO()

        # 2. 암호 해제 시도
        try:
            office_file = msoffcrypto.OfficeFile(input_buffer)
            
            # 영문 비밀번호의 경우 특정 인코딩에서 문제가 생길 수 있으므로 
            # 라이브러리가 지원하는 기본 문자열 방식으로 전달
            office_file.load_key(password=str(password))
            office_file.decrypt(decrypted_buffer)
            decrypted_buffer.seek(0)
        except Exception as e:
            logging.error(f"Decryption failed: {str(e)}")
            return jsonify({
                "error": "비밀번호가 일치하지 않습니다.",
                "details": "영문 대소문자를 다시 확인해주세요."
            }), 403

        # 3. 엑셀 데이터 로드
        try:
            # 해독된 버퍼를 읽을 때 엔진을 openpyxl로 고정
            df_raw = pd.read_excel(decrypted_buffer, engine='openpyxl', header=None)
        except Exception as e:
            return jsonify({"error": "엑셀 읽기 실패", "details": str(e)}), 400

        # 4. 데이터 시작점(헤더) 찾기
        header_row_idx = 0
        for i, row in df_raw.head(20).iterrows():
            if "수취인명" in [str(v).replace(" ", "") for v in row.values]:
                header_row_idx = i
                break
        
        df = df_raw.iloc[header_row_idx + 1:].copy()
        df.columns = df_raw.iloc[header_row_idx]
        df = df.reset_index(drop=True)

        num_rows = len(df)
        if num_rows == 0:
            return jsonify({"error": "변환할 데이터가 없습니다."}), 400

        # 5. 데이터 매핑 함수 (열 이름 기반)
        def get_col(name):
            # 실제 엑셀의 컬럼명에서 공백을 제거하고 비교
            for real_col in df.columns:
                if str(real_col).replace(" ", "") == name:
                    return df[real_col]
            return [""] * num_rows

        mapping = {
            "받는 분": get_col("수취인명"),
            "우편번호": get_col("우편번호"),
            "주소(시도+시군구+도로명+건물번호)": get_col("기본배송지"),
            "상세주소(동, 호수, 洞명칭, 아파트, 건물명 등)": get_col("상세배송지"),
            "일반전화(02-1234-5678)": get_col("수취인연락처2"),
            "휴대전화(010-1234-5678)": get_col("수취인연락처1"),
            "중량(kg)": [3] * num_rows,
            "부피(cm)=가로+세로+높이": [80] * num_rows,
            "내용품코드": ["농/수/축산물(일반)"] * num_rows,
            "내용물": get_col("상품명"),
            "배달방식": ["일반소포"] * num_rows,
            "배송시요청사항": get_col("배송메세지"),
            "분할접수 여부(Y/N)": ["N"] * num_rows
        }
        
        post_df = pd.DataFrame(mapping)

        # 6. 파일 생성 및 반환
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            post_df.to_excel(writer, index=False)
        output_buffer.seek(0)

        return send_file(
            output_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name="post_office_converted.xlsx"
        )

    except Exception as e:
        logging.error(f"Error: {str(e)}")
        return jsonify({"error": f"변환 실패: {str(e)}"}), 500
