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
        
        # 1. 파일 데이터를 메모리에 로드
        file_data = file.read()
        input_buffer = io.BytesIO(file_data)
        decrypted_buffer = io.BytesIO()

        # 2. 암호 해제 (숫자 비밀번호에 최적화)
        try:
            office_file = msoffcrypto.OfficeFile(input_buffer)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted_buffer)
            decrypted_buffer.seek(0)
        except Exception as e:
            return jsonify({"error": "비밀번호가 일치하지 않습니다.", "details": str(e)}), 403

        # 3. 엑셀 데이터 로드
        try:
            # 헤더 없이 전체 데이터를 읽어옵니다.
            df_raw = pd.read_excel(decrypted_buffer, engine='openpyxl', header=None)
        except Exception as e:
            return jsonify({"error": "엑셀 읽기 실패", "details": str(e)}), 400

        # 4. '수취인명'이 있는 진짜 헤더 행 찾기
        header_row_idx = 0
        for i, row in df_raw.head(20).iterrows():
            # 행의 값들 중 '수취인명'이 포함되어 있는지 확인 (공백 제거 후 비교)
            row_values = [str(v).replace(" ", "") for v in row.values]
            if "수취인명" in row_values:
                header_row_idx = i
                break
        
        # 헤더 아래부터 데이터로 설정
        df = df_raw.iloc[header_row_idx + 1:].copy()
        df.columns = df_raw.iloc[header_row_idx]
        df = df.reset_index(drop=True)

        num_rows = len(df)
        if num_rows == 0:
            return jsonify({"error": "변환할 데이터가 없습니다."}), 400

        # 5. 안전하게 데이터를 가져오는 함수 (공백 무시)
        def get_col(target_name):
            target_clean = target_name.replace(" ", "")
            for real_col in df.columns:
                if str(real_col).replace(" ", "") == target_clean:
                    return df[real_col]
            return [""] * num_rows

        # 6. 요청하신 1:1 매핑 규칙 적용
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

        # 7. 엑셀 파일 생성 및 전송
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            post_df.to_excel(writer, index=False)
        output_buffer.seek(0)

        return send_file(
            output_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"우체국_변환완료_{pd.Timestamp.now().strftime('%m%d')}.xlsx"
        )

    except Exception as e:
        logging.error(f"Server Error: {str(e)}")
        return jsonify({"error": f"서버 오류: {str(e)}"}), 500
