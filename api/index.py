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

        # 2. 엑셀 로드
        try:
            df = pd.read_excel(decrypted_buffer, engine='openpyxl')
        except Exception as e:
            return jsonify({"error": "엑셀 읽기 실패", "details": str(e)}), 400

        num_rows = len(df)
        if num_rows == 0:
            return jsonify({"error": "엑셀에 데이터가 없습니다."}), 400

        # 3. 데이터 추출 함수 (알려주신 이름을 최우선으로 찾습니다)
        def get_val(naver_col):
            if naver_col in df.columns:
                return df[naver_col]
            return [""] * num_rows

        # 4. 우체국 양식에 맞게 1:1 매핑
        mapping = {
            "받는 분": get_val("수취인명"),
            "우편번호": get_val("우편번호"),
            "주소(시도+시군구+도로명+건물번호)": get_val("기본배송지"),
            "상세주소(동, 호수, 洞명칭, 아파트, 건물명 등)": get_val("상세배송지"),
            "일반전화(02-1234-5678)": get_val("수취인연락처2"), # 보조 연락처
            "휴대전화(010-1234-5678)": get_val("수취인연락처1"),
            "중량(kg)": [3] * num_rows,
            "부피(cm)=가로+세로+높이": [80] * num_rows,
            "내용품코드": ["농/수/축산물(일반)"] * num_rows,
            "내용물": get_val("상품명"), # 상품명도 함께 가져옵니다
            "배달방식": ["일반소포"] * num_rows,
            "배송시요청사항": get_val("배송메세지"),
            "분할접수 여부(Y/N)": ["N"] * num_rows
        }
        
        post_df = pd.DataFrame(mapping)

        # 5. 파일 생성
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
        logging.error(f"Mapping error: {str(e)}")
        return jsonify({"error": f"데이터 변환 중 오류 발생: {str(e)}"}), 500
