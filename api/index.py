from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import msoffcrypto
import pandas as pd
import io
import logging

app = Flask(__name__)
CORS(app)

# 로그 설정 (Vercel 대시보드에서 더 자세히 볼 수 있게 함)
logging.basicConfig(level=logging.INFO)

@app.route('/api', methods=['POST'])
def convert():
    try:
        if 'file' not in request.files:
            return jsonify({"error": "파일이 선택되지 않았습니다."}), 400
        
        file = request.files['file']
        # 입력된 비밀번호에서 앞뒤 공백 및 제어문자 제거
        password = request.form.get('password', '').strip()
        
        if not password:
            return jsonify({"error": "비밀번호를 입력해주세요."}), 400

        # 1. 원본 파일을 바이트 단위로 끝까지 읽기
        file_data = file.read()
        if not file_data:
            return jsonify({"error": "파일 내용이 비어있습니다."}), 400
        
        input_buffer = io.BytesIO(file_data)
        decrypted_buffer = io.BytesIO()

        # 2. 암호 해제 시도
        try:
            office_file = msoffcrypto.OfficeFile(input_buffer)
            # 비밀번호를 이용해 복호화 키 로드
            office_file.load_key(password=password)
            office_file.decrypt(decrypted_buffer)
            decrypted_buffer.seek(0)
        except Exception as decrypt_err:
            logging.error(f"Decryption failed: {str(decrypt_err)}")
            return jsonify({
                "error": "비밀번호가 일치하지 않거나, 지원하지 않는 암호화 방식입니다.",
                "details": str(decrypt_err)
            }), 403

        # 3. 데이터 로드 (pandas)
        try:
            # 해독된 스트림을 엑셀로 읽기
            df = pd.read_excel(decrypted_buffer, engine='openpyxl')
        except Exception as excel_err:
            logging.error(f"Excel read failed: {str(excel_err)}")
            return jsonify({"error": "엑셀 데이터를 읽는 데 실패했습니다.", "details": str(excel_err)}), 400

        # 4. 우체국 양식 매핑
        # 네이버 엑셀의 실제 컬럼명과 일치하는지 확인 필수!
        mapping = {
            "받는 분": df.get("수취인명", ""),
            "우편번호": df.get("우편번호", ""),
            "주소(시도+시군구+도로명+건물번호)": df.get("기본배송지", ""),
            "상세주소(동, 호수, 洞명칭, 아파트, 건물명 등)": df.get("상세배송지", ""),
            "일반전화(02-1234-5678)": df.get("수취인연락처2", ""),
            "휴대전화(010-1234-5678)": df.get("수취인연락처1", ""),
            "중량(kg)": 3,
            "부피(cm)=가로+세로+높이": 80,
            "내용품코드": "농/수/축산물(일반)",
            "내용물": df.get("상품명", ""),
            "배달방식": "일반소포",
            "배송시요청사항": df.get("배송메세지", ""),
            "분할접수 여부(Y/N)": "N"
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
            download_name=f"post_office_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx"
        )

    except Exception as e:
        logging.error(f"Server error: {str(e)}")
        return jsonify({"error": f"서버 오류 발생: {str(e)}"}), 500
