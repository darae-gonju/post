from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import msoffcrypto
import pandas as pd
import io

app = Flask(__name__)
CORS(app)

# Vercel은 이 'app' 변수를 자동으로 찾아 서버를 구동합니다.

@app.route('/api', methods=['POST'])
def convert():
    try:
        if 'file' not in request.files:
            return jsonify({"error": "파일이 없습니다."}), 400
        
        file = request.files['file']
        password = request.form.get('password', '')
        
        # 1. 암호 해제
        decrypted_ptr = io.BytesIO()
        try:
            ms_file = msoffcrypto.OfficeFile(file)
            ms_file.load_key(password=password)
            ms_file.decrypt(decrypted_ptr)
        except:
            return jsonify({"error": "비밀번호가 틀렸습니다."}), 403
        
        # 2. 데이터 읽기 및 변환
        df = pd.read_excel(decrypted_ptr)
        
        # 3. 우체국 양식 매핑
        post_df = pd.DataFrame({
            "받는 분": df["수취인명"],
            "우편번호": df["우편번호"],
            "주소(시도+시군구+도로명+건물번호)": df["기본배송지"],
            "상세주소(동, 호수, 洞명칭, 아파트, 건물명 등)": df["상세배송지"],
            "일반전화(02-1234-5678)": df.get("수취인연락처2", ""),
            "휴대전화(010-1234-5678)": df.get("수취인연락처1", ""),
            "중량(kg)": 3,
            "부피(cm)=가로+세로+높이": 80,
            "내용품코드": "농/수/축산물(일반)",
            "내용물": df["상품명"],
            "배달방식": "일반소포",
            "배송시요청사항": df.get("배송메세지", ""),
            "분할접수 여부(Y/N)": "N"
        })

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            post_df.to_excel(writer, index=False)
        output.seek(0)

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name="post_office_list.xlsx"
        )

    except Exception as e:
        return jsonify({"error": f"서버 오류: {str(e)}"}), 500

# 여기서 app.run()이나 handler 함수를 절대 넣지 마세요!
