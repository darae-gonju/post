from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import msoffcrypto
import pandas as pd
import io

app = Flask(__name__)
CORS(app)

@app.route('/api', methods=['POST'])
def convert():
    try:
        if 'file' not in request.files:
            return jsonify({"error": "파일이 없습니다."}), 400
        
        file = request.files['file']
        password = request.form.get('password', '') # 입력한 비번 그대로 가져옴
        
        # 1. 파일 데이터를 메모리에 안전하게 로드
        file_content = file.read()
        file_bytes = io.BytesIO(file_content)
        
        # 2. 암호 해제 시도
        decrypted_ptr = io.BytesIO()
        try:
            ms_file = msoffcrypto.OfficeFile(file_bytes)
            # 사용자가 입력한 비밀번호를 '있는 그대로' 사용
            ms_file.load_key(password=password)
            ms_file.decrypt(decrypted_ptr)
            decrypted_ptr.seek(0) # 읽기 포인터 초기화
        except Exception as e:
            # 실패 시 로그를 남기고 사용자에게 알림
            print(f"암호 해제 에러: {str(e)}")
            return jsonify({"error": "비밀번호가 일치하지 않습니다. 다시 확인해주세요."}), 403
        
        # 3. 데이터 읽기 및 변환
        # engine='openpyxl'을 명시하여 최신 엑셀 형식 지원
        df = pd.read_excel(decrypted_ptr, engine='openpyxl')
        
        # 4. 우체국 양식 매핑
        post_df = pd.DataFrame({
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
        })

        # 5. 결과 파일 생성 (메모리에서 즉시 처리)
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
        print(f"서버 내부 에러: {str(e)}")
        return jsonify({"error": f"서버 오류가 발생했습니다: {str(e)}"}), 500
