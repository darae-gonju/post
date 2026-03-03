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
            return jsonify({"error": "파일이 업로드되지 않았습니다."}), 400
        
        file = request.files['file']
        password = request.form.get('password', '')
        
        # 1. 파일을 메모리(BytesIO)에 담기
        file_bytes = io.BytesIO(file.read())
        
        # 2. 암호 해제 시도
        decrypted_ptr = io.BytesIO()
        try:
            ms_file = msoffcrypto.OfficeFile(file_bytes)
            ms_file.load_key(password=password) # 비밀번호 입력
            ms_file.decrypt(decrypted_ptr)
            decrypted_ptr.seek(0) # 중요: 데이터를 읽기 위해 포인터를 맨 앞으로 이동
        except Exception as e:
            print(f"Decryption Error: {e}")
            return jsonify({"error": "비밀번호가 틀렸거나 암호화 형식이 맞지 않습니다."}), 403
        
        # 3. 데이터 읽기 및 변환
        try:
            df = pd.read_excel(decrypted_ptr)
        except Exception as e:
            return jsonify({"error": f"엑셀 읽기 실패: {str(e)}"}), 400
        
        # 4. 우체국 양식 매핑 (컬럼명이 다를 경우를 대비해 get 메서드 사용)
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

        # 5. 결과 파일 생성
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
        return jsonify({"error": f"서버 내부 오류: {str(e)}"}), 500

    except Exception as e:
        return jsonify({"error": f"서버 오류: {str(e)}"}), 500

# 여기서 app.run()이나 handler 함수를 절대 넣지 마세요!
