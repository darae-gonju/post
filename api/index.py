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
        if 'file' not in request.files:
            return jsonify({"error": "파일이 없습니다."}), 400
        
        file = request.files['file']
        password = request.form.get('password', '').strip()
        
        if not password:
            return jsonify({"error": "비밀번호를 입력해주세요."}), 400

        file_data = file.read()
        input_buffer = io.BytesIO(file_data)
        decrypted_buffer = io.BytesIO()

        # 1. 암호 해제
        try:
            office_file = msoffcrypto.OfficeFile(input_buffer)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted_buffer)
            decrypted_buffer.seek(0)
        except Exception:
            return jsonify({"error": "비밀번호가 일치하지 않습니다."}), 403

        # 2. 엑셀 로드
        try:
            df_raw = pd.read_excel(decrypted_buffer, engine='openpyxl', header=None, dtype=str)
        except Exception as e:
            return jsonify({"error": f"엑셀 읽기 실패: {str(e)}"}), 400

        # 3. 네이버 엑셀 헤더 탐색
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

        # 4. 상세주소 필수값 보정 (공백 한 칸이라도 넣어야 에러가 안 납니다)
        addr_detail = get_col_safe(df, "상세배송지", num_rows)
        addr_detail = addr_detail.apply(lambda x: " " if not x or x.strip() == "" else x)

        # 5. [중요] 보내주신 템플릿과 100% 동일한 17개 컬럼 및 제목 설정
        mapping = {
            "받는 분": get_col_safe(df, "수취인명", num_rows),
            "우편번호": get_col_safe(df, "우편번호", num_rows),
            "주소(시도+시군구+도로명+건물번호)": get_col_safe(df, "기본배송지", num_rows),
            "상세주소(동, 호수, 洞명칭, 아파트, 건물명 등)": addr_detail,
            "일반전화(02-1234-5678)": get_col_safe(df, "수취인연락처2", num_rows),
            "휴대전화(010-1234-5678)": get_col_safe(df, "수취인연락처1", num_rows),
            "중량(kg)": ["3"] * num_rows,
            "부피(cm)=가로+세로+높이": ["80"] * num_rows,
            "내용품코드": ["농/수/축산물(일반)"] * num_rows,
            "내용물": get_col_safe(df, "상품명", num_rows).str.slice(0, 20),
            "배달방식": ["미신청"] * num_rows,
            "배송시요청사항": get_col_safe(df, "배송메세지", num_rows),
            "분할접수 여부(Y/N)": ["N"] * num_rows,
            "분할접수 첫번째 중량(kg)": [""] * num_rows,
            "분할접수 첫번째 부피(cm)": [""] * num_rows,
            "분할접수 두번째 중량(kg)": [""] * num_rows,
            "분할접수 두번째 부피(cm)": [""] * num_rows
        }

        final_df = pd.DataFrame(mapping)

        # 6. 엑셀 파일 생성
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            # 템플릿 그대로 제목 줄 포함(header=True)
            final_df.to_excel(writer, index=False, header=True, sheet_name='Sheet1')
            
            # 모든 셀을 텍스트 형식으로 고정
            ws = writer.sheets['Sheet1']
            for r in range(1, num_rows + 2):
                for c in range(1, 18):
                    ws.cell(row=r, column=c).number_format = '@'
                    
        output_buffer.seek(0)

        return send_file(
            output_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"우체국_양식일치_{pd.Timestamp.now().strftime('%m%d')}.xlsx"
        )

    except Exception as e:
        logging.exception("Conversion failed")
        return jsonify({"error": f"서버 오류: {str(e)}"}), 500
