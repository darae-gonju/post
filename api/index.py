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
            return jsonify({"error": "비밀번호가 일치하지 않습니다."}), 403

        # 2. 엑셀 로드 (모든 데이터 문자열 처리)
        df_raw = pd.read_excel(decrypted_buffer, engine='openpyxl', header=None, dtype=str)

        # 3. '수취인명' 기반 데이터 시작점 찾기
        header_row_idx = 0
        for i, row in df_raw.head(20).iterrows():
            row_values = [str(v).replace(" ", "") for v in row.values if pd.notna(v)]
            if "수취인명" in row_values:
                header_row_idx = i
                break
        
        df = df_raw.iloc[header_row_idx + 1:].copy()
        df.columns = df_raw.iloc[header_row_idx]
        df = df.reset_index(drop=True)

        num_rows = len(df)
        if num_rows == 0:
            return jsonify({"error": "변환할 데이터가 없습니다."}), 400

        # 4. 데이터 정제 함수 (빈칸 처리 및 공백 제거)
        def get_val(target_name):
            target_clean = target_name.replace(" ", "")
            for real_col in df.columns:
                if str(real_col).replace(" ", "") == target_clean:
                    return df[real_col].fillna("").apply(lambda x: str(x).strip())
            return pd.Series([""] * num_rows)

        # 5. [우체국 표준 17개 컬럼 매핑]
        # 상세주소가 비어 있으면 '상세주소 없음' 또는 ' ' (공백)을 넣어 필수값 에러 방지
        addr_detail = get_val("상세배송지")
        addr_detail = addr_detail.replace("", " ") # 빈 칸을 공백 한 칸으로 채움

        mapping_data = {
            "받는 분": get_val("수취인명"),
            "우편번호": get_col_safe(df, "우편번호", num_rows),
            "주소(시도+시군구+도로명+건물번호)": get_val("기본배송지"),
            "상세주소(동, 호수, 洞명칭, 아파트, 건물명 등)": addr_detail,
            "일반전화(02-1234-5678)": get_val("수취인연락처2"),
            "휴대전화(010-1234-5678)": get_val("수취인연락처1"),
            "중량(kg)": ["3"] * num_rows,
            "부피(cm)=가로+세로+높이": ["80"] * num_rows,
            "내용품코드": ["농/수/축산물(일반)"] * num_rows,
            "내용물": get_val("상품명").apply(lambda x: x[:20]), # 글자수 제한 방어
            "배달방식": ["미신청"] * num_rows,
            "배송시요청사항": get_val("배송메세지"),
            "분할접수 여부(Y/N)": ["N"] * num_rows,
            "분할접수 첫번째 중량(kg)": [""] * num_rows,
            "분할접수 첫번째 부피(cm)": [""] * num_rows,
            "분할접수 두번째 중량(kg)": [""] * num_rows,
            "분할접수 두번째 부피(cm)": [""] * num_rows
        }

        post_df = pd.DataFrame(mapping_data)

        # 6. 표준 엑셀 파일 생성
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            # 우체국 지침: "표준 엑셀 양식을 사용하십시오" -> 제목 줄 포함(header=True)
            post_df.to_excel(writer, index=False, header=True, sheet_name='Sheet1')
            
            # 모든 셀 서식을 '텍스트'로 지정하여 우편번호/전화번호 깨짐 방지
            worksheet = writer.sheets['Sheet1']
            for row in range(2, num_rows + 2):
                for col in range(1, 18):
                    worksheet.cell(row=row, column=col).number_format = '@'

        output_buffer.seek(0)

        return send_file(
            output_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"우체국_표준양식_{pd.Timestamp.now().strftime('%m%d')}.xlsx"
        )

    except Exception as e
