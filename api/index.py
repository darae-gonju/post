from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import msoffcrypto
import pandas as pd
import io
import logging
import re

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

        # 2. 엑셀 로드 (dtype=str로 모든 0 누락 방지)
        df_raw = pd.read_excel(decrypted_buffer, engine='openpyxl', header=None, dtype=str)

        # 3. 헤더 탐색
        header_row_idx = 0
        for i, row in df_raw.head(20).iterrows():
            row_values = [str(v).replace(" ", "") for v in row.values if pd.notna(v)]
            if "수취인명" in row_values:
                header_row_idx = i
                break
        
        df = df_raw.iloc[header_row_idx + 1:].copy()
        df.columns = df_raw.iloc[header_row_idx]
        df = df.reset_index(drop=True)

        # 4. 데이터 정제 함수 (특수문자 및 공백 정리)
        def clean_val(target_name):
            target_clean = target_name.replace(" ", "")
            for real_col in df.columns:
                if str(real_col).replace(" ", "") == target_clean:
                    # NaN 처리 후 양 끝 공백 제거, 문자열 강제 변환
                    return df[real_col].fillna("").apply(lambda x: str(x).strip()).tolist()
            return [""] * len(df)

        num_rows = len(df)
        names = clean_val("수취인명")
        zips = clean_val("우편번호")
        addr1 = clean_val("기본배송지")
        addr2 = clean_val("상세배송지")
        tel1 = clean_val("수취인연락처2")
        tel2 = clean_val("수취인연락처1")
        products = clean_val("상품명")
        messages = clean_val("배송메세지")

        # 5. 우체국 17개 컬럼 규격에 맞춰 행 생성
        final_rows = []
        for i in range(num_rows):
            row = [
                names[i],           # 1: 받는 분
                zips[i],            # 2: 우편번호 (0 누락 방지됨)
                addr1[i],           # 3: 주소
                addr2[i],           # 4: 상세주소
                tel1[i],            # 5: 일반전화
                tel2[i],            # 6: 휴대전화
                "3",                # 7: 중량 (표준값)
                "80",               # 8: 부피 (표준값)
                "농/수/축산물(일반)",     # 9: 내용품코드
                products[i][:20],   # 10: 내용물 (글자수 제한 대비 20자로 자름)
                "미신청",           # 11: 배달방식 (미신청)
                messages[i],        # 12: 배송시요청사항
                "N",                # 13: 분할접수 여부
                "", "", "", ""      # 14~17: 분할접수 관련 빈칸
            ]
            final_rows.append(row)

        post_df = pd.DataFrame(final_rows)

        # 6. 파일 생성 (우체국 표준에 맞춘 텍스트 전용 저장)
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            # 제목(header) 없이 데이터만 출력
            post_df.to_excel(writer, index=False, header=False)
            
            # 모든 셀을 '텍스트' 형식으로 강제 지정하기 위한 워크시트 접근
            worksheet = writer.sheets['Sheet1']
            for i in range(len(final_rows)):
                for j in range(17):
                    worksheet.cell(row=i+1, column=j+1).number_format = '@'

        output_buffer.seek(0)

        return send_file(
            output_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"우체국_업로드용_{pd.Timestamp.now().strftime('%m%d_%H%M')}.xlsx"
        )

    except Exception as e:
        logging.error(f"Error: {str(e)}")
        return jsonify({"error": "데이터 변환 중 오류가 발생했습니다."}), 500
