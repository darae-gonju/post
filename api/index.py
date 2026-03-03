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

        # 2. 엑셀 로드 (모든 데이터를 일단 문자열로 읽음)
        try:
            df_raw = pd.read_excel(decrypted_buffer, engine='openpyxl', header=None, dtype=str)
        except Exception as e:
            return jsonify({"error": "엑셀 읽기 실패", "details": str(e)}), 400

        # 3. '수취인명' 행 찾기
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

        # 4. 데이터 추출 함수 (빈칸은 공백으로, 모든 값은 문자열로)
        def get_col(target_name):
            target_clean = target_name.replace(" ", "")
            for real_col in df.columns:
                if str(real_col).replace(" ", "") == target_clean:
                    return df[real_col].fillna("").astype(str).tolist()
            return [""] * num_rows

        # 5. [우체국 최신 17개 컬럼] 순서대로 리스트 생성
        # 제목 없이 데이터만 넣기 위해 리스트의 리스트(행 단위)로 재구성합니다.
        final_data = []
        
        names = get_col("수취인명")
        zips = get_col("우편번호")
        addr1 = get_col("기본배송지")
        addr2 = get_col("상세배송지")
        tel1 = get_col("수취인연락처2")
        tel2 = get_col("수취인연락처1")
        items = get_col("상품명")
        msg = get_col("배송메세지")

        for i in range(num_rows):
            row = [
                names[i],       # 0: 받는 분
                zips[i],        # 1: 우편번호
                addr1[i],       # 2: 주소
                addr2[i],       # 3: 상세주소
                tel1[i],        # 4: 일반전화
                tel2[i],        # 5: 휴대전화
                "3",            # 6: 중량
                "80",           # 7: 부피
                "농/수/축산물(일반)", # 8: 내용품코드
                items[i],       # 9: 내용물
                "미신청",       # 10: 배달방식 (대면/비대면/미신청 중 미신청)
                msg[i],         # 11: 배송시요청사항
                "N",            # 12: 분할접수 여부
                "",             # 13: 분할1 중량
                "",             # 14: 분할1 부피
                "",             # 15: 분할2 중량
                ""              # 16: 분할2 부피
            ]
            final_data.append(row)
        
        # 제목 줄 없는 데이터프레임 생성
        post_df = pd.DataFrame(final_data)

        # 6. 파일 생성 (header=False로 1행 제목 제거)
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            # index=False, header=False로 오직 데이터만 출력
            post_df.to_excel(writer, index=False, header=False)
        output_buffer.seek(0)

        return send_file(
            output_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"우체국_업로드용_{pd.Timestamp.now().strftime('%m%d_%H%M')}.xlsx"
        )

    except Exception as e:
        logging.error(f"Error: {str(e)}")
        return jsonify({"error": f"변환 실패: {str(e)}"}), 500
