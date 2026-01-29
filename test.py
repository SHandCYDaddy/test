import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import easyocr
import numpy as np
from PIL import Image as PILImage
import io

st.title("🚗 번호판 인식 및 엑셀 자동 배치 도구")

# 1. OCR 리더기 설정 (한글/영어)
@st.cache_resource
def load_ocr():
    return easyocr.Reader(['ko', 'en'])

reader = load_ocr()

# 2. 파일 업로드
uploaded_file = st.file_uploader("사진을 선택하세요 (1장)", type=['jpg', 'jpeg', 'png'])

if uploaded_file:
    st.image(uploaded_file, caption="업로드된 사진", use_container_width=True)
    
    if st.button("엑셀 파일 생성 및 번호 인식"):
        try:
            # 3. OCR 번호 추출
            with st.spinner("번호판을 인식하는 중입니다..."):
                img = PILImage.open(uploaded_file)
                result = reader.readtext(np.array(img), detail=0)
                detected_text = "".join(result).replace(" ", "") # 공백 제거

            # 4. 엑셀 서식 파일 불러오기
            # 파일명을 'test.xlsm'으로 사용합니다.
            wb = load_workbook("test.xlsm", keep_vba=True)
            ws = wb.active # 혹은 ws = wb["시트이름"]

            # 5. 데이터 입력
            # - 추출된 번호를 A38에 입력
            ws['A38'] = detected_text
            
            # - 사진을 A3:H36 영역의 시작점인 A3에 삽입
            # (VBA가 실행되면 A3:H36 영역에 맞춰 꽉 채워질 것입니다)
            img_for_excel = Image(uploaded_file)
            img_for_excel.anchor = 'A3' 
            ws.add_image(img_for_excel)

            # 6. 결과 저장 및 다운로드
            output = io.BytesIO()
            wb.save(output)
            
            st.success(f"인식 완료: {detected_text}")
            st.download_button(
                label="📥 결과 엑셀 다운로드",
                data=output.getvalue(),
                file_name="result_final.xlsm",
                mime="application/vnd.ms-excel.sheet.macroEnabled.12"
            )
            st.info("💡 엑셀을 연 후, 미리 넣어둔 VBA 매크로를 실행하면 사진이 A3:H36 영역에 꽉 채워집니다.")

        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
            st.info("GitHub 저장소에 'test.xlsm' 파일이 있는지 확인해주세요.")