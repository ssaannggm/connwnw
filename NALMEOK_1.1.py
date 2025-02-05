# %%
#pyinstaller --onefile --noconsole --icon=myicon2.ico NALMEOK_1.0.py
import sys
import sqlite3
import pandas as pd
import unicodedata
from PyQt5.QtWidgets import (
    QApplication, QVBoxLayout, QLineEdit, QTableWidget, QTableWidgetItem,QFrame, QListWidget, QSizePolicy,
    QComboBox, QWidget, QTextEdit, QDialog, QLabel, QPushButton, QHBoxLayout, QFileDialog, QMainWindow, QMessageBox, 
)
from PyQt5.QtCore import Qt
from pyhwpx import Hwp
import fitz  # pymupdf
import os
import re
import datetime

#region 전역 변수 선언
app = None  # PyQt5 애플리케이션 객체
conn = None  # SQLite 데이터베이스 연결 객체
cursor = None  # SQLite 커서 객체
table_selector = None  # 테이블 선택 드롭다운
table_widget = None  # 테이블 데이터를 표시하는 위젯
file_list_widget = None
id_input = None  # 사용자 입력 필드
fetch_button = None  # 데이터 가져오기 버튼
paste_button = None  # 복붙 내용 삽입 버튼
hwpfile_button= None
reset_button= None
data1_button= None
data2_button= None
temp_button= None
selected_table = None  # 선택된 테이블 이름
hwpfile_label = None  # 라벨을 전역 변수로 선언
selected_ids = []  # 사용자가 입력한 ID 목록
custom_data = []  # 외부 함수에서 가져온 데이터 저장

cursor_positions = []  # 문단 앞 커서 위치를 담는 리스트 (list, para, pos)
section_titles = []    # 각 위치에 대응하는 목차 또는 섹션 이름 **목차명=필드명이되도록할것
정렬된_cursor_positions=[]
정렬된_section_titles=[]
추출한_섹션명_리스트=[]

안전계획서_리스트=[]
pdf_files =[]
xlsx_files =[]
hwp = None
hwp11 = None

안전계획서1_dic ={
    "비고" : "",
    "파일이름" : "",
    "현장명" : "",
    "현장소재지" : "",
    "공사금액" : 0.0,
    "보고서_날짜_년" : "",
    "보고서_날짜_월" : "",
    "시공자" : "",
    "설계자" : "",
    "공사개요_대상공사" : "",
    "공사개요_구조" : "",
    "공사개요_개소" : 0,
    "공사개요_층수지하" : 0,
    "공사개요_층수지상" : 0,
    "공사개요_굴착깊이" : 0.0,
    "공사개요_최고높이" : 0.0,
    "공사개요_연면적" : 0.0,
    "기타특수구조물개요" : "",
    "주요공법1" : "",
    "주요공법2" : "",
    "주요공법3" : "",
    "주요공법4" : "",
    "주요공법5" : "",
    "주요공법6" : "",
    "주요공법7" : "",
    "주요공법8" : "",
    "주요공법9" : "",
    "주요공법10" : "",
    "파일경로" : "",
}
안전계획서2_1_dic={
    "비고" : "",
    "파일이름" : "",
    "파일경로" : "",
}
안전계획서2_2_dic={
    "비고" : "",
    "파일이름" : "",
    "파일경로" : "",
}
안전계획서2_3_dic={
    "비고" : "",
    "파일이름" : "",
    "파일경로" : "",
}
안전계획서2_4_dic={
    "비고" : "",
    "파일이름" : "",
    "파일경로" : "",
}
def 딕셔너리_데이터_초기화():
    """
    공통 데이터를 초기화합니다.
    """
    for key in 안전계획서1_dic.keys():
        안전계획서1_dic[key] = "" if isinstance(안전계획서1_dic[key], str) else 0 if isinstance(안전계획서1_dic[key], int) else 0.0
    for key in 안전계획서2_1_dic.keys():
        안전계획서2_1_dic[key] = "" if isinstance(안전계획서2_1_dic[key], str) else 0 if isinstance(안전계획서2_1_dic[key], int) else 0.0
    for key in 안전계획서2_2_dic.keys():
        안전계획서2_2_dic[key] = "" if isinstance(안전계획서2_2_dic[key], str) else 0 if isinstance(안전계획서2_2_dic[key], int) else 0.0
    for key in 안전계획서2_3_dic.keys():
        안전계획서2_3_dic[key] = "" if isinstance(안전계획서2_3_dic[key], str) else 0 if isinstance(안전계획서2_3_dic[key], int) else 0.0
    for key in 안전계획서2_4_dic.keys():
        안전계획서2_4_dic[key] = "" if isinstance(안전계획서2_4_dic[key], str) else 0 if isinstance(안전계획서2_4_dic[key], int) else 0.0
def 커서_섹션명_커스텀_리스트_데이터_초기화():
    global cursor_positions
    global section_titles
    global custom_data 
    global 정렬된_cursor_positions
    global 정렬된_section_titles
    global 추출한_섹션명_리스트
    cursor_positions = []  # 문단 앞 커서 위치를 담는 리스트 (list, para, pos)
    section_titles = []    # 각 위치에 대응하는 목차 또는 섹션 이름 **목차명=필드명이되도록할것
    custom_data = []
    정렬된_cursor_positions=[]
    정렬된_section_titles=[]
    추출한_섹션명_리스트=[]

main_pdf_path = None
sub_pdf_dir = None
output_pdf_path = "PDF끼워넣기_초안"
search_text = "후면 첨부참조"
sub_doc = None
main_doc = None
insertions = []
sub_pdf_names = []  # PDF 파일 이름 (확장자 없는 형태)
sub_pdf_paths = []  # PDF 파일의 실제 경로
#endregion

#################################################################

#region db 함수
def connect_db():
    """
    SQLite 데이터베이스에 연결합니다.
    데이터베이스 파일 이름은 기본적으로 'data.db'로 고정됩니다.
    데이터베이스 파일이 없으면 자동으로 생성됩니다.

    :return: 데이터베이스 연결 객체와 커서
    """
    db_name = 'data.db'  # 고정된 데이터베이스 파일 이름
    conn = sqlite3.connect(db_name)
    cursor = conn.cursor()
    return conn, cursor

# 테이블 생성 함수
def initialize_db(table_name, data_dict):
    """
    딕셔너리 키를 기반으로 데이터베이스 테이블을 생성합니다.
    """
    conn = sqlite3.connect('data.db')
    cursor = conn.cursor()
    
    # 딕셔너리에서 열 생성 쿼리 동적 작성
    columns = ', '.join([f"{key} {get_sqlite_type(value)}" for key, value in data_dict.items()])
    fixed_columns = "섹션명 TEXT, 리스트 INTEGER, para INTEGER, pos INTEGER"
    query = f"CREATE TABLE IF NOT EXISTS {table_name} (id INTEGER PRIMARY KEY AUTOINCREMENT, {columns}, {fixed_columns})"
    
    cursor.execute(query)
    conn.commit()
    conn.close()
def get_sqlite_type(value):
    """
    딕셔너리 값의 데이터 타입에 따라 SQLite 데이터 타입 반환.
    """
    if isinstance(value, int):
        return "INTEGER"
    elif isinstance(value, float):
        return "REAL"
    elif isinstance(value, str):
        return "TEXT"
    else:
        return "TEXT"  # 기본값

# 데이터 삽입 함수 / 문서 전체를 한방에 db화함.
def insert_data(
    table_name, data_dict, 섹션명_리스트= 정렬된_section_titles, 데이터_리스트=정렬된_cursor_positions
):
    """
    딕셔너리와 리스트를 사용해 데이터베이스에 데이터를 삽입합니다.
    
    :param db_name: 데이터베이스 파일 이름
    :param table_name: 테이블 이름
    :param data_dict: 삽입할 공통 데이터 딕셔너리
    :param 섹션명_리스트: 섹션명을 담은 리스트
    :param 데이터_리스트: (리스트, para, pos) 튜플 리스트
    """
    if len(섹션명_리스트) != len(데이터_리스트):
        raise ValueError("섹션명_리스트와 데이터_리스트의 길이가 다릅니다.")
    
    conn = sqlite3.connect('data.db')
    cursor = conn.cursor()

    # 공통 데이터 열과 값
    common_columns = ', '.join(data_dict.keys())
    common_values = tuple(data_dict.values())

    # 섹션명_리스트와 데이터_리스트를 조합하여 삽입
    for 섹션명, (리스트, para, pos) in zip(섹션명_리스트, 데이터_리스트):
        query = f'''
            INSERT INTO {table_name} (
                {common_columns}, 섹션명, 리스트, para, pos
            ) VALUES ({', '.join(['?'] * (len(data_dict) + 4))})
        '''
        values = common_values + (섹션명, 리스트, para, pos)
        cursor.execute(query, values)

    conn.commit()
    conn.close()
    print(f"{table_name} 테이블에 데이터가 삽입되었습니다.")

# 모든 데이터 조회 함수
def print_table_data(table_name):
    conn = sqlite3.connect('data.db')
    cursor = conn.cursor() 

    cursor.execute(f"SELECT * FROM {table_name}")
    rows = cursor.fetchall()
    
    for row in rows:
        print(row)
    
    conn.close()

#custom_data로 특정id만 추출하기(id의 위치),(id+1의 위치)(id의 섹션명)
def fetch_custom_data_with_next_positions(table_name, id_list):
    """
    주어진 테이블 이름과 ID 리스트에 해당하는 데이터를 custom_data 형식으로 반환.
    (id의 시작 좌표), (id+1의 시작 좌표), "id 섹션명" 구조로 변환.

    :param table_name: 데이터베이스 테이블 이름
    :param id_list: ID 리스트
    :return: custom_data 리스트
    """
    conn = sqlite3.connect('data.db')
    cursor = conn.cursor()
    
    # ID 리스트를 조건으로 데이터 가져오기
    placeholders = ', '.join('?' for _ in id_list)
    query = f'''
        SELECT id, 리스트, para, pos, 섹션명, 파일경로 
        FROM {table_name}
        WHERE id IN ({placeholders}) 
        ORDER BY id
    '''
    cursor.execute(query, id_list)
    rows = cursor.fetchall()
 

    # ID+1 위치 한 번에 가져오기
    next_positions = {}
    for id_ in id_list:
        query_next = f'''
            SELECT 리스트, para, pos 
            FROM {table_name}
            WHERE id = ?
        '''
        cursor.execute(query_next, (id_ + 1,))  # `id + 1`으로 수정
        next_row = cursor.fetchone()
        if next_row:
            next_positions[id_] = next_row
        else:
            next_positions[id_] = "문서끝"

    # 데이터 변환
    custom_data = []
    for row in rows:
        current_id, 리스트, para, pos, 섹션명, 파일경로 = row
        next_position = next_positions.get(current_id, "문서끝")
        custom_data.append(
            ((리스트, para, pos), next_position, 섹션명, 파일경로)
        )
    
    conn.close()
    return custom_data

# 데이터 삭제 함수
def delete_data(table_name, row_id):
    """
    특정 테이블에서 특정 id 값을 가진 데이터를 삭제합니다.

    :param table_name: 테이블 이름
    :param row_id: 삭제할 행의 id 값
    """
    conn, cursor = connect_db() 
    query = f'DELETE FROM {table_name} WHERE id = ?'
    cursor.execute(query, (row_id,))
    conn.commit()
    conn.close()

# 데이터 업데이트 함수  ##이것도 잘모르겠음. 
def update_data(table_name, row_id, **kwargs):
    """
    특정 테이블에서 특정 ID의 데이터를 업데이트합니다.
    
    :param table_name: 테이블 이름
    :param row_id: 업데이트할 행의 ID
    :param kwargs: 업데이트할 열 이름과 값의 딕셔너리 (예: {"파일이름": "새 파일명", "현장명": "새 현장명"})
    """
    conn, cursor = connect_db()
    updates = []
    values = []

    for key, value in kwargs.items():
        updates.append(f"{key} = ?")
        values.append(value)
    
    values.append(row_id)
    query = f'UPDATE {table_name} SET {", ".join(updates)} WHERE id = ?'
    cursor.execute(query, values)
    conn.commit()
    conn.close()

#테이블 삭제
def reset_table(table_name):
    """
    특정 테이블을 삭제하고 초기화합니다.
    테이블 삭제만 처리하며, 생성은 별도의 함수로 처리해야 합니다.

    :param table_name: 초기화할 테이블 이름
    """
    conn, cursor = connect_db()
    
    # 테이블 삭제
    cursor.execute(f'DROP TABLE IF EXISTS {table_name}')
    conn.commit()
    conn.close()
    
    print(f"{table_name} 테이블이 초기화되었습니다.")

#열 목록 확인
def check_table_structure(table_name):
    conn = sqlite3.connect('data.db')
    cursor = conn.cursor()
    cursor.execute(f"PRAGMA table_info({table_name})")
    rows = cursor.fetchall()
    conn.close()
    
    print(f"{table_name} 테이블 구조:")
    for row in rows:
        print(row)
    
# 데이터베이스 초기화
#initialize_db('con_1',안전계획서1_dic)

#endregion

#################################################################

#region hwp 초기화
#hwp = Hwp(new=True)
def 한글1_실행(hwp=hwp): ### todo : 번호 매기기? 등 여러 한글파일을 다뤄야 할 수 있으므로 / clear quit도 고려하기
    """메인 파일은 무조건 hwp"""
    if hwp.FileOpen():
        hwp.MoveDocBegin()
def 한글1_종료(hwp=hwp):
    """내용 버린 후 종료"""
    hwp.clear()

#상용함수
def 오른_표이동(hwp=hwp, n = int ):
    """n번 오른쪽으로 이동"""
    for _ in range(n):
        hwp.TableRightCell()
def 고정폭빈칸삭제(hwp=hwp):
    hwp.HAction.GetDefault("DeleteCtrls", hwp.HParameterSet.HDeleteCtrls.HSet)
    hwp.HParameterSet.HDeleteCtrls.CreateItemArray("DeleteCtrlType", 1)
    hwp.HParameterSet.HDeleteCtrls.DeleteCtrlType.SetItem(0, 7)  # <--- Item을 SetItem으로 고쳤음.
    hwp.HAction.Execute("DeleteCtrls", hwp.HParameterSet.HDeleteCtrls.HSet)
#endregion

#################################################################

#region 문서 공통 info 추출 -> 딕셔너리 저장
def 안전1편_공통info추출(hwp=hwp):
    고정폭빈칸삭제()
    #파일경로 추출
    안전계획서1_dic["파일경로"] = hwp.Path

    #파일이름 추출
    안전계획서1_dic["파일이름"] = os.path.basename(안전계획서1_dic["파일경로"])
    print(안전계획서1_dic["파일경로"], 안전계획서1_dic["파일이름"])

    ##처음 위치로##
    hwp.Cancel()
    hwp.MoveDocBegin()

    # 정규표현식으로 연도와 월 추출
    연월텍스트 = hwp.GetPageText()
    패턴 = r'(\d{4})\.\s*(\d{2})'
    매칭 = re.search(패턴, 연월텍스트)

    if 매칭:
        안전계획서1_dic["보고서_날짜_년"], 안전계획서1_dic["보고서_날짜_월"] = 매칭.groups()
        #print(f"연도: {안전계획서1_dic["보고서_날짜_년"]}")
        #print(f"월: {안전계획서1_dic["보고서_날짜_월"]}")
    else:
        print("연도와 월을 찾을 수 없습니다.")

    #region###--공사개요서 뒤지기--###
    if hwp.find_forward("1. 공사 개요서"):
        hwp.SelectCtrlFront()
        hwp.ShapeObjTableSelCell() # 첫번째 셀 선택
    else : print("공사개요서 못찾음.")

    #현장명
    오른_표이동(2)
    안전계획서1_dic["현장명"] = hwp.get_selected_text()

    #소재지
    오른_표이동(2)
    안전계획서1_dic["현장소재지"] = hwp.get_selected_text()

    #공사금액
    오른_표이동(4)
    안전계획서1_dic["공사금액"] = hwp.get_selected_text()

    #시공자회사명
    오른_표이동(3)
    안전계획서1_dic["시공자"] = hwp.get_selected_text()

    #설계자 회사명
    오른_표이동(25)
    안전계획서1_dic["설계자"] = hwp.get_selected_text()

    #대상공사
    오른_표이동(41)
    안전계획서1_dic["공사개요_대상공사"] = hwp.get_selected_text()

    #구조
    오른_표이동(1)
    안전계획서1_dic["공사개요_구조"] = hwp.get_selected_text()

    #개소
    오른_표이동(1)
    안전계획서1_dic["공사개요_개소"] = hwp.get_selected_text()

    오른_표이동(1)
    안전계획서1_dic["공사개요_층수지하"] = hwp.get_selected_text()

    오른_표이동(1)
    안전계획서1_dic["공사개요_층수지상"] = hwp.get_selected_text()

    오른_표이동(1)
    안전계획서1_dic["공사개요_굴착깊이"] = hwp.get_selected_text()

    오른_표이동(1)
    안전계획서1_dic["공사개요_최고높이"] = hwp.get_selected_text()

    오른_표이동(1)
    안전계획서1_dic["공사개요_연면적"] = hwp.get_selected_text()

    오른_표이동(2)
    안전계획서1_dic["기타특수구조물개요"] = hwp.get_selected_text()

    #공법리스트
    오른_표이동(2)
    주요공법 = hwp.get_selected_text()  # 텍스트 가져오기
    공법리스트 = [공법.strip() for 공법 in 주요공법.split('\n') if 공법.strip()]  # 줄 단위로 나누고 양쪽 공백 제거 및 빈 줄 제거
    print(공법리스트)
    for i in range(1, 11):  # "주요공법1" ~ "주요공법10"
        if i <= len(공법리스트):  # 공법리스트에 값이 남아 있는 경우
            안전계획서1_dic[f"주요공법{i}"] = 공법리스트[i - 1]
        else:  # 공법리스트에 값이 없는 경우 빈 문자열 유지
            안전계획서1_dic[f"주요공법{i}"] = ""
    print(안전계획서1_dic)
    #endregion

    안전계획서1_dic["비고"] = ""  # 비고 초기화

    hwp.Cancel()
def 안전2_1_공통info추출(hwp=hwp):
    고정폭빈칸삭제()
    #파일경로 추출
    안전계획서2_1_dic["파일경로"] = hwp.Path

    #파일이름 추출
    안전계획서2_1_dic["파일이름"] = os.path.basename(안전계획서2_1_dic["파일경로"])

    ##처음 위치로##
    hwp.Cancel()
    hwp.MoveDocBegin()
def 안전2_2_공통info추출(hwp=hwp):
    고정폭빈칸삭제()
    #파일경로 추출
    안전계획서2_2_dic["파일경로"] = hwp.Path

    #파일이름 추출
    안전계획서2_2_dic["파일이름"] = os.path.basename(안전계획서2_2_dic["파일경로"])

    ##처음 위치로##
    hwp.Cancel()
    hwp.MoveDocBegin()
def 안전2_3_공통info추출(hwp=hwp):
    고정폭빈칸삭제()
    #파일경로 추출
    안전계획서2_3_dic["파일경로"] = hwp.Path

    #파일이름 추출
    안전계획서2_3_dic["파일이름"] = os.path.basename(안전계획서2_3_dic["파일경로"])

    ##처음 위치로##
    hwp.Cancel()
    hwp.MoveDocBegin()
def 안전2_4_공통info추출(hwp=hwp):
    고정폭빈칸삭제()
    #파일경로 추출
    안전계획서2_4_dic["파일경로"] = hwp.Path

    #파일이름 추출
    안전계획서2_4_dic["파일이름"] = os.path.basename(안전계획서2_4_dic["파일경로"])

    ##처음 위치로##
    hwp.Cancel()
    hwp.MoveDocBegin()
#endregion

#################################################################

#region 한글 함수 : 섹션명추출 temp로 저장후 표삭제,빈칸삭제,공백삭제 후 텍스트 스캔 후 리스트로 반환
def 섹션명_추출(hwp=hwp, save_temp_path="템플릿\\Temp.hwp"):
    """
    1. temp로 저장 후 새 객체로 오픈
    2. 모든 표 삭제
    3. 모든 고정폭 빈칸 삭제
    4. 공백 정리 
    5. 강제쪽나눔 삭제
    6. 문자열 스캔
    7. return 리스트
    :param hwp: 한글(HWP) 객체
    :param save_temp_path: 임시 저장 경로 (기본값: "템플릿\\Temp.hwp")

    """
    # 1. 템플릿 파일 저장
    hwp.SaveAs(save_temp_path, arg="lock:false")
    hwp3 = Hwp(new= True, visible=True)
    hwp3.open(save_temp_path)
    # 2. 모든 컨트롤(표, 사각형) 삭제
    for ctrl in reversed(hwp3.ctrl_list):
        if ctrl.UserDesc == "표":  # 컨트롤이 표일 경우
            hwp3.delete_ctrl(ctrl)  # 바로 삭제
        if ctrl.UserDesc == "사각형":  # 컨트롤이 표일 경우
            hwp3.delete_ctrl(ctrl)  # 바로 삭제
    # 3. 고정폭빈칸삭제
    hwp3.HAction.GetDefault("DeleteCtrls", hwp3.HParameterSet.HDeleteCtrls.HSet)
    hwp3.HParameterSet.HDeleteCtrls.CreateItemArray("DeleteCtrlType", 1)
    hwp3.HParameterSet.HDeleteCtrls.DeleteCtrlType.SetItem(0, 7)  # <--- Item을 SetItem으로 고쳤음.
    hwp3.HAction.Execute("DeleteCtrls", hwp3.HParameterSet.HDeleteCtrls.HSet)
    # 4. 문서의 공백 및 불필요한 내용 정리
    hwp3.find_replace_all(src="  ", dst="")#두칸 띄어쓰기 삭제
    hwp3.MoveDocBegin()
    while hwp3.MoveSelRight():
        selected_text = hwp3.get_selected_text()
        if not selected_text.strip():  # 공백 문자열 삭제
            hwp3.Delete()
        elif selected_text in ["-", "※", "▸", "▣","","<" ]:  # todo 특정 문자일 경우 삭제해버려서 섹션명을 추출하지 않는 것은 어떤가?
            hwp3.MoveSelParaEnd()
            hwp3.Delete()
        elif hwp3.MoveNextParaBegin():
            continue
        else:
            break
    
    # 5. 강제쪽나눔 전체 삭제
    def delete_forced_page_breaks():
        hwp3.SetMessageBoxMode(0x00020000)
        pset = hwp3.HParameterSet.HGotoE
        hwp3.HAction.GetDefault("Goto", pset.HSet)
        while True:
            try:
                pset.HSet.SetItem("DialogResult", 54)  # 강제쪽나눔으로 이동
                pset.SetSelectionIndex = 5
                if not hwp3.HAction.Execute("Goto", pset.HSet):  # 이동 실패 시 종료
                    break
                hwp3.DeleteBack()  # 강제쪽나눔 삭제
            except Exception as e:
                print(f"오류 발생: {e}")
                break
    hwp3.MoveDocBegin()
    delete_forced_page_breaks()
    

    # 6. 텍스트 스캔 및 섹션 타이틀 추출
    hwp3.MoveDocBegin()
    hwp3.init_scan()
    extracted_texts = []  # 추출된 텍스트 저장
    while True:
        state, text = hwp3.get_text()
        if text and text.strip():  # 공백 제외
            clean_text = text.replace("\r\n", "").replace("\n", "").replace("\r", "")
            print(clean_text)  # 정리된 텍스트 출력
            extracted_texts.append(clean_text)
        if state <= 1:  # 종료 조건
            break
    hwp3.release_scan()
    hwp3.save_as(save_temp_path)
    hwp3.clear()
    hwp3.quit()
    # 7. 결과 반환 (추출된 텍스트 리스트)
    return extracted_texts

# hwp 객체를 가져와서 함수 호출 사용법
####추출한_섹션명_리스트 = 섹션명_추출(hwp)

#endregion

#################################################################

#region 한글 함수 : 문서 포지션 추출 -> 위치 순서대로 정렬
def 처음지점추가(hwp0 = hwp):
    hwp0.MoveDocBegin()
    cursor_positions.append(hwp0.GetPos())
    section_titles.append("문서처음")
def 문단시작지점추가(hwp0 = hwp):
        문단시작_list =[]
        for i in hwp0.ctrl_list:
            if i.UserDesc == "새 번호":#컨트롤이 표일 경우
                문단시작_list.append(i)#리스트에 저장
        if 문단시작_list:
            for i in range(len(문단시작_list)):
                hwp0.move_to_ctrl(문단시작_list[i])
                hwp0.MoveLeft()
                cursor_positions.append(hwp0.GetPos())
                section_titles.append(f"문단{i+1}")
def 마지막지점추가(hwp0 = hwp):
    hwp0.MoveDocEnd()
    cursor_positions.append(hwp0.GetPos())
    section_titles.append("문서끝")

def 위치추가(hwp = hwp):
    # 목차 문단 맨 앞의 위치 가져오기
    pos = hwp.GetPos()  # pos는 (리스트, para, pos) 형태의 튜플
    
    # 리스트 값이 0인지 확인
    if pos[0] != 0:  # 리스트 값이 0이 아니면 함수 종료
        print(f"리스트 값이 0이 아니므로 추가하지 않음: {pos}")
        #다시 찾기 추가해야함.
        return
    cursor_positions.append(hwp.GetPos())

    #목차 문단 셀선택 후 추가하고 다시 첫 위치로 돌아가기
    hwp.MoveSelParaEnd()
    section_titles.append(hwp.get_selected_text())
    hwp.MoveParaBegin()
def 중간위치추가(hwp0 = hwp, sec_list=추출한_섹션명_리스트):
    """추출한 섹션명 리스트를 넣으면 리스트를 돌면서
        find후 포지션을 cursor_positions, section_titles에 추가(표안은 추가안함)

    """
    hwp0.MoveDocBegin()###이걸 왜 실패하고있지???todotodotodotodo
    for section in sec_list:
        while hwp0.find(section, 'AllDoc', WholeWordOnly=1, SeveralWords=0, UseWildCards=0):
            # 현재 위치의 커서 정보를 가져옴
            hwp0.MoveLeft()  # 커서를 찾은 위치로 보정
            hwp0.MoveRight()
            pos = hwp0.GetPos()

            if pos[0] != 0:  # 리스트 값이 0이 아니면 다음 위치로 검색
                print(f"리스트 값이 0이 아니므로 다음 위치 검색: {pos}")
                hwp0.MoveRight()  # 커서를 한 칸 오른쪽으로 이동하여 다음 검색 준비
                continue  # 다음 위치 검색
            else:
                # 커서 위치가 유효하다면, 데이터를 추가
                cursor_positions.append(pos)
                hwp0.MoveSelParaEnd()  # 현재 커서가 위치한 문단의 끝으로 이동
                section_titles.append(hwp0.get_selected_text())
                print(f'{section} : 완료')
                hwp0.MoveDocBegin()
                break  # 섹션 검색 완료 후 다음 섹션으로 이동
        else:
            print(f"{section} : 실패")  # 검색 실패 처리
###중간위치추가(추출한_섹션명_리스트)

#리스트 para기준 재정렬
def 재정렬_파라기준(cursor_positions, section_titles):
    """
    cursor_positions의 'para' 값을 기준으로 cursor_positions와 section_titles를 재정렬하는 함수.
    
    매개변수:
    - cursor_positions (list of tuples): (table, para, dotPos) 형식의 커서 위치 리스트.
    - section_titles (list of str): 해당 커서 위치에 대응하는 섹션 제목 리스트.
    
    반환값:
    - tuple: 정렬된 cursor_positions와 section_titles.
    """
    # cursor_positions와 section_titles를 결합
    결합_데이터 = list(zip(cursor_positions, section_titles))
    
    # 'para' 값(튜플의 두 번째 값)을 기준으로 정렬
    정렬된_데이터 = sorted(결합_데이터, key=lambda x: x[0][1])
    
    # 정렬된 데이터를 다시 두 개의 리스트로 분리
    정렬된_cursor_positions, 정렬된_section_titles = zip(*정렬된_데이터)
    
    return list(정렬된_cursor_positions), list(정렬된_section_titles)
###정렬된_cursor_positions, 정렬된_section_titles = 재정렬_파라기준(cursor_positions, section_titles)

#endregion

#################################################################
 
#region 한글 함수 : 복붙하기 수정필요함!!

#내용 삽입 함수
def 복붙_내용삽입(custom_data):
    """
    custom_data: [(섹션시작좌표, 섹션끝좌표, 섹션명), ...] 형태의 리스트
    붙혀넣을 템플릿의 필드명은 '섹션명'이어야 한다.
    """
    global selected_table
    hwp2 = Hwp(new=True)
    hwp2.FileOpen()##todo 타이틀에 따라 템플릿 구성되게
    hwp5 = Hwp(visible=False)
    for section_start, section_end, section_name, 파일경로 in custom_data:
        # 섹션 시작 위치로 이동
        if hwp5.Path == 파일경로 :
            pass
        else:

            hwp5.clear()
            hwp5.open(파일경로)
            
        hwp5.SetPos(*section_start)  
        hwp5.MoveNextParaBegin()  # 다음 단락으로 이동
        hwp5.Select()  # 섹션 시작 위치에서 선택 시작
        # 섹션 끝 위치로 이동
        if section_end == None:
            print("잘못된 id를 추출하였습니다.(문서끝id입력함)")
        else:
            hwp5.SetPos(*section_end)
            hwp5.MoveLeft()  # 끝 위치에서 한 글자 왼쪽으로 이동해 선택 범위 조정

        # 복사 작업 수행
        hwp5.Copy()

        # 두 번째 HWP 파일의 필드로 이동하여 붙여넣기
        
        hwp2.MoveToField(section_name)
        hwp2.Paste()  # 복사한 내용 붙여넣기
    hwp5.clear()
    hwp5.quit()
    hwp2.save_as(f"{selected_table}.hwp")
    hwp2.clear(
    hwp2.quit()
    )
##복붙_내용삽입(custom_data)
#endregion

#################################################################

#region 데이터, db테이블 버튼
#카피가능 메시지박스
def show_copyable_message(title, message, info):
    """
    사용자가 내용을 복사할 수 있는 팝업 메시지 창
    """
    dialog = QDialog()
    dialog.setWindowTitle(title)
    dialog.setGeometry(300, 300, 500, 300)  # 창 크기 설정

    layout = QVBoxLayout()

    text_box = QTextEdit()
    text_box.setPlainText(f"{message}\n\n{info}")  # 메시지 설정
    text_box.setReadOnly(True)  # 편집 불가 (읽기 전용)
    text_box.setTextInteractionFlags(Qt.TextSelectableByMouse | Qt.TextSelectableByKeyboard)  # 마우스 및 키보드 선택 허용

    close_button = QPushButton("닫기")
    close_button.clicked.connect(dialog.close)

    layout.addWidget(text_box)
    layout.addWidget(close_button)
    dialog.setLayout(layout)

    dialog.exec_()  # 다이얼로그 실행

# 메시지 박스
def show_message(title, message, info=None, icon=QMessageBox.Information):
    """
    PyQt5 메시지 박스를 표시하는 함수.
    :param title: 메시지 박스 제목
    :param message: 메시지 내용
    :param info: 추가 정보 (선택 사항)
    :param icon: 메시지 박스 아이콘 (기본값: Information)
    """
    msg = QMessageBox()
    msg.setIcon(icon)  # 메시지 아이콘 설정
    msg.setWindowTitle(title)  # 창 제목
    msg.setText(message)  # 메시지 내용 설정
    if info:
        msg.setInformativeText(info)  # 추가 정보 (선택 사항)
    msg.exec_()  # 메시지 박스 실행

# DB 연결 초기화
def init_db(db_path="data.db"):
    """
    SQLite 데이터베이스에 연결합니다.
    :param db_path: 데이터베이스 파일 경로
    """
    global conn, cursor
    conn = sqlite3.connect(db_path)  # 데이터베이스 연결 생성
    cursor = conn.cursor()  # 커서 객체 생성

# 테이블 목록 로드
def load_table_list():
    """
    SQLite 데이터베이스에서 테이블 목록을 가져와 드롭다운에 추가합니다.
    """
    global cursor, table_selector
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")  # 테이블 목록 조회
    tables = [row[0] for row in cursor.fetchall()]  # 결과에서 테이블 이름 추출
    table_selector.addItems(tables)  # 드롭다운에 테이블 목록 추가

# 선택한 테이블 데이터 로드
def load_table_data():
    """
    드롭다운에서 선택한 테이블의 데이터를 QTableWidget에 로드합니다.
    """
    global cursor, table_selector, table_widget,selected_table
    selected_table = table_selector.currentText()  # 선택된 테이블 이름 가져오기
    if not selected_table:
        return

    # 테이블 데이터 조회
    cursor.execute(f"SELECT * FROM {selected_table}")
    rows = cursor.fetchall()  # 테이블의 모든 행 가져오기
    columns = [description[0] for description in cursor.description]  # 컬럼 이름 가져오기

    # QTableWidget에 데이터 삽입
    table_widget.setRowCount(len(rows))
    table_widget.setColumnCount(len(columns))
    table_widget.setHorizontalHeaderLabels(columns)  # 열 제목 설정

    for i, row in enumerate(rows):
        for j, value in enumerate(row):
            table_widget.setItem(i, j, QTableWidgetItem(str(value)))  # 각 셀에 데이터 삽입

    # 초기 숨김 설정 - 특정 열 숨기기
    hidden_columns = ["파일이름", "파일경로", "리스트", "para", "pos"]
    for col_index, col_name in enumerate(columns):
        if col_name in hidden_columns:
            table_widget.setColumnHidden(col_index, True)

    # show_message(
    #     title="테이블 로드 완료",
    #     message=f"'{selected_table}' 테이블의 데이터를 로드했습니다.",
    #     info=f"총 {len(rows)}개의 행이 로드되었습니다."
    # )

# id가져오기 버튼
def fetch_custom_data():
    """
    사용자가 입력한 ID를 처리하고 외부 함수로 데이터를 가져옵니다.
    """
    global id_input, selected_ids, custom_data, paste_button, selected_table

    # 사용자 입력 처리
    input_text = id_input.text()
    try:
        # 쉼표로 구분된 숫자 리스트 생성
        selected_ids = [int(x.strip()) for x in input_text.split(",") if x.strip().isdigit()]
        if not selected_ids:
            raise ValueError("숫자를 입력해야 합니다.")
    except ValueError as e:
        show_message(
            title="입력 오류",
            message="유효한 숫자를 입력해야 합니다.",
            info=str(e),
            icon=QMessageBox.Warning
        )
        return

    show_message(
        title="ID 선택 완료",
        message="선택된 ID가 처리되었습니다.",
        info=f"선택된 ID: {selected_ids}"
    )

    # 외부 함수 호출 - 예제
    table_name = selected_table
    custom_data = fetch_custom_data_with_next_positions(table_name, selected_ids)  # 외부 함수 호출
    if custom_data:
        # custom_data에서 섹션명(튜플의 마지막 요소)만 추출
        section_names = [item[-2] for item in custom_data]  # 각 항목의 마지막 요소 추출
        section_names_str = "\n".join(section_names)  # 줄바꿈으로 섹션명을 연결

        show_message(
            title="데이터 가져오기 완료",
            message="데이터를 성공적으로 가져왔습니다.",
            info=f"가져온 섹션명:\n{section_names_str}"
        )
        paste_button.setEnabled(True)
    else:
        show_message(
            title="데이터 없음",
            message="선택한 ID에 대한 데이터를 찾을 수 없습니다.",
            icon=QMessageBox.Warning
        )

# 내용 입력 버튼
def execute_paste():
    """
    외부 함수를 호출하여 복붙 데이터를 삽입합니다.
    """
    global custom_data, selected_table
    if not custom_data:
        show_message(
            title="삽입 실패",
            message="삽입할 데이터가 없습니다.",
            icon=QMessageBox.Warning
        )
        return
    
    복붙_내용삽입(custom_data)  # 외부 함수 호출
    show_message(
        title="삽입 성공",
        message="데이터를 성공적으로 삽입했습니다.",
        info=f"저장 경로: {selected_table}.hwp"
    )

def 한글파일선택버튼(): 
    #데이터 초기화
    #한글열기
    #path 추출하고 라벨 업데이트
    global hwpfile_label, data1_button, hwpfile_button,reset_button  # 전역 변수 참조
    딕셔너리_데이터_초기화()
    커서_섹션명_커스텀_리스트_데이터_초기화()
    try:
        hwp.FileOpen()  # 파일 열기
        file_path = hwp.Path  # 파일 경로 가져오기
        hwpfile_label.setText(f"path : {file_path}")  # 라벨 업데이트
        data1_button.setEnabled(True)
        reset_button.setEnabled(True)
        hwpfile_button.setEnabled(False)
        print(f"파일 열기: {file_path}")  # 디버깅용 출력
        show_message(
            title="한글 파일 연결",
            message="한글 파일 연결 성공. / [데이터추출] 가능",
            info=f"연결 경로:\n{file_path}"
        )
    except Exception as e:
        show_message(
            title="오류 발생",
            message="오류 발생!",
            info=f"발생 오류 :\n{e}"
        )
        
def 초기화버튼():
    global hwpfile_label, data1_button, data2_button, temp_button
    try:   
        hwp.Clear()  # 파일 닫기
        딕셔너리_데이터_초기화()
        커서_섹션명_커스텀_리스트_데이터_초기화()
        hwpfile_button.setEnabled(True)
        data1_button.setEnabled(False)
        data2_button.setEnabled(False)
        hwpfile_label.setText("path : ")  # 라벨 초기화
        show_message(
            title="연결 해제",
            message="연결 해제",
        )
    except Exception as e:
        show_message(
            title="오류 발생",
            message="오류 발생!",
            info=f"발생 오류 :\n{e}"
        )
def 데이터추출버튼():
#     #현재 선택된 테이블 가져오기
#     #테이블에 맞춰 인포 추출
#     #테이블 별 섹션명 (미리만들기/매번만들기)
#     #포지션 추출 후 확인
    global selected_table, 정렬된_cursor_positions, 정렬된_section_titles, data2_button, temp_button
    if selected_table == '안전관리계획서1':
        안전1편_공통info추출()
    elif selected_table =='안전관리계획서2_1':
        안전2_1_공통info추출()
    elif selected_table =='안전관리계획서2_2':
        안전2_2_공통info추출()
    elif selected_table =='안전관리계획서2_3':
        안전2_3_공통info추출()
    elif selected_table =='안전관리계획서2_4':
        안전2_4_공통info추출()
    try:
        show_message(
                title="데이터 추출중",
                message="데이터를 추출하고 있습니다.(시간 소요)\n한글 '찾기' 경고가 나올때까지 기다리세요.\n완료후 [데이터입력],[템플릿 제작] 사용가능 ",
                info=f"선택한 테이블 :\n{selected_table}"
            )
        추출한_섹션명_리스트.clear()
        추출한_섹션명_리스트 = 섹션명_추출(hwp)
        처음지점추가()
        문단시작지점추가()
        마지막지점추가()
        중간위치추가(추출한_섹션명_리스트)
        정렬된_cursor_positions.clear() 
        정렬된_section_titles.clear()
        정렬된_cursor_positions, 정렬된_section_titles = 재정렬_파라기준(cursor_positions, section_titles)
        show_message(
                title="데이터 추출완료",
                message="데이터 추출 완료!",
                info=f"데이터화 한 섹션 :\n{정렬된_section_titles}"
            )
        data2_button.setEnabled(True)
    except Exception as e:
        print(f"API 호출 중 오류 발생: {e}")
        show_message(
            title="오류 발생",
            message="오류 발생!",
            info=f"발생 오류 :\n{e}"
        )

def 데이터입력버튼():
    #테이블 가져오기
    #인서트데이터
    global selected_table, 정렬된_cursor_positions, 정렬된_section_titles
    if selected_table == '안전관리계획서1':
        data_dic = 안전계획서1_dic
    elif selected_table =='안전관리계획서2_1':
        data_dic = 안전계획서2_1_dic
    elif selected_table =='안전관리계획서2_2':
        data_dic = 안전계획서2_2_dic
    elif selected_table =='안전관리계획서2_3':
        data_dic = 안전계획서2_3_dic
    elif selected_table =='안전관리계획서2_4':
        data_dic = 안전계획서2_4_dic
    try:
        insert_data(selected_table,data_dic,정렬된_section_titles, 정렬된_cursor_positions )
        show_message(
                title="db 입력 완료",
                message="데이터를 db에 저장하였습니다.",
                info=f"테이블:\n{selected_table}"
            )
        #테이블 ui리셋하기 추가
    except Exception as e:
        print(f"API 호출 중 오류 발생: {e}")
        show_message(
            title="오류 발생",
            message="오류 발생!",
            info=f"발생 오류 :\n{e}"
        )
    
def 템플릿제작버튼():
    try:
        hwp9 = Hwp()
        if hwp9.Open("템플릿\\Temp.hwp"):
            # 6. 텍스트 스캔 및 섹션 타이틀 추출
            hwp9.MoveDocBegin()
            hwp9.init_scan()
            extracted_texts = []  # 추출된 텍스트 저장
            while True:
                state, text = hwp9.get_text()
                if text and text.strip():  # 공백 제외
                    clean_text = text.replace("\r\n", "").replace("\n", "").replace("\r", "")
                    print(clean_text)  # 정리된 텍스트 출력
                    extracted_texts.append(clean_text)
                if state <= 1:  # 종료 조건
                    break
            hwp9.release_scan()
        
            for i in extracted_texts:
                if hwp9.find_forward(i):
                    hwp9.MoveParaEnd()
                    hwp9.BreakPara()
                    hwp9.create_field(name=i, direction=i)
            hwp9.save_as(r"템플릿\자동필드생성.hwp")
            hwp9.clear()
            hwp9.quit()

        show_message(
                title="템플릿 만들기",
                message="temp파일로 템플릿 초안을 만들었습니다",
                info=f"저장 경로 : 템플릿\\자동필드생성.hwp"
            )
    except Exception as e:
        show_message(
            title="오류 발생",
            message="오류 발생!",
            info=f"발생 오류 :\n{e}"
        )
#endregion

#-------------------------------

#region 엑셀 입출력
# 내보내기
def export_to_excel():
    """
    현재 선택된 테이블 데이터를 엑셀 파일로 내보냅니다.
    """
    global table_selector, cursor
    selected_table = table_selector.currentText()  # 선택된 테이블 이름 가져오기
    if not selected_table:
        return

    # 데이터 가져오기
    cursor.execute(f"SELECT * FROM {selected_table}")
    rows = cursor.fetchall()
    columns = [description[0] for description in cursor.description]

    # Pandas DataFrame으로 변환
    df = pd.DataFrame(rows, columns=columns)

    # 파일 저장 다이얼로그 열기
    file_path, _ = QFileDialog.getSaveFileName(None, "엑셀로 저장", "", "Excel Files (*.xlsx);;All Files (*)")
    if file_path:
        df.to_excel(file_path, index=False)  # 엑셀로 저장
        show_message(
            title="엑셀 내보내기 완료",
            message="데이터를 엑셀 파일로 내보냈습니다.",
            info=f"파일 경로: {file_path}"
        )

# 엑셀 불러오기
def import_from_excel():
    """
    엑셀 파일에서 데이터를 읽어와 선택된 테이블에 삽입합니다.
    """
    global cursor, conn, table_selector
    selected_table = table_selector.currentText()  # 선택된 테이블 이름 가져오기
    if not selected_table:
        return

    # 파일 열기 다이얼로그
    file_path, _ = QFileDialog.getOpenFileName(None, "엑셀 파일 열기", "", "Excel Files (*.xlsx);;All Files (*)")
    if not file_path:
        return

    # 엑셀 파일 읽기
    df = pd.read_excel(file_path)

    # 기존 데이터 삭제 및 새 데이터 삽입
    cursor.execute(f"DELETE FROM {selected_table}")
    for _, row in df.iterrows():
        placeholders = ", ".join(["?"] * len(row))
        cursor.execute(f"INSERT INTO {selected_table} VALUES ({placeholders})", tuple(row))

    conn.commit()
    show_message(
        title="엑셀 불러오기 완료",
        message="엑셀 데이터를 성공적으로 DB에 반영했습니다.",
        info=f"파일 경로: {file_path}"
    )
    load_table_data()  # 테이블 데이터 갱신
#endregion

#-------------------------------

#region 안전 템플릿 변환 버튼 함수
def 안전계획서업로드버튼():
    """파일 5개 주소 받아와서 파일명만 리스트에 표시"""
    global file_list_widget  # 전역 변수 사용
    global 안전계획서_리스트, pdf_files, xlsx_files  # 확장자별 파일 리스트

     # 기존 리스트 초기화
    안전계획서_리스트.clear()
    pdf_files.clear()
    xlsx_files.clear()


    if file_list_widget is None:
        print("오류: file_list_widget이 초기화되지 않았습니다.")
        return  # 위젯이 없으면 함수 종료

    # 다중 파일 선택 다이얼로그
    files, _ = QFileDialog.getOpenFileNames(
        None, "파일 선택", "", "All Files (*);;HWP Files (*.hwp);;PDF Files (*.pdf);;Excel Files (*.xlsx)", options=QFileDialog.Options()
    )
    
    if files:
        # 최대 10개까지만 추가
        selected_files = files[:10]

        # 확장자별로 파일 분류
        for file in selected_files:
            ext = os.path.splitext(file)[1].lower()  # 확장자 추출
            if ext == ".hwp":
                안전계획서_리스트.append(file)
            elif ext == ".pdf":
                pdf_files.append(file)
            elif ext in [".xls", ".xlsx"]:
                xlsx_files.append(file)

        # UI 리스트 업데이트 (파일명만 추가)
        file_list_widget.clear()
        file_list_widget.addItems([os.path.basename(f) for f in selected_files])
    print("📄 HWP 파일 리스트:", 안전계획서_리스트)
    print("📂 PDF 파일 리스트:", pdf_files)
    print("📊 Excel 파일 리스트:", xlsx_files)

def 안전계획서데이터추출버튼():
    """ 
        1. 자료 합치기(원본) 
        2. 섹션명 추출(표, 띄어쓰기 등 삭제 후 스캔) 
        3. 원본에서 섹션명으로 포지션추출
    """
    global hwp11, 추출한_섹션명_리스트, 정렬된_cursor_positions, 정렬된_section_titles, cursor_positions, section_titles
    # 1. 자료 합치기
    hwp11 = Hwp(new= True, visible=True) #통짜파일 원본 인스턴스
    for i in 안전계획서_리스트:
        hwp11.insert(i, format="HWP",move_doc_end=True)

    # 2. 섹션명 추출(표, 띄어쓰기 등 삭제 후 스캔)  -> 포지션 검색
    try:
        추출한_섹션명_리스트.clear()
        cursor_positions.clear()
        section_titles.clear()
        정렬된_cursor_positions.clear() 
        정렬된_section_titles.clear()

        추출한_섹션명_리스트 = 섹션명_추출(hwp11)

        처음지점추가(hwp11)
        문단시작지점추가(hwp11)
        마지막지점추가(hwp11)
        print(추출한_섹션명_리스트)

        중간위치추가(hwp11, sec_list=추출한_섹션명_리스트)
        
        정렬된_cursor_positions, 정렬된_section_titles = 재정렬_파라기준(cursor_positions, section_titles)
        print(정렬된_cursor_positions, 정렬된_section_titles)
        show_message(
                title="데이터 추출완료",
                message="데이터 추출 완료!",
                info=f"데이터화 한 섹션 :\n{정렬된_section_titles}"
            )
    except Exception as e:
        print(f"API 호출 중 오류 발생: {e}")
        show_message(
            title="데이터 추출 오류 발생",
            message="데이터 추출 오류 발생!",
            info=f"발생 오류 :\n{e}"
        )

def 안전계획서_변환템플릿삽입버튼():
    """
    외부 함수를 호출하여 복붙 데이터를 삽입합니다.
    """
    global hwp11, 추출한_섹션명_리스트, 정렬된_cursor_positions, 정렬된_section_titles, cursor_positions, section_titles
    if not 정렬된_cursor_positions or not 정렬된_section_titles:
        show_message(
            title="삽입 실패",
            message="삽입할 데이터가 없습니다.",
            icon=QMessageBox.Warning
        )
        return
    hwp12 = Hwp(new=True)##템플릿 인스턴스 
    # 현재 날짜와 시간 가져오기
    now = datetime.datetime.now()
    formatted_date = now.strftime("%y%m%d.")  # 날짜: YYMMDD.
    #if hwp12.open("템플릿\\유해방지계획서_템플릿.hwp"):

    #템플릿 불러와서 필드에 입력하기
    if hwp12.FileOpen():
        try:
            hwp12.pdf
            filename = os.path.basename(hwp12.Path)
            filename = os.path.splitext(filename)[0]  # 확장자 제거
            for i in range(len(정렬된_cursor_positions) - 1):  # 마지막 인덱스는 n+1을 위해 제외
                section_start = 정렬된_cursor_positions[i]      # 현재 인덱스(n)의 (list, para, pos)
                section_end = 정렬된_cursor_positions[i + 1]    # 다음 인덱스(n+1)의 (list, para, pos)
                section_name = 정렬된_section_titles[i]         # 현재 인덱스(n)의 섹션 타이틀
                
                #필드가 있을 때만
                if hwp12.move_to_field(section_name):
                    #복
                    print(f"{section_name} 필드 있음 / 복붙 시도")
                    # 섹션 시작 위치로 이동
                    hwp11.SetPos(*section_start)  
                    hwp11.MoveNextParaBegin()  # 다음 단락으로 이동
                    hwp11.Select()  # 섹션 시작 위치에서 선택 시작
                    # 섹션 끝 위치로 이동
                    hwp11.SetPos(*section_end)
    
                    if i != len(정렬된_cursor_positions) - 2: #맨 마지막이 아닐때
                        hwp11.MoveLeft()  # 끝 위치에서 한 글자 왼쪽으로 이동해 선택 범위 조정
                    data = hwp11.GetTextFile("HWP","saveblock")
                    #붙
                    if data is None: print('data가 없습니다: 복붙 실패')
                    else : 
                        hwp12.SetTextFile(data,"HWP")
                        print(f"{section_start},{section_end},{section_name} 복붙 성공")
                    hwp11.Cancel() # 셀선택 초기화
                    data = None #data 변수 초기화
                else: print(f'{section_name}필드가 없습니다.')
        except Exception as e:
            print(f"API 호출 중 오류 발생: {e}")
            show_message(
                title="복붙 오류 발생",
                message="복붙 오류 발생!",
                info=f"발생 오류 :\n{e}"
            )
    #저장하기
    try:
        if hwp12.save_as(f"{formatted_date} {filename}_변환_초안.hwp"):
            hwp12.clear()
            hwp12.quit()

        show_message(
            title="삽입 성공",
            message="데이터를 성공적으로 삽입했습니다.",
            info=f"저장 경로: {formatted_date} {filename}_변환_초안.hwp"
        )
    except Exception as e:
        print(f"API 호출 중 오류 발생: {e}")
        show_message(
            title="저장 종료 오류 발생",
            message="저장 종료 오류 발생!",
            info=f"발생 오류 :\n{e}"
        )

def 변환초기화버튼():
    """초기화"""
    global file_list_widget, hwp11, 추출한_섹션명_리스트, 정렬된_cursor_positions, 정렬된_section_titles, cursor_positions, section_titles, main_pdf_path
    global sub_pdf_dir, output_pdf_path, search_text, insertions, sub_pdf_names, sub_pdf_paths
    추출한_섹션명_리스트.clear()
    cursor_positions.clear()
    section_titles.clear()
    정렬된_cursor_positions.clear() 
    정렬된_section_titles.clear()
    file_list_widget.clear()
    hwp11.clear()
    hwp11.quit()
    main_pdf_path = None
    sub_pdf_dir = None
    output_pdf_path = None
    search_text = "후면 첨부참조"
    insertions = []
    sub_pdf_names = []  # PDF 파일 이름 (확장자 없는 형태)
    sub_pdf_paths = []  # PDF 파일의 실제 경로
    sub_doc.close()
    main_doc.close()

def 설정버튼():
    """
    설정
    -pdf아웃풋 값 수정
    -pdf 검색 단어 수정
    -pdf 회의록 옮기기 추가 시에 여기에서 값 수정
    """
#endregion

#-------------------------------

def 계획서PDF병합():
    """
    업로드된 계획서를 PDF변환 후 PDF병합하기
    """
    ###todo

def 메인PDF업로드():
    """
    1.메인PDF경로지정
    """
    global main_pdf_path  # 전역 변수로 설정하여 다른 곳에서도 접근 가능
    global file_list_widget  # 전역 변수 사용
    # 파일 탐색기 창 열기 (PDF 파일만 선택 가능)
    file_path, _ = QFileDialog.getOpenFileName(None, "메인 PDF 선택", "", "PDF 파일 (*.pdf)")

    if file_path:  # 사용자가 파일을 선택했을 경우
        main_pdf_path = file_path  # 선택한 경로 저장
        print(f"📂 선택된 메인 PDF: {main_pdf_path}")  # 콘솔 출력 (디버깅용)

        # 선택된 경로를 UI에 표시하는 경우 (예: QLabel 사용)
        # pdf_title.setText(main_pdf_path)  
        
        QMessageBox.information(None, "PDF 선택 완료", f"선택된 파일:\n{main_pdf_path}")  # 팝업창 표시
        # UI 리스트 업데이트 (파일명만 추가)
        file_list_widget.clear()
        file_list_widget.addItem(os.path.basename(main_pdf_path))  # ✅ 파일명만 추가
    else:
        QMessageBox.warning(None, "선택 취소", "메인 PDF를 선택하지 않았습니다.")  # 취소한 경우 경고 메시지

def clean_text(text):
    """
    불필요한 유니코드 문자 및 공백을 정리하는 함수
    """
     # 모든 제어 문자 제거
    text = ''.join(ch for ch in text if unicodedata.category(ch)[0] != 'C' or ch == '\n')
    
    # 연속된 공백 정리 (줄바꿈 제외)
    text = re.sub(r'[^\S\n]+', ' ', text).strip()  
    return text

def extract_text_from_pdf(pdf_path):
    """
    PDF 파일에서 텍스트를 추출하여 정제하는 함수
    """
    text = ""
    try:
        doc = fitz.open(pdf_path)
        for page in doc:
            raw_text = page.get_text("text")
            cleaned_text = clean_text(raw_text)  # 유니코드 정제 적용
            text += cleaned_text + "\n\n"
    except Exception as e:
        print(f"오류 발생: {e}")
    return text
def 병합된PDF생성(pdf_list):
    """
    여러 개의 PDF를 병합하여 하나의 PDF로 반환
    """
    merged_pdf = fitz.open()

    for pdf in pdf_list:
        doc = fitz.open(pdf)  # PDF 열기
        merged_pdf.insert_pdf(doc)  # 병합
        doc.close()

    return merged_pdf  # 병합된 PDF 반환
def 서브PDF폴더선택():
    """
    1. 서브PDF폴더선택
    2. 메인PDF의 '후면첨부참조'의 앞 키워드와 PDF명 매칭 확인
    """
    global main_pdf_path, sub_pdf_dir, search_text, insertions, sub_pdf_names, sub_pdf_paths, sub_doc, main_doc, pdf_files
    global file_list_widget  # 전역 변수 사용
    # 폴더 선택 창 열기
    folder_path = QFileDialog.getExistingDirectory(None, "서브 PDF 폴더 선택", "")

    if folder_path:  # 사용자가 폴더를 선택했을 경우
        sub_pdf_dir = folder_path  # 선택한 폴더 경로 저장
        print(f"📂 선택된 서브 PDF 폴더: {sub_pdf_dir}")  # 콘솔 출력 (디버깅용)

        # 서브 PDF 목록 가져오기
        sub_pdf_names.clear()
        sub_pdf_paths.clear()

        for filename in os.listdir(sub_pdf_dir):
            if filename.endswith(".pdf"):
                clean_name = re.sub(r"^\d+[-\d]*\.\s*", "", os.path.splitext(filename)[0])  # 앞 숫자 제거
                sub_pdf_names.append(clean_name)  # 파일 이름 저장
                sub_pdf_paths.append(os.path.join(sub_pdf_dir, filename))  # 파일 경로 저장
        
         # ✅ pdf_files 사용하여 메인 PDF 설정
        if not pdf_files:
            QMessageBox.warning(None, "오류", "메인 PDF 파일이 선택되지 않았습니다.")
            return

        if len(pdf_files) == 1:
            main_doc = fitz.open(pdf_files[0])  # 단일 파일이면 그대로 사용
        else:
            main_doc = 병합된PDF생성(pdf_files)  # 다중 PDF 병합 후 사용

        삽입예정list =[]
        opened_docs = []
        matched_keywords = []  # 🔹 매칭할 키워드 리스트
        unmatched_keywords = []  # 🔹 매칭되지 않은 키워드 리스트

        try:
            # 🔹 1단계: 모든 페이지에서 검색어 추출
            for page_num in range(len(main_doc)):
                raw_text = main_doc[page_num].get_text("text")  # 현재 페이지의 텍스트 추출
                text = clean_text(raw_text)  # 🔹 유니코드 정제 적용
                lines = text.split("\n")  # 줄 단위로 나누기

                for line in lines:
                    if search_text in line:  # 🔹 검색어 포함된 줄 찾기
                        matched_keyword = line.split(search_text)[0]  # 검색어 앞부분만 추출
                        matched_keyword = re.sub(r"[^\w\s()·]", "", matched_keyword).strip()  # 특수문자 제거
                        matched_keywords.append((page_num, matched_keyword))  # 🔹 페이지 번호와 함께 저장
                        break  # 해당 페이지에서 검색어를 찾으면 더 이상 검사하지 않음

            # 🔹 2단계: 검색어 리스트와 서브 PDF 비교
            for page_num, matched_keyword in matched_keywords:
                best_match_index = None

                for i, sub_pdf_name in enumerate(sub_pdf_names):
                    clean_sub_pdf_name = re.sub(r"^\d+[-\d]*\.\s*", "", sub_pdf_name)  # 🔹 앞 숫자 제거
                    if matched_keyword == clean_sub_pdf_name:  # 🔹 완전 일치하는 경우
                        best_match_index = i
                        break  # 첫 번째 완전 일치 항목을 선택하고 루프 종료
                
                # 🔹 3단계: 매칭된 경우 PDF 삽입 준비
                if best_match_index is not None:
                    best_match_name = sub_pdf_names[best_match_index]  # 매칭된 PDF 파일명 가져오기
                    sub_doc = fitz.open(sub_pdf_paths[best_match_index])  # 서브 PDF 열기
                    opened_docs.append(sub_doc)  # 🔹 열린 PDF 저장
                    insertions.append((page_num + 1, sub_doc))  # 해당 페이지 뒤에 삽입할 정보 저장
                    print(f"✅ 삽입 예정: {matched_keyword} → {best_match_name}")
                    삽입예정list.append((matched_keyword, best_match_name))
                else:
                    unmatched_keywords.append(matched_keyword)  # 🔹 매칭되지 않은 키워드 저장

            # 🔹 4단계: 매칭되지 않은 검색어 처리
            if unmatched_keywords:
                unmatched_list = "\n".join(unmatched_keywords)
                print(f"⚠ 매칭 실패: 다음 키워드에 대한 PDF가 없습니다.\n{unmatched_list}")
                show_copyable_message(
                    title="매칭 실패!",
                    message=f"검색 단어: {search_text}",
                    info=f"다음 키워드가 매칭되지 않았습니다. 모든 PDF를 확인한 후 다시 시도하세요:\n{unmatched_list}"
                )
                raise Exception("매칭 실패: 일부 파일이 누락됨")  # 예외 발생 -> except 블록으로 이동
        except Exception as e:
            # 🔴 오류 발생 시, 열린 모든 sub_doc 닫기
            print(f"🚨 오류 발생: {e}")
            for doc in opened_docs:
                doc.close()
            print("모든 열린 서브 PDF 파일을 닫았습니다.")
            raise  # 오류를 다시 발생시켜 프로그램이 종료되거나 처리 가능하도록 함

        # 삽입 예정 리스트를 문자열로 변환하여 표시
        message = "\n".join([f"{i+1}. {keyword} → {pdf}" for i, (keyword, pdf) in enumerate(삽입예정list)])

        # 팝업 창 띄우기
        QMessageBox.information(None, "매칭 성공 목록", message)

        # UI 리스트 업데이트 (파일명만 추가)
        file_list_widget.clear()
        file_list_widget.addItem(f"검색 단어 : {search_text if search_text else '없음'}")
        file_list_widget.addItems([f"삽입 예정: {keyword} → {pdf}" for keyword, pdf in 삽입예정list])

    else:
        QMessageBox.warning(None, "선택 취소", "서브 PDF 폴더를 선택하지 않았습니다.")  # 취소한 경우 경고 메시지

def PDF끼워넣기():
    """
    PDF 삽입 후 저장
    """
    global sub_doc, main_doc, insertions, output_pdf_path
    for insert_page, sub_doc in sorted(insertions, key=lambda x: x[0], reverse=True):
        main_doc.insert_pdf(sub_doc, start_at=insert_page)  # 서브 PDF 삽입
        sub_doc.close()  # 사용이 끝난 서브 PDF 닫기

    # 현재 날짜와 시간 가져오기
    now = datetime.datetime.now()
    formatted_date = now.strftime("%y%m%d.")  # 날짜: YYMMDD.

    # 최종 결과 저장
    main_doc.save(f"{formatted_date} {output_pdf_path}.pdf")  # 결과물을 새로운 PDF 파일로 저장
    main_doc.close()  # 메인 PDF 닫기
    print(f"처리 완료! 파일 저장: {output_pdf_path}")  # 완료 메시지 출력
    show_message(
            title="처리 성공",
            message="처리 성공!",
            info=f"저장 경로: {formatted_date} {output_pdf_path}.pdf"
        )
#-------------------------------

# UI 생성
def create_ui():

    """
    PyQt5 기반의 UI를 생성하고 초기 설정을 수행합니다.
    """
    #region ui초기화
    global app, table_selector, table_widget, id_input, fetch_button, paste_button, hwpfile_label, hwpfile_button, reset_button,data1_button, data2_button, temp_button
    global file_list_widget
    app = QApplication(sys.argv)  # QApplication 생성

    window = QMainWindow()
    window.setWindowTitle("NALMEOK1.1")  # 창 제목 설정
    window.setGeometry(100, 100, 800, 600)  # 창 크기 설정

    layout = QVBoxLayout()  # 전체 레이아웃
    #endregion

    #region 테이블 선택, 한글파일 선택
    table_selector_label = QLabel("테이블을 선택하세요:")
    layout.addWidget(table_selector_label)

    table_selector = QComboBox()
    layout.addWidget(table_selector)
    table_selector.currentIndexChanged.connect(load_table_data)

    # 한글 파일선택 버튼
    hwpfile_layout = QHBoxLayout()
    hwpfile_button = QPushButton("추출할 한글 파일 선택 / 해제")
    hwpfile_button.clicked.connect(한글파일선택버튼)  # 엑셀 내보내기 연결
    hwpfile_button.setEnabled(True)  # 초기 상태 비활성화
    hwpfile_layout.addWidget(hwpfile_button)

    reset_button = QPushButton("연결해제/초기화")
    reset_button.clicked.connect(초기화버튼)  # 엑셀 내보내기 연결
    reset_button.setEnabled(False)  # 초기 상태 비활성화
    hwpfile_layout.addWidget(reset_button)

    # path라벨
    hwpfile_label = QLabel("path : ")
    hwpfile_label.setFixedHeight(30)  # 라벨 높이 설정
    #hwpfile_label.setWordWrap(True)   # 텍스트 줄 바꿈 활성화
    hwpfile_label.setStyleSheet("QLabel { font-size: 8pt; padding: 5px; border: 1px solid #ccc; }")  # 스타일 추가
    hwpfile_layout.addWidget(hwpfile_label, stretch=1)
    
    layout.addLayout(hwpfile_layout)
    #endregion

    #region 데이터 추출, db 입력 등
    hwpdata_layout = QHBoxLayout()
    data1_button = QPushButton("데이터 추출")
    data1_button.clicked.connect(데이터추출버튼)  
    data1_button.setEnabled(False)  # 초기 상태 비활성화
    hwpdata_layout.addWidget(data1_button)

    # 데이터 입력 버튼
    data2_button = QPushButton("DB 입력")
    data2_button.clicked.connect(데이터입력버튼) 
    data2_button.setEnabled(False)  # 초기 상태 비활성화
    hwpdata_layout.addWidget(data2_button)

    layout.addLayout(hwpdata_layout)

    #endregion

    #region id입력, 내용삽입, 테이블
    input_layout = QHBoxLayout()
    input_label = QLabel("입력할 ID 입력 (쉼표로 구분):")
    input_layout.addWidget(input_label)

    id_input = QLineEdit()
    id_input.setPlaceholderText("예: 55, 60, 70")  # 힌트 텍스트
    input_layout.addWidget(id_input)

    fetch_button = QPushButton("데이터 가져오기")
    fetch_button.clicked.connect(fetch_custom_data)  # 버튼 클릭 시 fetch_custom_data 호출
    input_layout.addWidget(fetch_button)

    paste_button = QPushButton("복붙 내용 삽입")
    paste_button.clicked.connect(execute_paste)  # 버튼 클릭 시 execute_paste 호출
    paste_button.setEnabled(False)  # 초기 상태 비활성화
    input_layout.addWidget(paste_button)

    layout.addLayout(input_layout)

    # 테이블 위젯
    table_widget = QTableWidget()
    layout.addWidget(table_widget)
    #endregion

    #region 엑셀 입출, 초기화

    # ✅ QFrame 생성 
    xl_frame = QFrame()
    xl_frame.setFrameShape(QFrame.Shape.Box)  # 테두리 추가
    xl_frame.setFrameShadow(QFrame.Shadow.Raised)  # 그림자 효과 추가
    xl_frame.setStyleSheet("QFrame { border: 1.5px solid black; padding: 1px; }")  # 스타일 설정

    # ✅ 전체 레이아웃 (수직)
    xl_layout = QVBoxLayout()

    # ✅ 타이틀 추가 (프레임 내부에 포함)
    xl_title = QLabel("✅ 엑셀 입력 / 출력 ")  
    xl_title.setStyleSheet("""font-size: 10pt; 
                               font-weight: bold; 
                               padding: 1px;
                               background: transparent;
                               border:none;
                               """)
    xl_title.setAlignment(Qt.AlignCenter)  # 텍스트 중앙 정렬
    xl_layout.addWidget(xl_title)  # 가장 위에 추가
    
    # ✅ 버튼 레이아웃 (두 번째 줄)
    button_layout = QHBoxLayout()

    #엑셀 입력 출력
    export_button = QPushButton("📤 엑셀로 내보내기")
    export_button.clicked.connect(export_to_excel)  # 엑셀 내보내기 연결
    button_layout.addWidget(export_button)

    import_button = QPushButton("📥 엑셀에서 불러오기")
    import_button.clicked.connect(import_from_excel)  # 엑셀 불러오기 연결
    button_layout.addWidget(import_button)
    
    # 템플릿 제작 버튼
    temp_button = QPushButton("🔧 temp 스캔 후 필드 생성")# temp파일 문자열 스캔 후 필드 넣기
    temp_button.clicked.connect(템플릿제작버튼)
    button_layout.addWidget(temp_button)

    # ✅ 버튼 레이아웃을 전체 레이아웃에 추가
    xl_layout.addLayout(button_layout)

    # ✅ 프레임에 최종 레이아웃 설정
    xl_frame.setLayout(xl_layout)

    # ✅ 전체 레이아웃에 추가
    layout.addWidget(xl_frame)

    #endregion
    
    #region 안전 템플릿 변환 
    # ✅ QFrame 생성 (안전 -> 유해 변환 관련 버튼 그룹)
    utrans_frame = QFrame()
    utrans_frame.setFrameShape(QFrame.Shape.Box)  # 테두리 추가
    utrans_frame.setFrameShadow(QFrame.Shadow.Raised)  # 그림자 효과 추가
    utrans_frame.setStyleSheet("QFrame { border: 1.5px solid black; padding: 1px; }")  # 스타일 설정

    # ✅ 버튼 그룹 정리 (2개씩 묶어서 배치)
    utrans_layout = QVBoxLayout()

    # 공통 버튼
    row1_layout = QHBoxLayout()

    utrans1_button = QPushButton("📂 파일 불러오기 [1]")
    utrans1_button.clicked.connect(안전계획서업로드버튼)  
    row1_layout.addWidget(utrans1_button)

    utrans8_button = QPushButton("⛔ 초기화")
    utrans8_button.clicked.connect(변환초기화버튼)  
    row1_layout.addWidget(utrans8_button)

    utrans8_button = QPushButton("⚙ 설정")
    utrans8_button.clicked.connect(설정버튼)  
    row1_layout.addWidget(utrans8_button)

    utrans_layout.addLayout(row1_layout)

    # 파일 리스트 표시할 QListWidget    
    file_list_widget = QListWidget()
    file_list_widget.setStyleSheet("QListWidget { font-size: 10pt; padding: 5px; border: 1px solid #ccc; }")  # 스타일 추가
    #file_list_widget.setFixedHeight(110)  # 높이를 110px로 고정
    # 고정 높이 제거 & 유동적 크기 설정
    file_list_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
    utrans_layout.addWidget(file_list_widget)

    # ✅ 타이틀 추가 (프레임 내부에 포함)
    utrans_title = QLabel("✅ 안전계획서 : 템플릿 변환 ")  
    utrans_title.setStyleSheet("""font-size: 10pt; 
                               font-weight: bold; 
                               padding: 1px;
                               background: transparent;
                               border:none;
                               """)
    utrans_title.setAlignment(Qt.AlignCenter)  # 텍스트 중앙 정렬
    utrans_layout.addWidget(utrans_title)  # 가장 위에 추가

    # 줄 버튼
    row2_layout = QHBoxLayout()

    utrans2_button = QPushButton("📊 데이터 추출(temp생성) [2]") 
    utrans2_button.clicked.connect(안전계획서데이터추출버튼)  
    row2_layout.addWidget(utrans2_button)

    utrans3_button = QPushButton("📝 변환 템플릿에 삽입 [3]")
    utrans3_button.clicked.connect(안전계획서_변환템플릿삽입버튼) 
    row2_layout.addWidget(utrans3_button)

    utrans_layout.addLayout(row2_layout)

    # ✅ 타이틀 추가 (프레임 내부에 포함)
    pdf_title = QLabel("✅ PDF 병합 + 취합 ")  
    pdf_title.setStyleSheet("""font-size: 10pt; 
                               font-weight: bold; 
                               padding-top: 10px;
                               background: transparent;
                               border:none;
                               """)
    pdf_title.setAlignment(Qt.AlignCenter)  # 텍스트 중앙 정렬
    utrans_layout.addWidget(pdf_title)  # 가장 위에 추가

    # 1줄 버튼
    pdf1_layout = QHBoxLayout()

    # pdf1_button = QPushButton("📜 한글→PDF 변환 + 병합 [2]") 
    # pdf1_button.clicked.connect(계획서PDF병합)  
    # pdf1_layout.addWidget(pdf1_button)

    # pdf2_button = QPushButton("메인 PDF 업로드 [2-2]")
    # pdf2_button.clicked.connect(메인PDF업로드) 
    # pdf1_layout.addWidget(pdf2_button)
    
    utrans_layout.addLayout(pdf1_layout)

    #2번째 줄
    pdf2_layout = QHBoxLayout()

    pdf3_button = QPushButton("🗂 서브 PDF 폴더 선택 [2]")
    pdf3_button.clicked.connect(서브PDF폴더선택) 
    pdf1_layout.addWidget(pdf3_button)

    pdf4_button = QPushButton("📝 PDF 끼워넣기 [3]")
    pdf4_button.clicked.connect(PDF끼워넣기) 
    pdf1_layout.addWidget(pdf4_button)
    
    utrans_layout.addLayout(pdf2_layout)

    utrans_frame.setLayout(utrans_layout)

    # ✅ 테이블 내부에 추가 (자연스럽게 배치됨)
    layout.addWidget(utrans_frame)  


    

    # 메인 위젯 설정
    container = QWidget()
    container.setLayout(layout)
    window.setCentralWidget(container)

    return window

    #endregion


if __name__ == "__main__":
    initialize_db('안전관리계획서1',안전계획서1_dic)
    initialize_db('안전관리계획서2_1',안전계획서2_1_dic)
    initialize_db('안전관리계획서2_2',안전계획서2_2_dic)
    initialize_db('안전관리계획서2_3',안전계획서2_3_dic)
    initialize_db('안전관리계획서2_4',안전계획서2_4_dic)
    init_db()  # 데이터베이스 초기화
    main_window = create_ui()  # UI 생성
    load_table_list()  # 테이블 목록 로드
    main_window.show()  # 메인 창 표시
    sys.exit(app.exec_())  # 이벤트 루프 실행
#endregion


