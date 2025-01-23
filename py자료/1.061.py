#pyinstaller --onefile --noconsole --icon=myicon.ico 1.061.py
import tkinter as tk
from tkinter import ttk
from pyhwpx import Hwp
import re

#region 초기화 부분
#인스턴스생성
hwp = Hwp()

# Tkinter 창 생성
root = tk.Tk()
root.title("송나겸 돈 복사 매크로(v1.061)")

#파일 경로 표시
file_path_frame = tk.Frame(root, highlightbackground="black", highlightthickness=1, padx=10, pady=10)
file_path_frame.grid(row=0, column=0, padx=15, pady=5, sticky="ew")

# 공지사항 라벨
notice_label = tk.Label(file_path_frame, text="2002방식 조판부호 사용 해야 작동함. / 창이 활성화된 상태에서만 단축키가 작동함.", font=("Arial", 10), anchor="w", fg="red")
notice_label.grid(row=1, column=0, columnspan=2, sticky="w", padx=5)  # 아래쪽 전체 너비로 표시
notice_labe2 = tk.Label(file_path_frame, text="파일 선택 눌러서 편집할 파일을 연결 / 연결 안되면 다 끄고 다시 해주세요...(후졌음)", font=("Arial", 10), anchor="w", fg="red")
notice_labe2.grid(row=2, column=0, columnspan=2, sticky="w", padx=5)  # 아래쪽 전체 너비로 표시
# 파일 경로 라벨
file_path_label = tk.Label(file_path_frame, text="파일 경로: 없음", font=("Arial", 10), anchor="e")
file_path_label.grid(row=0, column=0, sticky="w", padx=70)  # 왼쪽 정렬

#표 초기화
Table_list = [] # 모든 표의 갯수 찾기
Table_Total = 0
Table_index = 0 #전역으로 사용할 표 인덱스

#그림 초기화
pic_list = []
pic_Total = 0
pic_index =0

#문단 변수
para = 0
dic = hwp.get_style_dict()

# Notebook(탭 컨트롤러) 생성
notebook = ttk.Notebook(root)
notebook.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")  # pack 대신 grid 사용
# 첫 번째 탭
tab1 = ttk.Frame(notebook)
notebook.add(tab1, text="표 매크로")
# 두 번째 탭
tab2 = ttk.Frame(notebook)
notebook.add(tab2, text="그림 매크로")

# 탭의 열 비율 설정
tab1.grid_columnconfigure(0, weight=1, minsize=300)  # frame_table_movement가 있는 열
tab1.grid_columnconfigure(1, weight=3, minsize=300)  # frame_inspector가 있는 열
tab2.grid_columnconfigure(0, weight=1, minsize=300)  
tab2.grid_columnconfigure(1, weight=3, minsize=300)  





def Table_init():
    global Table_list
    global Table_Total
    Table_list.clear()
    Table_Total = 0
    for i in hwp.ctrl_list:
        if i.UserDesc == "표":#컨트롤이 표일 경우
            Table_list.append(i)#리스트에 저장
    Table_Total = len(Table_list) # 문서안의 모든 표의 갯수
    align_ep_label.config(text=f"전체 표 갯수={Table_Total}")
    
def pic_init():
     global pic_list 
     global pic_Total
     pic_list.clear()
     pic_Total =0
     for i in hwp.ctrl_list:
         if i.UserDesc == "그림":#컨트롤이 그림일 경우
            pic_list.append(i)#리스트에 저장
     pic_Total = len(pic_list) # 문서안의 모든 표의 갯수
     position_pic_label.config(text=f"전체 그림 갯수={pic_Total}")
def update_file_path_label():
    """파일 경로를 업데이트하고 Label에 표시"""
    if hwp.Path:
        file_path = hwp.Path
        file_path_label.config(text=f"파일 경로: {file_path}")
    else:
        file_path_label.config(text="파일 경로: 없음")
        print("파일 경로: 없음")

def select_file():
    """파일을 재선택"""
    global dic
    try:
        hwp.Run("FileOpen")
        Table_init()
        pic_init()
        dic = hwp.get_style_dict()
        update_file_path_label()
        처음으로()
    except Exception as e:
        print(f"파일 선택 중 오류: {e}")
def res_ctrl():
        Table_init()    
        pic_init()

# 파일 선택 버튼
file_reselect_button = tk.Button(file_path_frame, text="파일 선택", command=select_file)
file_reselect_button.grid(row=0, column=0, sticky="w", padx=5)  # 오른쪽 정렬

#표/그림 재탐색
res_button = tk.Button(file_path_frame, text="표/그림 재탐색", command=res_ctrl)
res_button.grid(row=0, column=1, sticky="e", padx=5)  # 오른쪽 정렬
#endregion

#region 매크로 함수
def 셀_전체선택():
    hwp.TableCellBlock() 
    hwp.TableCellBlockExtend()
    hwp.TableCellBlockExtend()
def 셀_전체선택ex():
    hwp.ShapeObjTableSelCell()
    hwp.TableCellBlockExtend()
    hwp.TableCellBlockExtend()
def 표라인_전체투명():
    hwp.TableCellBorderNo()
def 표라인_안쪽실선():
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.GetDefault("CellBorder", pset.HSet)
    pset.WidthVert = hwp.HwpLineWidth("0.12mm")
    pset.TypeVert = hwp.HwpLineType("Solid")
    pset.WidthHorz = hwp.HwpLineWidth("0.12mm")
    pset.TypeHorz = hwp.HwpLineType("Solid")
    hwp.HAction.Execute("CellBorder", pset.HSet)
def 표라인_위선(t : str):
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.GetDefault("CellBorder", pset.HSet)
    pset.BorderWidthTop = hwp.HwpLineWidth(f"{t}mm")
    pset.BorderTypeTop = hwp.HwpLineType("Solid")
    return hwp.HAction.Execute("CellBorder", pset.HSet)
def 표라인_아래선(t : str):
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.GetDefault("CellBorder", pset.HSet)
    pset.BorderWidthBottom = hwp.HwpLineWidth(f"{t}mm")
    pset.BorderTypeBottom = hwp.HwpLineType("Solid")
    return hwp.HAction.Execute("CellBorder", pset.HSet)
def 표라인_양옆투명():
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.GetDefault("CellBorder", pset.HSet)
    pset.BorderWidthRight = hwp.HwpLineWidth("0.12mm")
    pset.BorderTypeRight = hwp.HwpLineType("None")
    pset.BorderWidthLeft = hwp.HwpLineWidth("0.12mm")
    pset.BorderTypeLeft = hwp.HwpLineType("None")
    pset.FillAttr.GradationAlpha = 0
    pset.FillAttr.ImageAlpha = 0
    hwp.HAction.Execute("CellBorder", pset.HSet)
def 표라인_양옆선(t : str):
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.GetDefault("CellBorder", pset.HSet)
    pset.BorderWidthRight = hwp.HwpLineWidth(f"{t}mm")
    pset.BorderTypeRight = hwp.HwpLineType("Solid")
    pset.BorderWidthLeft = hwp.HwpLineWidth(f"{t}mm")
    pset.BorderTypeLeft = hwp.HwpLineType("Solid")
    pset.FillAttr.GradationAlpha = 0
    pset.FillAttr.ImageAlpha = 0
    hwp.HAction.Execute("CellBorder", pset.HSet)
def 표색_없음():#안됨
    #hwp.cell_fill([255,255,255]) #흰색으로 채우기
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.GetDefault("CellBorder", pset.HSet)
    pset.FillAttr.type = hwp.BrushType("NullBrush") # 소문자임.. ㅜㅜㅜㅜㅜ
    pset.FillAttr.GradationAlpha = 0
    pset.FillAttr.WindowsBrush = 0
    pset.FillAttr.ImageAlpha = 0
    return hwp.HAction.Execute("CellBorder", pset.HSet)#false,,,,,,,
def 셀세로정렬():
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
    pset.HSet.SetItem("ShapeType", 3)
    pset.HSet.SetItem("ShapeCellSize", 0)
    pset.ShapeTableCell.VertAlign = hwp.VAlign("Center")
    hwp.HAction.Execute("TablePropertyDialog", pset.HSet)
def 안여백지정해제():
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
    pset.HSet.SetItem("ShapeType", 3)
    pset.HSet.SetItem("ShapeCellSize", 0)
    pset.ShapeTableCell.HasMargin = 0
    hwp.HAction.Execute("TablePropertyDialog", pset.HSet)
def 글자처럼해제_자리차지():
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
    pset.TextWrap = hwp.TextWrapType("TopAndBottom")
    pset.TreatAsChar = 0
    pset.HSet.SetItem("ShapeType", 3)
    pset.HSet.SetItem("ShapeCellSize", 0)
    hwp.HAction.Execute("TablePropertyDialog", pset.HSet)
def 셀단위로나눔_제목줄반복():
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
    pset.RepeatHeader = 1
    pset.PageBreak = hwp.TableBreak("Table")
    pset.HSet.SetItem("ShapeType", 3)
    pset.HSet.SetItem("ShapeCellSize", 0)
    hwp.HAction.Execute("TablePropertyDialog", pset.HSet)
def 표색_헤드_회색217():
    hwp.TableColPageUp()
    hwp.cell_fill([217,217,217]) 
def 표색_헤드_없음():
    hwp.TableColPageUp()
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.GetDefault("CellBorder", pset.HSet)
    pset.FillAttr.type = hwp.BrushType("NullBrush") # 소문자임.. ㅜㅜㅜㅜㅜ
    pset.FillAttr.GradationAlpha = 0
    pset.FillAttr.WindowsBrush = 0
    pset.FillAttr.ImageAlpha = 0
    return hwp.HAction.Execute("CellBorder", pset.HSet)#false,,,,,,,
def 표라인_헤드_밑줄(t :str):
    hwp.TableColPageUp()
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.GetDefault("CellBorder", pset.HSet)
    pset.BorderWidthBottom = hwp.HwpLineWidth(f"{t}mm")
    pset.BorderTypeBottom = hwp.HwpLineType("Solid")
    hwp.HAction.Execute("CellBorder", pset.HSet)
    hwp.TableCellBlockExtend()
def 표라인_헤드_두줄():
    hwp.TableColPageUp()
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.GetDefault("CellBorder", pset.HSet)
    pset.BorderWidthBottom = hwp.HwpLineWidth("0.5mm")
    pset.BorderTypeBottom = hwp.HwpLineType("DoubleSlim")
    hwp.HAction.Execute("CellBorder", pset.HSet)
    hwp.TableCellBlockExtend()
def 표라인_표주_윗선(t:str):
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.Run("TableCellBlockRow")
    hwp.HAction.GetDefault("CellBorder", pset.HSet)
    pset.BorderWidthTop = hwp.HwpLineWidth(f"{t}mm")
    pset.BorderTypeTop = hwp.HwpLineType("Solid")
    pset.TypeVert = hwp.HwpLineType("None")
    pset.BorderTypeBottom = hwp.HwpLineType("None")
    pset.BorderTypeRight = hwp.HwpLineType("None")
    pset.BorderTypeLeft = hwp.HwpLineType("None")
    pset.FillAttr.GradationAlpha = 0
    pset.FillAttr.ImageAlpha = 0
    hwp.HAction.Execute("CellBorder", pset.HSet)
    hwp.TableCellBlockExtend()
def 표좌우맞춤():
    pagedef_dict = hwp.get_pagedef_as_dict()
    paper_width = int(pagedef_dict['용지폭'])       # 용지 폭
    margin_left = int(pagedef_dict['왼쪽'])         # 왼쪽 여백
    margin_right = int(pagedef_dict['오른쪽'])      # 오른쪽 여백
    # 실제 문서 폭 계산
    actual_width = paper_width - margin_left - margin_right
    epsilon = align_ep.get()
    hwp.set_table_width(actual_width-epsilon)
    셀_전체선택()
def 위캡션2mm():
    hwp.ShapeObjAttachCaption()
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
    pset.HSet.SetItem("ShapeType", 3)
    pset.ShapeCaption.Side = hwp.SideType("Top")
    pset.ShapeCaption.Width = hwp.MiliToHwpUnit(30.0)
    pset.ShapeCaption.Gap = hwp.MiliToHwpUnit(2.0)
    pset.ShapeCaption.CapFullSize = 0
    pset.HSet.SetItem("ShapeCellSize", 0)
    hwp.HAction.Execute("TablePropertyDialog", pset.HSet)
    #표 안쪽으로 들어가는 코드 추가
    hwp.MoveDown()
def 아래캡션3mm():
    hwp.ShapeObjAttachCaption()
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
    pset.HSet.SetItem("ShapeType", 3)
    pset.ShapeCaption.Side = hwp.SideType("Bottom")
    pset.ShapeCaption.Width = hwp.MiliToHwpUnit(30.0)
    pset.ShapeCaption.Gap = hwp.MiliToHwpUnit(3.0)
    pset.ShapeCaption.CapFullSize = 0
    pset.HSet.SetItem("ShapeCellSize", 0)
    hwp.HAction.Execute("TablePropertyDialog", pset.HSet)
    #표 안쪽으로 들어가는 코드 추가
    hwp.MoveUp()
def 그림아래캡션3mm():
    global pic_index
    global pic_list
    hwp.ShapeObjAttachCaption()
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
    pset.HSet.SetItem("ShapeType", 3)
    pset.ShapeCaption.Side = hwp.SideType("Bottom")
    pset.ShapeCaption.Width = hwp.MiliToHwpUnit(30.0)
    pset.ShapeCaption.Gap = hwp.MiliToHwpUnit(3.0)
    pset.ShapeCaption.CapFullSize = 0
    pset.HSet.SetItem("ShapeCellSize", 0)
    hwp.HAction.Execute("TablePropertyDialog", pset.HSet)
    hwp.select_ctrl(pic_list[pic_index])
def 표위치_2단(Horz: str, Vert: str):
    """ Justify왼  Center(왼),Left(가운데),Right(오른)  /
        Vert는 Top, Bottom
    """
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
    pset.HorzAlign = hwp.HAlign(Horz)
    pset.HorzRelTo = hwp.HorzRel("Page")
    pset.VertAlign = hwp.VAlign(Vert)
    pset.VertRelTo = hwp.VertRel("Page")
    pset.HorzOffset = hwp.MiliToHwpUnit(0.0)
    pset.VertOffset = hwp.MiliToHwpUnit(0.0)
    if Vert =="Top":
        pset.OutsideMarginBottom = hwp.MiliToHwpUnit(7.0)
    else:
        pset.OutsideMarginBottom = hwp.MiliToHwpUnit(0.0)
    if Vert == "Bottom":
        pset.OutsideMarginTop = hwp.MiliToHwpUnit(7.0)
    else:
        pset.OutsideMarginTop = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginRight = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginLeft = hwp.MiliToHwpUnit(0.0)
    pset.HSet.SetItem("ShapeType", 3)
    pset.HSet.SetItem("ShapeCellSize", 0)
    hwp.HAction.Execute("TablePropertyDialog", pset.HSet)
def 그림위치_2단(Horz: str, Vert: str):
    """Justify왼    "Left": 중앙   Right:오른쪽
        Vert는 Top, Bottom
    """
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("ShapeObjDialog", pset.HSet)
    pset.HorzAlign = hwp.HAlign(Horz)
    pset.HorzRelTo = hwp.HorzRel("Page")
    pset.VertAlign = hwp.VAlign(Vert)
    pset.VertRelTo = hwp.VertRel("Page")
    pset.HorzOffset = hwp.MiliToHwpUnit(0.0)
    pset.VertOffset = hwp.MiliToHwpUnit(0.0)
    if Vert =="Top":
        pset.OutsideMarginBottom = hwp.MiliToHwpUnit(7.0)
    else:
        pset.OutsideMarginBottom = hwp.MiliToHwpUnit(0.0)
    if Vert == "Bottom":
        pset.OutsideMarginTop = hwp.MiliToHwpUnit(7.0)
    else:
        pset.OutsideMarginTop = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginRight = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginLeft = hwp.MiliToHwpUnit(0.0)
    pset.HSet.SetItem("ShapeType", 1)
    hwp.HAction.Execute("ShapeObjDialog", pset.HSet)

def 표위치_1단():
    """ Horz은  Center(왼),Left(가운데),Right(오른)  /
        Vert는 Top, Bottom
    """
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
    pset.HorzAlign = hwp.HAlign("Left")
    pset.HorzRelTo = hwp.HorzRel("Page")
    pset.VertAlign = hwp.VAlign("Top")
    pset.VertRelTo = hwp.VertRel("Para")
    pset.HorzOffset = hwp.MiliToHwpUnit(0.0)
    pset.VertOffset = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginBottom = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginTop = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginRight = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginLeft = hwp.MiliToHwpUnit(0.0)
    pset.HSet.SetItem("ShapeType", 3)
    pset.HSet.SetItem("ShapeCellSize", 0)
    hwp.HAction.Execute("TablePropertyDialog", pset.HSet)
def 표위치_1단_2미리():
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)
    pset.HorzAlign = hwp.HAlign("Left")
    pset.HorzRelTo = hwp.HorzRel("Page")
    pset.VertAlign = hwp.VAlign("Top")
    pset.VertRelTo = hwp.VertRel("Para")
    pset.HorzOffset = hwp.MiliToHwpUnit(0.0)
    pset.VertOffset = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginBottom = hwp.MiliToHwpUnit(2.0)
    pset.OutsideMarginTop = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginRight = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginLeft = hwp.MiliToHwpUnit(0.0)
    pset.HSet.SetItem("ShapeType", 3)
    pset.HSet.SetItem("ShapeCellSize", 0)
    hwp.HAction.Execute("TablePropertyDialog", pset.HSet)
def 그림위치_1단():
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("ShapeObjDialog", pset.HSet)
    pset.HorzAlign = hwp.HAlign("Left")
    pset.HorzRelTo = hwp.HorzRel("Page")
    pset.VertAlign = hwp.VAlign("Top")
    pset.VertRelTo = hwp.VertRel("Page")
    pset.HorzOffset = hwp.MiliToHwpUnit(0.0)
    pset.VertOffset = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginBottom = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginTop = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginRight = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginLeft = hwp.MiliToHwpUnit(0.0)
    pset.HSet.SetItem("ShapeType", 1)
    hwp.HAction.Execute("ShapeObjDialog", pset.HSet)
def 그림위치_1단_2미리():
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("ShapeObjDialog", pset.HSet)
    pset.HorzAlign = hwp.HAlign("Left")
    pset.HorzRelTo = hwp.HorzRel("Page")
    pset.VertAlign = hwp.VAlign("Top")
    pset.VertRelTo = hwp.VertRel("Page")
    pset.HorzOffset = hwp.MiliToHwpUnit(0.0)
    pset.VertOffset = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginBottom = hwp.MiliToHwpUnit(2.0)
    pset.OutsideMarginTop = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginRight = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginLeft = hwp.MiliToHwpUnit(0.0)
    pset.HSet.SetItem("ShapeType", 1)
    hwp.HAction.Execute("ShapeObjDialog", pset.HSet)
def 그림_안여백외곽선없음():
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("ShapeObjDialog", pset.HSet)
    pset.ShapeDrawImageAttr.InsideMarginBottom = hwp.MiliToHwpUnit(0.0)
    pset.ShapeDrawImageAttr.InsideMarginTop = hwp.MiliToHwpUnit(0.0)
    pset.ShapeDrawImageAttr.InsideMarginRight = hwp.MiliToHwpUnit(0.0)
    pset.ShapeDrawImageAttr.InsideMarginLeft = hwp.MiliToHwpUnit(0.0)
    pset.ShapeDrawLineAttr.style = hwp.HwpLineType("None")
    pset.HSet.SetItem("ShapeType", 1)
    hwp.HAction.Execute("ShapeObjDialog", pset.HSet)
def 그림글자처럼():
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("ShapeObjDialog", pset.HSet)
    pset.TreatAsChar = 1
    pset.HSet.SetItem("ShapeType", 1)
    hwp.HAction.Execute("ShapeObjDialog", pset.HSet)
def 그림밖여백0():
    pset = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("ShapeObjDialog", pset.HSet)
    pset.OutsideMarginBottom = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginTop = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginRight = hwp.MiliToHwpUnit(0.0)
    pset.OutsideMarginLeft = hwp.MiliToHwpUnit(0.0)
    pset.HSet.SetItem("ShapeType", 1)
    hwp.HAction.Execute("ShapeObjDialog", pset.HSet)

#endregion

# region 함수 모음
def 되돌리기(event=None):
    hwp.Run("Undo")
def 다시실행(event=None):  
    hwp.Run("Redo")  
def on_CellLineMacro(event=None):
    """버튼 클릭 이벤트: 드롭다운 값 출력"""
    윗선        = dropdown1.get()
    아래선      = dropdown2.get()
    양옆선      = dropdown3.get()
    내부선      = dropdown4.get()
    전체배경    = dropdown5.get()
    헤드배경    = dropdown6.get()
    헤드밑줄    = dropdown7.get()
    표주윗선    = dropdown8.get()

    if 윗선 == "냅둠":
        pass
    elif 윗선 == "0.12":
        표라인_위선("0.12")
    elif 윗선 == "0.4":
        표라인_위선("0.4")
    else:
        print(f"{윗선}이 선택되지 않았습니다.")    

    if 아래선 == "냅둠":
        pass
    elif 아래선 == "0.12":
        표라인_아래선("0.12")
    elif 아래선 == "0.4":
        표라인_아래선("0.4")
    else:
        print(f"{아래선}이 선택되지 않았습니다.")    

    if 양옆선 == "냅둠":
        pass
    elif 양옆선 == "투명":
        표라인_양옆투명()
    elif 양옆선 == "0.12":
        표라인_양옆선("0.12")
    elif 양옆선 == "0.4":
        표라인_양옆선("0.4")
    else:
        print(f"{양옆선}이 선택되지 않았습니다.")   
    
    if 내부선 == "냅둠":
        pass
    elif 내부선 == "0.12":
        표라인_안쪽실선()
    else:
        print(f"{내부선}이 선택되지 않았습니다.") 

    if 전체배경 == "냅둠":
        pass
    elif 전체배경 == "색없음":
        표색_없음()
    else:
        print(f"{전체배경}이 선택되지 않았습니다.")     

    if 헤드배경 == "냅둠":
        pass
    elif 헤드배경 == "없음":
        표색_헤드_없음()
        셀_전체선택()
    elif 헤드배경 == "회색(217)":
        표색_헤드_회색217()
        셀_전체선택()
    else:
        print(f"{헤드배경}이 선택되지 않았습니다.")  
    
    if 헤드밑줄 == "냅둠":
        pass
    elif 헤드밑줄 == "0.12":
        표라인_헤드_밑줄("0.12")
    elif 헤드밑줄 == "0.4":
        표라인_헤드_밑줄("0.4")
    elif 헤드밑줄 == "두줄":
        표라인_헤드_두줄()
    else:
        print(f"{헤드밑줄}이 선택되지 않았습니다.")  

    if 표주윗선 == "표주 없음(냅둠)":
        pass
    elif 표주윗선 == "0.12(아래투명/위는0.12)":
        표라인_표주_윗선("0.12")
    elif 표주윗선 == "0.4(아래투명/위0.4)":
        표라인_표주_윗선("0.4")
    else:
        print(f"{표주윗선}이 선택되지 않았습니다.")  
def 스타일적용():
    표내용 = style7.get() 
    #표캡션 = style8.get()
    표헤드 = style9.get() 
    표주 = style10.get()
    if 표내용 == 1:
        hwp.Run("StyleShortcut8")
    # if 표캡션 == 1:
    #     #hwp.get_into_table_caption()
    #     hwp.Run("ShapeObjInsertCaptionNum")
    #     #hwp.Run("SelectAll")
    #     #hwp.Run("StyleShortcut8")
    #     #hwp.Run("ReturnPrevPos")
    #     셀_전체선택()
    if 표헤드 == 1:
        hwp.TableColPageUp()
        hwp.Run("TableColEnd")
        hwp.Run("StyleShortcut9")
        hwp.TableCellBlockExtend()
    if 표주 == 1:
        hwp.HAction.Run("TableCellBlockRow")
        hwp.Run("StyleShortcut10")
        hwp.TableCellBlockExtend()
def 표여백정렬초기화():
    """안여백1,밖여백0,셀세로중앙정렬,안여백지정해제,글자처럼해제,자리차지,셀단위로나눔,제목줄반복,위캡션2미리"""
    if var1.get() == 1:
        hwp.set_table_inside_margin(1,1,1,1) #안여백 1mm로 밀기
    if var2.get() == 1:
        hwp.TableVAlignCenter() #셀 세로 중앙정렬
    if var3.get() == 1:
        안여백지정해제()
    if var4.get() == 1:
        글자처럼해제_자리차지()#1 위계가 있음.
    if var5.get() == 1:
        셀단위로나눔_제목줄반복()#2 위계가 있음.
    if var6.get() == 1:
        hwp.set_table_outside_margin(0,0,0,0)#밖여백 0mm로 밀기 
    셀_전체선택()   
def 그림초기화():
    """"""
    if pic_var1.get() == 1: #글자처럼. 자리차지
        글자처럼해제_자리차지()
    else :
        그림글자처럼()
    if pic_var2.get() == 1: #그림캡션달기
        그림아래캡션3mm()
    if pic_var3.get() == 1: #밖여백0
        그림밖여백0()
    if pic_var4.get() == 1: #그림외곽선 없므
        그림_안여백외곽선없음()
def 표캡션():
    위캡션2mm()
    셀_전체선택()   
def 양옆맞추기(event=None):
    """양옆맞추기(오차0.4임)"""
    표좌우맞춤()
def 단2맞추기(event=None):
    """2단일때 양옆 맞추기"""
    width = align2_ep.get()
    hwp.set_table_width(width)
    셀_전체선택()
def 그림용(event=None):
    """그림 표용 세팅"""
    아래캡션3mm()
    셀_전체선택()
def handle_1단용(event=None):
    표위치_1단()
def handle_1단용_2미리(event=None):
    표위치_1단_2미리()
def handle_2단_왼_상(event=None):
    표위치_2단("Justify","Top")
def handle_2단_가운_상(event=None):
    표위치_2단("Left","Top")
def handle_2단_오른_상(event=None):
    표위치_2단("Right","Top")
def handle_2단_왼_하(event=None):
    표위치_2단("Justify","Bottom")
def handle_2단_가운_하(event=None):
    표위치_2단("Left","Bottom")
def handle_2단_오른_하(event=None):
    표위치_2단("Right","Bottom")

def 그림_1단용(event=None):
    그림위치_1단()
def 그림_1단용_2미리(event=None):
    그림위치_1단_2미리()
def 그림_2단_왼_상(event=None):
    그림위치_2단("Justify","Top")
def 그림_2단_가운_상(event=None):
    그림위치_2단("Left","Top")
def 그림_2단_오른_상(event=None):
    그림위치_2단("Right","Top")
def 그림_2단_왼_하(event=None):
    그림위치_2단("Justify","Bottom")
def 그림_2단_가운_하(event=None):
    그림위치_2단("Left","Bottom")
def 그림_2단_오른_하(event=None):
    그림위치_2단("Right","Bottom")
def 이전표(event=None):
    global Table_index
    global Table_list
    # Spinbox에서 현재 값을 가져오고 -1
    current_value = current_table_index.get()
    prev_index = max(0, current_value - 1)  # 최소값 제한
    current_table_index.set(prev_index)  # Spinbox에 값 설정
    Table_index = prev_index  # 전역 변수 업데이트
    print(f"현재 Table_index: {Table_index}")
    hwp.get_into_nth_table(Table_index)#인덱스 표의 첫번째 셀로 이동
    셀_전체선택()
def 다음표(event=None):
    global Table_index
    global Table_list
    # Spinbox에서 현재 값을 가져오고 +1
    current_value = current_table_index.get()
    next_index = max(0, min(current_value + 1, Table_Total - 1))  # 최대값 제한
    current_table_index.set(next_index)  # Spinbox에 값 설정
    Table_index = next_index  # 전역 변수 업데이트
    print(f"현재 Table_index: {Table_index}")
    hwp.get_into_nth_table(Table_index)#인덱스 표의 첫번째 셀로 이동
    셀_전체선택()
def 처음으로(event=None):
    """처음으로"""
    global Table_index
    global Table_list
    # Spinbox와 Table_index를 동기화
    Table_index = 0  # 전역 변수 Table_index를 0으로 설정
    current_table_index.set(Table_index)  # Spinbox에 값 반영
    try:
        hwp.get_into_nth_table(Table_index)  # 첫 번째 표로 이동
        셀_전체선택()
    except Exception as e:
        print(f"처음으로 이동 중 오류 발생: {e}")
def 이전그림(event=None):
    global pic_index
    global pic_list
    # Spinbox에서 현재 값을 가져오고 -1
    current_value = current_pic_index.get()
    prev_index = max(0, current_value - 1)  # 최소값 제한
    current_pic_index.set(prev_index)  # Spinbox에 값 설정
    pic_index = prev_index  # 전역 변수 업데이트
    print(f"현재 Pic_index: {pic_index}")
    hwp.select_ctrl(pic_list[pic_index])
    #hwp.get_into_nth_table(pic_index)#인덱스 표의 첫번째 셀로 이동
    #셀_전체선택()
def 다음그림(event=None):
    global pic_index
    global pic_list
    # Spinbox에서 현재 값을 가져오고 +1
    current_value = current_pic_index.get()
    next_index = max(0, min(current_value + 1, pic_Total - 1))  # 최대값 제한
    current_pic_index.set(next_index)  # Spinbox에 값 설정
    pic_index = next_index  # 전역 변수 업데이트
    print(f"현재 Pic_index: {pic_index}")
    hwp.select_ctrl(pic_list[pic_index])
    
def 처음으로_그림(event=None):
    """처음으로"""
    global pic_index
    global pic_list
    if not pic_list:  # 그림 리스트가 비어 있을 경우 예외 처리
        print("그림이 없습니다.")
        return
    # Spinbox와 Table_index를 동기화
    pic_index = 0  # 전역 변수 Table_index를 0으로 설정
    current_pic_index.set(pic_index)  # Spinbox에 값 반영
    try:
        hwp.select_ctrl(pic_list[pic_index])
    except Exception as e:
        print(f"처음으로 이동 중 오류 발생: {e}")
def bind_button_to_key(button, key, modifier=None):
    """
    버튼을 특정 키에 바인딩. Alt, Ctrl과 같은 modifier도 처리 가능.
    
    Args:
        button: 바인딩할 Tkinter 버튼.
        key: 바인딩할 키 (특수 문자 포함).
        modifier: 'Alt' 또는 'Control' 등과 같은 수정 키.
    """
    def key_action(event=None):
        # 버튼 눌림 효과
        button.config(relief="sunken")  # 눌림 상태
        button.update_idletasks()  # 즉시 업데이트
        root.after(100, lambda: button.config(relief="raised"))  # 100ms 후 원래 상태 복구
        button.invoke()  # 버튼의 command 실행

    # 특수 문자 처리용 매핑
    special_key_map = {
        ",": "comma",
        ".": "period",
        "/": "slash",
        ";": "semicolon",
        "'": "apostrophe",
        "[": "bracketleft",
        "]": "bracketright",
        "-": "minus",
        "=": "equal"
    }

    # 키 이벤트 문자열 생성
    if modifier == "Alt":
        if key in special_key_map:
            root.bind(f"<Alt-Key-{special_key_map[key]}>", key_action)
        else:
            root.bind(f"<Alt-Key-{key}>", key_action)
    elif modifier == "Control":
        if key in special_key_map:
            root.bind(f"<Control-Key-{special_key_map[key]}>", key_action)
        else:
            root.bind(f"<Control-Key-{key}>", key_action)
    else:
        # 기본 바인딩 (대소문자 자동 처리)
        if key in special_key_map:
            root.bind(f"<Key-{special_key_map[key]}>", key_action)
        else:
            root.bind(f"<Key-{key}>", key_action)
#endregion

#region 첫번째 탭 : 표 매크로
#region 되돌리기 다시실행
btn_undo = tk.Button(file_path_frame, text="되돌리기 [Z]", command=되돌리기)
btn_undo.grid(row=0, column=2, padx=5, pady=5, sticky="e")
bind_button_to_key(btn_undo, "z")

btn_redo = tk.Button(file_path_frame, text="다시실행 [X]", command=다시실행)
btn_redo.grid(row=0, column=3, padx=5, pady=5, sticky="e")
bind_button_to_key(btn_redo, "x")
#endregion
#region 상태 변수
current_table_index = tk.IntVar(value=0)
align_ep = tk.DoubleVar(value= 0.0)
align2_ep = tk.DoubleVar(value= 66)
# 체크박스 상태 변수 (IntVar)
var1 = tk.IntVar(value=1)
var2 = tk.IntVar(value=1)
var3 = tk.IntVar(value=1)
var4 = tk.IntVar(value=1)
var5 = tk.IntVar(value=1)
var6 = tk.IntVar(value=1)

style7 = tk.IntVar(value=1)
style8 = tk.IntVar(value=1)
style9 = tk.IntVar(value=1)
style10 = tk.IntVar(value=1)
#endregion
##########################################
#region 표이동 프레임
frame_table_movement = tk.Frame(tab1, highlightbackground="black", highlightthickness=2, padx=10, pady=10)
frame_table_movement.grid(row=0, column=0, pady=5, sticky="ew")

position_description = tk.Label(frame_table_movement, text="표 이동", font=("Arial", 12, "bold"))
position_description.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

align_ep_label = tk.Label(frame_table_movement, text=f"전체 표 갯수={Table_Total}")
align_ep_label.grid(row=0, column=2)

current_index_label = tk.Label(frame_table_movement, text="현재 표 인덱스:")
current_index_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")

index_spinbox = tk.Spinbox(frame_table_movement, from_=0, to=0, textvariable=current_table_index, width=5)
index_spinbox.grid(row=1, column=1, padx=5, pady=5)

btn_prev = tk.Button(frame_table_movement, text="이전 [Q]", command=이전표)
btn_prev.grid(row=2, column=0, padx=5, pady=5, sticky="e")
bind_button_to_key(btn_prev, "q")

btn_next = tk.Button(frame_table_movement, text="다음 [W]", command=다음표)
btn_next.grid(row=2, column=1, padx=5, pady=5, sticky="w")
bind_button_to_key(btn_next, "w")

btn_begin = tk.Button(frame_table_movement, text="처음으로 [B]", command=처음으로)
btn_begin.grid(row=1, column=2, padx=5, pady=5)
bind_button_to_key(btn_begin, "b")
처음으로()
#endregion
##########################################
#region 'Style' 
frame_style = tk.Frame(tab1, highlightbackground="black", highlightthickness=2, padx=10, pady=10)
frame_style.grid(row=2, column=1, padx=5,  pady=5, sticky="nsew")

inspector_label = tk.Label(frame_style, text="style 적용", font=("Arial", 12, "bold"))
inspector_label.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

style_check7 = tk.Checkbutton(frame_style, text="표내용(스타일8)", variable=style7)
#style_check8 = tk.Checkbutton(frame_style, text="표캡션(스타일8)", variable=style8)
style_check9 = tk.Checkbutton(frame_style, text="표헤드(스타일9)", variable=style9)
style_check10 = tk.Checkbutton(frame_style, text="표주(스타일0)", variable=style10)
style_check7.grid(row=0, column=2,  padx=5, pady=5,sticky="w")
#style_check8.grid(row=1, column=2,  padx=5, pady=5,sticky="w")
style_check9.grid(row=1, column=2,  padx=5, pady=5,sticky="w")
style_check10.grid(row=2, column=2,  padx=5, pady=5,sticky="w")

btn_style = tk.Button(frame_style, text="실행 [G]", command=스타일적용)
btn_style.grid(row=2, column=0,  padx=5, pady=5,sticky="ew")
bind_button_to_key(btn_style, "g")
#endregion
##########################################
#region '표 기본설정' 프레임
frame_init_macro = tk.Frame(tab1, highlightbackground="black", highlightthickness=2, padx=10, pady=10)
frame_init_macro.grid(row=0, column=1, padx=5,pady=5,sticky="ew")

init_macro_label = tk.Label(frame_init_macro, text="표 기본설정", font=("Arial", 12, "bold")) # 기본 정렬 / 전체, 2단 채우기, 입실론, 그림 캡션
init_macro_label.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

btn_init = tk.Button(frame_init_macro, text="[E] 실행 >>>", command=표여백정렬초기화)
btn_init.grid(row=1, column=0,  padx=5, pady=5,sticky="ew")
bind_button_to_key(btn_init, "e")

check1 = tk.Checkbutton(frame_init_macro, text="안여백 1mm", variable=var1)
check2 = tk.Checkbutton(frame_init_macro, text="셀세로중앙정렬", variable=var2)
check3 = tk.Checkbutton(frame_init_macro, text="안여백지정해제", variable=var3)
check4 = tk.Checkbutton(frame_init_macro, text="글자해제,자리차지(1)", variable=var4)
check5 = tk.Checkbutton(frame_init_macro, text="셀단위로나눔,제목반복(2)", variable=var5)
check6 = tk.Checkbutton(frame_init_macro, text="밖여백 0mm", variable=var6)
check1.grid(row=0, column=2,  padx=5, pady=5,sticky="w")
check2.grid(row=1, column=2,  padx=5, pady=5,sticky="w")
check3.grid(row=2, column=2,  padx=5, pady=5,sticky="w")
check4.grid(row=0, column=3,  padx=5, pady=5,sticky="w")
check5.grid(row=1, column=3,  padx=5, pady=5,sticky="w")
check6.grid(row=2, column=3,  padx=5, pady=5,sticky="w")
#endregion
##########################################
#region '표 fill, 캡션' 프레임
frame_modi_macro = tk.Frame(tab1, highlightbackground="black", highlightthickness=2, padx=10, pady=10)
frame_modi_macro.grid(row=1, column=1,padx=5, pady=5,sticky="nsew")

# init_macro_label = tk.Label(frame_modi_macro, text="표 추가설정", font=("Arial", 12, "bold")) 
# init_macro_label.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

btn_align = tk.Button(frame_modi_macro, text="전체 Fill [A]", command=양옆맞추기)
btn_align.grid(row=3, column=0, padx=5, pady=5,sticky="ew")
bind_button_to_key(btn_align, "a")

btn_align2 = tk.Button(frame_modi_macro, text="2단 Fill [S]", command=단2맞추기)
btn_align2.grid(row=3, column=1, padx=5, pady=5,sticky="ew")
bind_button_to_key(btn_align2, "s")

btn_pic = tk.Button(frame_modi_macro, text="표 캡션 [d]", command=표캡션)
btn_pic.grid(row=3, column=2,padx=5, pady=5,sticky="ew")
bind_button_to_key(btn_pic, "d")

btn_pic = tk.Button(frame_modi_macro, text="그림 캡션 [F]", command=그림용)
btn_pic.grid(row=3, column=3,padx=5, pady=5,sticky="ew")
bind_button_to_key(btn_pic, "f")

align_ep_label2 = tk.Label(frame_modi_macro, text="Fill e =")
align_ep_label2.grid(row=2, column=0, padx=0, pady=5, sticky="w")

index_spinbox = tk.Spinbox(frame_modi_macro, from_=-100, to=100, increment=0.1, textvariable=align_ep, width=3)
index_spinbox.grid(row=2, column=0, padx=(45,10), pady=5, sticky="w")

align_ep2_label2 = tk.Label(frame_modi_macro, text="2단 w=")
align_ep2_label2.grid(row=2, column=1, padx=0, pady=5, sticky="w")

index_entry = tk.Entry(frame_modi_macro, textvariable=align2_ep, width=4)
index_entry.grid(row=2, column=1, padx=(45,5), pady=5, sticky="w")
#endregion
##########################################
#region '선/배경 설정' 프레임
frame_cell_macro = tk.Frame(tab1, highlightbackground="black", highlightthickness=2, padx=10, pady=10)
frame_cell_macro.grid(row=1, column=0,pady=5, rowspan=2,sticky="ew")

# cell_macro_label = tk.Label(frame_cell_macro, text="선/배경 설정", font=("Arial", 12, "bold"))#4개 받아서 바꾸기, 배경없음 선없음
# cell_macro_label.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

#----------------------------------# 선택창

dropdown1_label = tk.Label(frame_cell_macro, text="윗 선:")
dropdown1_label.grid(row=1, column=0, padx=5, pady=3)

dropdown1 = ttk.Combobox(frame_cell_macro, values=["냅둠", "0.12", "0.4"])
dropdown1.set("0.4")
dropdown1.state(["readonly"])
dropdown1.grid(row=1, column=1, padx=5, pady=3)

dropdown2_label = tk.Label(frame_cell_macro, text="아래 선:")
dropdown2_label.grid(row=2, column=0, padx=5, pady=3)

dropdown2 = ttk.Combobox(frame_cell_macro, values=["냅둠", "0.12", "0.4"])
dropdown2.set("0.4")
dropdown2.state(["readonly"])
dropdown2.grid(row=2, column=1, padx=5, pady=3)

dropdown3_label = tk.Label(frame_cell_macro, text="양옆 선:")
dropdown3_label.grid(row=3, column=0, padx=5, pady=3)

dropdown3 = ttk.Combobox(frame_cell_macro, values=["냅둠","투명", "0.12", "0.4"])
dropdown3.set("투명")
dropdown3.state(["readonly"])
dropdown3.grid(row=3, column=1, padx=5, pady=3)

dropdown4_label = tk.Label(frame_cell_macro, text="내부 선:")
dropdown4_label.grid(row=4, column=0, padx=5, pady=3)

dropdown4 = ttk.Combobox(frame_cell_macro, values=["냅둠","0.12"])
dropdown4.set("냅둠")
dropdown4.state(["readonly"])
dropdown4.grid(row=4, column=1, padx=5, pady=3)

dropdown5_label = tk.Label(frame_cell_macro, text="전체 배경:")
dropdown5_label.grid(row=5, column=0, padx=5, pady=3)

dropdown5 = ttk.Combobox(frame_cell_macro, values=["냅둠","색없음"])
dropdown5.set("냅둠")
dropdown5.state(["readonly"])
dropdown5.grid(row=5, column=1, padx=5, pady=3)

dropdown6_label = tk.Label(frame_cell_macro, text="헤드 배경:")
dropdown6_label.grid(row=6, column=0, padx=5, pady=3)

dropdown6 = ttk.Combobox(frame_cell_macro, values=["냅둠", "없음", "회색(217)"])
dropdown6.set("냅둠")
dropdown6.state(["readonly"])
dropdown6.grid(row=6, column=1, padx=5, pady=3)

dropdown7_label = tk.Label(frame_cell_macro, text="헤드 밑 줄:")
dropdown7_label.grid(row=7, column=0, padx=5, pady=3)

dropdown7 = ttk.Combobox(frame_cell_macro, values=["냅둠","0.12","0.4", "두줄"])
dropdown7.set("냅둠")
dropdown7.state(["readonly"])
dropdown7.grid(row=7, column=1, padx=5, pady=3)

dropdown8_label = tk.Label(frame_cell_macro, text="표주 윗선:")
dropdown8_label.grid(row=8, column=0, padx=5, pady=3)

dropdown8 = ttk.Combobox(frame_cell_macro, values=["표주 없음(냅둠)","0.12(아래투명/위는0.12)", "0.4(아래투명/위0.4)"])
dropdown8.set("표주 없음(냅둠)")
dropdown8.state(["readonly"])
dropdown8.grid(row=8, column=1, padx=5, pady=3)

#-------------------#


btn_적용 = tk.Button(frame_cell_macro, text="<<< 적용 [R]", command=on_CellLineMacro)
btn_적용.grid(row=1, column=2, padx=5, pady=3,sticky="ew")
bind_button_to_key(btn_적용, "r")

btn_배경없음 = tk.Button(frame_cell_macro, text="All배경없음 [T]", command=표색_없음)
btn_배경없음.grid(row=7, column=2,  padx=5, pady=3,sticky="ew")
bind_button_to_key(btn_배경없음, "t")

btn_선없음 = tk.Button(frame_cell_macro, text="  All선없음 [Y]", command=표라인_전체투명)
btn_선없음.grid(row=8, column=2,  padx=5, pady=3,sticky="ew")
bind_button_to_key(btn_선없음, "y")
#endregion
##########################################
#region '표 위치 설정' 프레임 추가
frame1 = tk.Frame(tab1, highlightbackground="black", highlightthickness=2, padx=10, pady=10)
frame1.grid(row=4, column=0, pady=5, sticky="ew")

# 설명 레이블
# position_description = tk.Label(frame1, text="표 위치 매크로", font=("Arial", 12, "bold"))
# position_description.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

# 오른 단독 버튼
left_button = tk.Button(frame1, text="1단[K]", command=handle_1단용_2미리)
left_button.grid(row=1, column=0 ,padx=10, pady=7, sticky="nsew" )
bind_button_to_key(left_button, "k")



# 6개의 버튼 배열 (1열 3개, 2열 3개)
button1 = tk.Button(frame1, text="2단 왼 상[L]", command=handle_2단_왼_상)
button1.grid(row=1, column=1, padx=10, pady=7)
bind_button_to_key(button1, "l")

button2 = tk.Button(frame1, text="2단 가운 상[;]", command=handle_2단_가운_상)
button2.grid(row=1, column=2, padx=10, pady=7)
bind_button_to_key(button2, ";")

button3 = tk.Button(frame1, text="2단 오른 상[']", command=handle_2단_오른_상)##
button3.grid(row=1, column=3, padx=10, pady=7)
bind_button_to_key(button3, "'")

button4 = tk.Button(frame1, text="2단 왼 하[,]", command=handle_2단_왼_하)
button4.grid(row=2, column=1, padx=10, pady=7)
bind_button_to_key(button4, ",")

button5 = tk.Button(frame1, text="2단 가운 하[.]", command=handle_2단_가운_하)
button5.grid(row=2, column=2, padx=10, pady=7)
bind_button_to_key(button5, ".")

button6 = tk.Button(frame1, text="2단 오른 하[/]", command=handle_2단_오른_하)
button6.grid(row=2, column=3, padx=10, pady=7)
bind_button_to_key(button6, "/")
#endregion
##########################################
#region 매크로1~0 버튼 
def 매크로(i:str):
        hwp.Run(f"ScrMacroPlay{i}")# 1~ 11번까지 매크로 가능
        
macro_frame = tk.Frame(root, highlightbackground="black", highlightthickness=1, padx=10, pady=5)
macro_frame.grid(row=1, column=0, padx=15, pady=1, sticky="ew")
for i in range(1, 12):  # 매크로1, 매크로2, ..., 매크로9 버튼 생성
    btn = tk.Button(macro_frame, text=f"매크로{i}", command=lambda x=i: 매크로(str(x)))
    btn.grid(row=0, column=i, sticky="nsew")
    if i == 10:  # 10은 0으로 매핑
        bind_button_to_key(btn, "0", "Alt")
    elif i == 11:  # 11은 -로 매핑
        bind_button_to_key(btn, "-", "Alt")
    else:  # 1~9는 숫자 그대로 매핑
        bind_button_to_key(btn, str(i), "Alt")
#endregion
#endregion

#region 두번째 탭 : 그림 매크로
#region 상태변수
current_pic_index = tk.IntVar(value=0)
# 체크박스 상태 변수 (IntVar)
pic_var1 = tk.IntVar(value=1)
pic_var2 = tk.IntVar(value=1)
pic_var3 = tk.IntVar(value=1)
pic_var4 = tk.IntVar(value=1)
pic_var5 = tk.IntVar(value=1)
pic_var6 = tk.IntVar(value=1)
#endregion
##########################################
#region '그림 이동' 프레임
frame_pic_movement = tk.Frame(tab2, highlightbackground="black", highlightthickness=2, padx=10, pady=10)
frame_pic_movement.grid(row=0, column=0, pady=5, sticky="ew")

position_pic = tk.Label(frame_pic_movement, text="그림 이동", font=("Arial", 12, "bold"))
position_pic.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

position_pic_label = tk.Label(frame_pic_movement, text=f"전체 그림 갯수={pic_Total}")
position_pic_label.grid(row=0, column=2)
# 현재 표 인덱스와 이동 버튼
current_pic_i_label = tk.Label(frame_pic_movement, text="현재 그림 인덱스:")
current_pic_i_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")

index_spinbox = tk.Spinbox(frame_pic_movement, from_=0, to=0, textvariable=current_pic_index, width=5)
index_spinbox.grid(row=1, column=1, padx=5, pady=5)

btn_p_prev = tk.Button(frame_pic_movement, text="이전 []", command=이전그림)
btn_p_prev.grid(row=2, column=0, padx=5, pady=5, sticky="e")
#bind_button_to_key(btn_prev, "q")

btn_p_next = tk.Button(frame_pic_movement, text="다음 []", command=다음그림)
btn_p_next.grid(row=2, column=1, padx=5, pady=5, sticky="w")
#bind_button_to_key(btn_next, "w")

btn_begin2 = tk.Button(frame_pic_movement, text="처음으로 []", command=처음으로_그림)####그림ㅊ퍼음으로
btn_begin2.grid(row=1, column=2, padx=5, pady=5)
#bind_button_to_key(btn_begin2, "b")

#endregion
##########################################
#region '그림 기본 설정'프레임
frame_pic_init_macro = tk.Frame(tab2, highlightbackground="black", highlightthickness=2, padx=10, pady=10)
frame_pic_init_macro.grid(row=0, column=1, padx=5,pady=5,sticky="nsew")

init_macro_label = tk.Label(frame_pic_init_macro, text="그림 기본설정", font=("Arial", 12, "bold")) # 기본 정렬 / 전체, 2단 채우기, 입실론, 그림 캡션
init_macro_label.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

btn_init = tk.Button(frame_pic_init_macro, text="[] 실행 >>>", command=그림초기화)
btn_init.grid(row=1, column=0,  padx=5, pady=5,sticky="ew")
#bind_button_to_key(btn_init, "e")

pic_check1 = tk.Checkbutton(frame_pic_init_macro, text="글자해제,자리차지", variable=pic_var1)
pic_check2 = tk.Checkbutton(frame_pic_init_macro, text="그림캡션달기", variable=pic_var2)
pic_check3 = tk.Checkbutton(frame_pic_init_macro, text="밖여백 0mm", variable=pic_var3)
pic_check4 = tk.Checkbutton(frame_pic_init_macro, text="안여백 0mm, 외곽선없음", variable=pic_var4)
#pic_check5 = tk.Checkbutton(frame_pic_init_macro, text="그림외곽선없음", variable=pic_var5)
#pic_check6 = tk.Checkbutton(frame_pic_init_macro, text="그림여백 0mm", variable=pic_var6)
pic_check1.grid(row=0, column=2,  padx=5, pady=5,sticky="w")
pic_check2.grid(row=1, column=2,  padx=5, pady=5,sticky="w")
pic_check3.grid(row=0, column=3,  padx=5, pady=5,sticky="w")
pic_check4.grid(row=1, column=3,  padx=5, pady=5,sticky="w")
#pic_check5.grid(row=1, column=3,  padx=5, pady=5,sticky="w")
#pic_check6.grid(row=2, column=3,  padx=5, pady=5,sticky="w")
#endregion
##########################################
#region '그림 위치 설정'프레임
pic_pos_frame = tk.Frame(tab2, highlightbackground="black", highlightthickness=2, padx=10, pady=10)
pic_pos_frame.grid(row=1, column=0, pady=5, sticky="nsew")


# 오른 단독 버튼
pic_left_button = tk.Button(pic_pos_frame, text="1단[]", command=그림_1단용_2미리)
pic_left_button.grid(row=1, column=0, padx=10, pady=7, sticky="nsew" )



# 6개의 버튼 배열 (1열 3개, 2열 3개)
pic_button1 = tk.Button(pic_pos_frame, text="2단 왼 상[]", command=그림_2단_왼_상)
pic_button1.grid(row=1, column=1, padx=10, pady=7)
#bind_button_to_key(button1, "l")

pic_button2 = tk.Button(pic_pos_frame, text="2단 가운 상[]", command=그림_2단_가운_상)
pic_button2.grid(row=1, column=2, padx=10, pady=7)
#bind_button_to_key(button2, ";")

pic_button3 = tk.Button(pic_pos_frame, text="2단 오른 상[]", command=그림_2단_오른_상)##
pic_button3.grid(row=1, column=3, padx=10, pady=7)
#bind_button_to_key(button3, "'")

pic_button4 = tk.Button(pic_pos_frame, text="2단 왼 하[]", command=그림_2단_왼_하)
pic_button4.grid(row=2, column=1, padx=10, pady=7)
#bind_button_to_key(button4, ",")

pic_button5 = tk.Button(pic_pos_frame, text="2단 가운 하[]", command=그림_2단_가운_하)
pic_button5.grid(row=2, column=2, padx=10, pady=7)
#bind_button_to_key(button5, ".")

pic_button6 = tk.Button(pic_pos_frame, text="2단 오른 하[]", command=그림_2단_오른_하)
pic_button6.grid(row=2, column=3, padx=10, pady=7)
#bind_button_to_key(button6, "/")
#endregion
##########################################
#region '문단 찾기' 프레임
para_frame = tk.Frame(tab2, highlightbackground="black", highlightthickness=2, padx=10, pady=10)
para_frame.grid(row=2, column=0, pady=5, sticky="nsew")

position_para = tk.Label(para_frame, text="글머리 패턴-스타일", font=("Arial", 12, "bold"))
position_para.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

#찾을 글머리 패턴 선택 콤보박스스
headlist = ['로마자','로마자.','숫자.','(숫자)','숫자)']
para_headcombo = ttk.Combobox(para_frame, values=headlist)
para_headcombo.grid(row=1, column=1, padx=10, pady=7)
para_headcombo.state(["readonly"])
para_headcombo.set("글머리 패턴  선택")  # 초기값 지정

#찾기 버튼
def 다음_찾기():
    global para
    selected_head = para_headcombo.get()
    regex_map = {
        "로마자": r"[IXV]",
        "로마자.": r"[IXV]+\.",
        "숫자.": r"\d+\.",
        "(숫자)": r"\(\d+\)",
        "숫자)": r"\d+\)",
    }
    regex = regex_map.get(selected_head, r"\d+\)")  # 기본값: 숫자)

    while hwp.set_pos(0, para, 0):
        hwp.MoveSelLineEnd()
        if re.match(regex, hwp.get_selected_text()):
            selected_style = para_stylecombo.get()
            hwp.set_style(f"{selected_style}")  # 2번스타일(개요 1) 적용
            para+=1
            break
        else:
            pass
        hwp.Cancel()
        para+=1
    else:
        # while 루프가 break 없이 종료된 경우 실행
        para = 0  # para 초기화
        position_para_label.config(text=f"탐색 문단={para}")
        return
    # 라벨 업데이트
    position_para_label.config(text=f"탐색 문단={para}")
para_find_btn = tk.Button(para_frame, text="다음 찾기", command=다음_찾기)
para_find_btn.grid(row=1, column=2, padx=10, pady=7)

#탐색 문단 라벨벨
position_para_label = tk.Label(para_frame, text=f"탐색 문단={para}")
position_para_label.grid(row=1, column=3)

#탐색 문단 초기화 버튼
def 문단초기화():
    global para
    para = 0  # 문단 초기화
    position_para_label.config(text=f"탐색 문단={para}")  # 라벨 업데이트
para_reset_btn = tk.Button(para_frame, text="문단 0으로", command=문단초기화)
para_reset_btn.grid(row=2, column=2, padx=10, pady=7)

#적용할 스타일 선택 콤보박스
names = [item['name'] for item in dic]
para_stylecombo = ttk.Combobox(para_frame, values=names)
para_stylecombo.grid(row=2, column=1, padx=10, pady=7)
para_stylecombo.state(["readonly"])
para_stylecombo.set("스타일 선택")  # 초기값 지정



#endregion
##########################################
#region '문자표 바꾸기'프레임
#주의사항 : 홑낫표를 바꾸더라도 겹낫+홑낫 조합의 경우 확인요망
#홑화살괄호의 경우 진짜 등호와 구별을 해야함.
char_frame = tk.Frame(tab2, highlightbackground="black", highlightthickness=2, padx=10, pady=10)
char_frame.grid(row=1, column=1, padx=5, pady=5, sticky="nsew")

find_char = tk.Label(char_frame, text="문자표 찾아바꾸기", font=("Arial", 12, "bold"))
find_char.grid(row=0, column=0, columnspan=2, padx=5, pady=5)


#가운데점 변환 #벡터곱주의
dot_dst = tk.StringVar(value="⋅") #22c5
dot_entry = tk.Entry(char_frame, textvariable=dot_dst, width=2)
dot_entry.grid(row=1, column=0, padx=5, pady=5, sticky="we")
def 가운데점_찾아바꾸기():
    dot_scr = ['·', #00b7
       '·', #0387
       'ּ', #05bc

       '᛫', #16eb
       'ᆞ', #119e

       '∙', # 2219
       '⦁', #2981
       '⋅', #22c5
       '•', #2022
       '․', #2024
       '‧', #2027
       '⸱', # 2E31
       '⸳', #2E33

       '・', #30fb
       'ㆍ', #318d

       'ꞏ', #a78f
       '', #f09e
       '･', #ff65
       ]
    _dst = dot_dst.get()
    _scr2 = f" {_dst} " #가운데 점 양쪽 띄어쓰기
    for i in dot_scr:
        hwp.find_replace_all(i, _dst)
    hwp.find_replace_all(_scr2, _dst) #양쪽 띄어쓰기만 찾아서 띄어쓰기 삭제제
dot_button = tk.Button(char_frame, text="가운데점", command=가운데점_찾아바꾸기)
dot_button.grid(row=1, column=1, padx=10, pady=7, sticky="nsw" )

#공지지
notice_char = tk.Label(char_frame, text="바꾼 후 확인 필수!", font=("Arial", 10))
notice_char.grid(row=0, column=2,  padx=5, pady=5)
notice_char = tk.Label(char_frame, text="키보드<>는 안바뀜", font=("Arial", 10))
notice_char.grid(row=1, column=2,  padx=5, pady=5)

#홑낫표 변환
ssl_dst = tk.StringVar(value="｢") #FF62
ssl_entry = tk.Entry(char_frame, textvariable=ssl_dst, width=2)
ssl_entry.grid(row=2, column=0, padx=2, pady=5, sticky="w")
ssr_dst = tk.StringVar(value="｣") #FF63
ssr_entry = tk.Entry(char_frame, textvariable=ssr_dst, width=2)
ssr_entry.grid(row=2, column=0, padx=2, pady=5, sticky="e")
def 홑낫표_찾아바꾸기():
    ssl_scr = ['｢',#FF62
               '󰠆',#0F0806
               '󰠑',#F0811
               '⌈',#2038
               '⌌',#230C
               '⌜',#231C
               '⎾',#23BE
               '┌',#250C
               '┍',#250D
               '┎',#250E
               '┏',#250F
               '「',#300C
                ]
    ssr_scr = ['｣',#FF63
               '󰠎',#0F080E
               '󰠙',#0F0819
               '⌋',#203B
               '⌏',#230F
               '⌟',#231F
               '⏌',#23CC
               '┘',#2518
               '┙',#2519
               '┚',#251A
               '┛',#251B
               '」',#300D
               ]
    l_dst = ssl_dst.get()
    r_dst = ssr_dst.get()
    for i in ssl_scr:
        hwp.find_replace_all(i, l_dst)
    for i in ssr_scr:
        hwp.find_replace_all(i, r_dst)
ssl_button = tk.Button(char_frame, text="홑낫표", command=홑낫표_찾아바꾸기)
ssl_button.grid(row=2, column=1, padx=10, pady=7, sticky="nsew" )



#겹낫표 변환
dsl_dst = tk.StringVar(value="󰡔") #0F0854
dsl_entry = tk.Entry(char_frame, textvariable=dsl_dst, width=2)
dsl_entry.grid(row=3, column=0, padx=2, pady=5, sticky="w")
dsr_dst = tk.StringVar(value="󰡕") #0F0855
dsr_entry = tk.Entry(char_frame, textvariable=dsr_dst, width=2)
dsr_entry.grid(row=3, column=0, padx=2, pady=5, sticky="e")
def 겹낫표_찾아바꾸기():
    dsl_scr = ['󰡔',#0F0854
               '󰠞',#0F081E
               '╔',#2554
               '『',#300E
                ]
    dsr_scr = ['󰡕', #0F0855
               '󰠦', #0F0826
               '╝', #255D
               '』', #300F
               ]
    l_dst = dsl_dst.get()
    r_dst = dsr_dst.get()
    for i in dsl_scr:
        hwp.find_replace_all(i, l_dst)
    for i in dsr_scr:
        hwp.find_replace_all(i, r_dst)
dsl_button = tk.Button(char_frame, text="겹낫표", command=겹낫표_찾아바꾸기)
dsl_button.grid(row=3, column=1, padx=10, pady=7, sticky="nsew" )



#홑화살괄호 변환
#진짜 등호를 변경할수있으므로 주의할것것
sabl_dst = tk.StringVar(value="〈") #3008
sabl_entry = tk.Entry(char_frame, textvariable=sabl_dst, width=2)
sabl_entry.grid(row=2, column=2, padx=2, pady=5, sticky="w")
sabr_dst = tk.StringVar(value="〉") #3009
sabr_entry = tk.Entry(char_frame, textvariable=sabr_dst, width=2)
sabr_entry.grid(row=2, column=2, padx=(25,5), pady=5, sticky="w")
def 홑화살괄호_찾아바꾸기():
    l_scr = [#'<', # (003C/키보드)#진짜 등호ㅇ도 바꾸면 안되기 땜누에에
             '˂', # (02C2)
             '‹', # (2039)
             '〈', #(2329)
             '❬', #(276C)
             '❮', #(276E)
             '❰', #(2770)
             '⟨', #(27E8)
             '〈', #(3008)
             '＜', #(FF1C)
                ]
    r_scr = [#'>', # (003E/키보드) #진짜 등호ㅇ도 바꾸면 안되기 땜누에에
             '˃', # (02C3)
             '›', # (203A)
             '〉', #(232A)
             '❭', #(276D)
             '❯', #(276F)
             '❱', #(2771)
             '⟩', #(27E9)
             '〉', #(3009)
             '＞', #(FF1E)
               ]
    l_dst = sabl_dst.get()
    r_dst = sabr_dst.get()
    for i in l_scr:
        hwp.find_replace_all(i, l_dst)
    for i in r_scr:
        hwp.find_replace_all(i, r_dst)
sabl_button = tk.Button(char_frame, text="홑꺽쇠(!!)", command=홑화살괄호_찾아바꾸기)
sabl_button.grid(row=2, column=2, padx=5, pady=7, sticky="e" )



#겹화살괄호 변환
dabl_dst = tk.StringVar(value="《") #300A  
dabl_entry = tk.Entry(char_frame, textvariable=dabl_dst, width=2)
dabl_entry.grid(row=3, column=2, padx=2, pady=5, sticky="w")
dabr_dst = tk.StringVar(value="》") #300B 
dabr_entry = tk.Entry(char_frame, textvariable=dabr_dst, width=2)
dabr_entry.grid(row=3, column=2, padx=(25,5), pady=5, sticky="w")
def 겹화살괄호_찾아바꾸기():
    l_scr = ['《', #300A
             '«', #00AB
             '≪', #226A
             '⪡', #2AA1
             '⟪', #27EA
                ]
    r_scr = ['》', #300B
             '»', #00BB
             '≫', #226B
             '⪢', #2AA2
             '⟫', #27EB
               ]
    l_dst = dabl_dst.get()
    r_dst = dabr_dst.get()
    for i in l_scr:
        hwp.find_replace_all(i, l_dst)
    for i in r_scr:
        hwp.find_replace_all(i, r_dst)
sabl_button = tk.Button(char_frame, text="겹꺽쇠", command=겹화살괄호_찾아바꾸기)
sabl_button.grid(row=3, column=2, padx=5, pady=7, sticky="e" )

##########################################
#endregion
##########################################
#endregion

#region 세 번째 탭 : 미정
tab3 = ttk.Frame(notebook)
notebook.add(tab3, text="추가 기능 2")

frame3 = tk.Frame(tab3, padx=10, pady=10)
frame3.grid(row=0, column=0, padx=15, pady=5, sticky="ew")

##############################################################
#endregion


#region UI 출력
root.update_idletasks()
root.geometry(f"{max(200, 15+ notebook.winfo_reqwidth())}x{max(300, notebook.winfo_reqheight() + 180)}")

# Tkinter 메인 루프 실행
root.mainloop()
#endregion








