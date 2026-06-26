"""Generate ZEGO manual PDFs for all 5 languages using reportlab + CJK fonts."""
import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER

FONTS_DIR = "C:/Windows/Fonts/"

def reg(name, file):
    try:
        pdfmetrics.registerFont(TTFont(name, FONTS_DIR + file))
        return True
    except Exception as e:
        print(f"  Font {name} failed: {e}")
        return False

reg("Malgun",    "malgun.ttf")
reg("MalgunB",   "malgunbd.ttf")
reg("Gothic",    "msgothic.ttc")
reg("SimSun",    "simsun.ttc")
reg("JhengHei",  "msjh.ttc")

BRAND = colors.HexColor("#CC1625")
BRAND_LIGHT = colors.HexColor("#FFF1F2")
GRAY_BG = colors.HexColor("#FAF9F7")
BORDER = colors.HexColor("#E7E5E4")
TEXT = colors.HexColor("#1C1917")
TEXT_MUTED = colors.HexColor("#57534E")
INFO_BG = colors.HexColor("#EFF6FF")
INFO_TEXT = colors.HexColor("#1D4ED8")
WARN_BG = colors.HexColor("#FFFBEB")
WARN_TEXT = colors.HexColor("#92400E")
TIP_BG = colors.HexColor("#F0FDF4")
TIP_TEXT = colors.HexColor("#15803D")
WHITE = colors.white

W, H = A4

def make_styles(font, fontb):
    styles = {
        "body": ParagraphStyle("body", fontName=font, fontSize=9, leading=14, textColor=TEXT, spaceAfter=4),
        "h1":   ParagraphStyle("h1",   fontName=fontb, fontSize=20, leading=24, textColor=WHITE),
        "h2":   ParagraphStyle("h2",   fontName=fontb, fontSize=11.5, leading=15, textColor=BRAND, spaceBefore=10, spaceAfter=6),
        "h3":   ParagraphStyle("h3",   fontName=fontb, fontSize=9.5, leading=13, textColor=TEXT, spaceBefore=8, spaceAfter=4),
        "info": ParagraphStyle("info", fontName=font, fontSize=8.5, leading=13, textColor=INFO_TEXT, backColor=INFO_BG),
        "warn": ParagraphStyle("warn", fontName=font, fontSize=8.5, leading=13, textColor=WARN_TEXT, backColor=WARN_BG),
        "tip":  ParagraphStyle("tip",  fontName=font, fontSize=8.5, leading=13, textColor=TIP_TEXT,  backColor=TIP_BG),
        "toc":  ParagraphStyle("toc",  fontName=font, fontSize=8.5, leading=13, textColor=TEXT_MUTED),
        "meta": ParagraphStyle("meta", fontName=font, fontSize=8,   leading=12, textColor=WHITE),
        "foot": ParagraphStyle("foot", fontName=font, fontSize=8,   leading=12, textColor=WARN_TEXT, alignment=TA_CENTER),
        "step": ParagraphStyle("step", fontName=font, fontSize=9,   leading=14, textColor=TEXT, leftIndent=18),
        "li":   ParagraphStyle("li",   fontName=font, fontSize=9,   leading=14, textColor=TEXT, leftIndent=12, bulletIndent=4),
    }
    return styles

def cover_table(title, sub, meta, cover_color):
    cover = Table(
        [[Paragraph(f"✈  ZEGO", ParagraphStyle("logo", fontName="MalgunB", fontSize=18, textColor=WHITE)),
          ""],
         [Paragraph(title, ParagraphStyle("cvt", fontName="MalgunB", fontSize=22, textColor=WHITE, leading=26)), ""],
         [Paragraph(sub,   ParagraphStyle("cvs", fontName="Malgun", fontSize=10, textColor=colors.HexColor("#FFD0D4"), leading=14)), ""],
         [Paragraph(meta,  ParagraphStyle("cvm", fontName="Malgun", fontSize=8, textColor=colors.HexColor("#FFD0D4"))), ""],
        ],
        colWidths=[W - 80*mm, 0],
        rowHeights=[14*mm, 18*mm, 10*mm, 8*mm],
    )
    cover.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), cover_color),
        ("TOPPADDING",    (0,0), (-1,0), 12),
        ("BOTTOMPADDING", (0,-1), (-1,-1), 12),
        ("LEFTPADDING",   (0,0), (-1,-1), 14),
        ("RIGHTPADDING",  (0,0), (-1,-1), 14),
        ("LINEBELOW",     (0,2), (-1,2), 0.5, colors.HexColor("#FF8B95")),
        ("TOPPADDING",    (0,3), (-1,3), 6),
    ]))
    return cover

def section_title(text, s):
    return Table(
        [[Paragraph(text, s["h2"])]],
        colWidths=[W - 50*mm],
        style=TableStyle([
            ("LEFTPADDING",  (0,0), (-1,-1), 8),
            ("RIGHTPADDING", (0,0), (-1,-1), 8),
            ("TOPPADDING",   (0,0), (-1,-1), 4),
            ("BOTTOMPADDING",(0,0), (-1,-1), 4),
            ("LINEBEFORE",   (0,0), (0,-1), 4, BRAND),
            ("BACKGROUND",   (0,0), (-1,-1), BRAND_LIGHT),
            ("ROUNDEDCORNERS", [0, 0, 4, 4]),
        ])
    )

def info_box(text, s, style="info"):
    st = s[style]
    icons = {"info": "ℹ", "warn": "⚠", "tip": "💡"}
    bgs   = {"info": INFO_BG, "warn": WARN_BG, "tip": TIP_BG}
    borders = {"info": colors.HexColor("#BFDBFE"), "warn": colors.HexColor("#FDE68A"), "tip": colors.HexColor("#BBF7D0")}
    t = Table(
        [[Paragraph(f"{icons[style]} {text}", st)]],
        colWidths=[W - 50*mm],
        style=TableStyle([
            ("BACKGROUND",   (0,0), (-1,-1), bgs[style]),
            ("BOX",          (0,0), (-1,-1), 0.8, borders[style]),
            ("LEFTPADDING",  (0,0), (-1,-1), 10),
            ("RIGHTPADDING", (0,0), (-1,-1), 10),
            ("TOPPADDING",   (0,0), (-1,-1), 6),
            ("BOTTOMPADDING",(0,0), (-1,-1), 6),
            ("ROUNDEDCORNERS", [4, 4, 4, 4]),
        ])
    )
    return t

def data_table(headers, rows, s, font, fontb):
    th_style = ParagraphStyle("th", fontName=fontb, fontSize=8.5, textColor=WHITE)
    td_style = ParagraphStyle("td", fontName=font,  fontSize=8.5, textColor=TEXT, leading=13)
    table_data = [[Paragraph(h, th_style) for h in headers]]
    for row in rows:
        table_data.append([Paragraph(str(c), td_style) for c in row])
    col_w = (W - 50*mm) / len(headers)
    t = Table(table_data, colWidths=[col_w]*len(headers))
    ts = [
        ("BACKGROUND",   (0,0), (-1,0), BRAND),
        ("ROWBACKGROUNDS",(0,1), (-1,-1), [WHITE, GRAY_BG]),
        ("GRID",         (0,0), (-1,-1), 0.4, BORDER),
        ("LEFTPADDING",  (0,0), (-1,-1), 8),
        ("RIGHTPADDING", (0,0), (-1,-1), 8),
        ("TOPPADDING",   (0,0), (-1,-1), 5),
        ("BOTTOMPADDING",(0,0), (-1,-1), 5),
        ("VALIGN",       (0,0), (-1,-1), "MIDDLE"),
    ]
    t.setStyle(TableStyle(ts))
    return t

def step(num, text, s):
    num_cell = Table(
        [[Paragraph(str(num), ParagraphStyle("sn", fontName="MalgunB", fontSize=8, textColor=WHITE, alignment=TA_CENTER))]],
        colWidths=[6*mm], rowHeights=[6*mm],
        style=TableStyle([("BACKGROUND",(0,0),(-1,-1),BRAND),("ROUNDEDCORNERS",[3,3,3,3]),
                          ("TOPPADDING",(0,0),(-1,-1),0),("LEFTPADDING",(0,0),(-1,-1),0)])
    )
    return Table(
        [[num_cell, Paragraph(text, s["body"])]],
        colWidths=[8*mm, W-58*mm],
        style=TableStyle([("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                          ("LEFTPADDING",(0,0),(0,-1),0),("LEFTPADDING",(1,0),(1,-1),6),
                          ("TOPPADDING",(0,0),(-1,-1),2),("BOTTOMPADDING",(0,0),(-1,-1),2)])
    )

def menu_tree(items, s):
    rows = []
    for item in items:
        if item.get("section"):
            rows.append([Paragraph(item["label"].upper(),
                ParagraphStyle("ms", fontName="MalgunB", fontSize=7, textColor=TEXT_MUTED,
                               letterSpacing=0.8))])
        else:
            rows.append([Paragraph(f"  {item['label']}",
                ParagraphStyle("mi", fontName="Malgun", fontSize=8.5, textColor=TEXT, leading=13))])
    t = Table(rows, colWidths=[W-50*mm],
              style=TableStyle([
                  ("BACKGROUND",(0,0),(-1,-1),GRAY_BG),
                  ("BOX",(0,0),(-1,-1),0.5,BORDER),
                  ("LEFTPADDING",(0,0),(-1,-1),10),
                  ("TOPPADDING",(0,0),(-1,-1),2),
                  ("BOTTOMPADDING",(0,0),(-1,-1),2),
              ]))
    return t

def footer_box(text, s):
    t = Table(
        [[Paragraph(text, s["foot"])]],
        colWidths=[W - 50*mm],
        style=TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),WARN_BG),
            ("BOX",(0,0),(-1,-1),0.5,colors.HexColor("#FDE68A")),
            ("LEFTPADDING",(0,0),(-1,-1),12),
            ("RIGHTPADDING",(0,0),(-1,-1),12),
            ("TOPPADDING",(0,0),(-1,-1),8),
            ("BOTTOMPADDING",(0,0),(-1,-1),8),
            ("ROUNDEDCORNERS",[4,4,4,4]),
        ])
    )
    return t

# ─── CONTENT PER LANGUAGE ────────────────────────────────────────────────────

LANGS = {
    "kor": {
        "font": "Malgun", "fontb": "MalgunB",
        "cover_color": BRAND,
        "title": "사용자 매뉴얼",
        "sub": "이스타항공 물류/비품 관리 시스템 — 전 기능 안내서",
        "meta": "대상: 전 지점 담당자 · 관리자  |  언어: 한국어  |  2026년 6월 기준",
        "footer": "문의사항은 공항서비스팀으로 연락 바랍니다.  |  hdqapo@eastarjet.com",
        "toc_title": "목차",
        "sections": [
            {"title": "1. 시스템 소개",
             "body": [
                ("p", "ZEGO는 이스타항공 전 지점의 운송양식 및 비품 재고를 통합 관리하는 웹 기반 시스템입니다."),
                ("h3", "관리 양식 종류"),
                ("table", (["분류","양식명"],[
                    ["승객 서류","악기 서약서 (DECLARATION OF INDEMNITY, Musical Instrument)"],
                    ["승객 서류","보호자 서약서 (DECLARATION OF PARENT GUARDIAN)"],
                    ["화물 서류","총기인수인계서 (Firearm Handover Form)"],
                    ["특수화물","NOTOC (Notification to Captain)"],
                    ["일반 비품","X-배너, 스탠션, 안내문, 탑승권 등"],
                ])),
                ("h3", "재고 상태"),
                ("table", (["색상","상태","의미"],[
                    ["녹색","정상 (Normal)","최소기준 초과, 재고 충분"],
                    ["노란색","부족 (Low)","수량>0, 최소기준 이하"],
                    ["빨간색","소진 (Empty)","수량=0"],
                ])),
             ]},
            {"title": "2. 로그인 & 비밀번호",
             "body": [
                ("step", (1, "브라우저에서 ZEGO 사이트 주소 접속")),
                ("step", (2, "아이디와 비밀번호 입력 후 로그인 버튼 클릭")),
                ("warn", "초기 비밀번호: $eastar12 — 첫 로그인 시 반드시 변경하세요."),
                ("p", "좌측 사이드바 하단 '비밀번호 변경' 메뉴에서 언제든 변경 가능합니다."),
                ("info", "비밀번호 만료 7일 전부터 상단에 경고 메시지가 표시됩니다."),
             ]},
            {"title": "3. 대시보드",
             "body": [
                ("p", "로그인 후 첫 화면. 통계 카드, 재고 경고 테이블, 최근 입출고 이력, 6개월 출고 트렌드 차트로 구성됩니다."),
                ("table", (["통계 카드","내용"],[
                    ["운영 지점 수","현재 등록된 전체 지점 수"],
                    ["양식 종류","관리 중인 양식 품목 수"],
                    ["당일 거래","오늘 처리된 입고·출고·이전 건수"],
                    ["재고 경고","부족 또는 소진 상태 항목 수"],
                ])),
             ]},
            {"title": "4. 재고현황",
             "body": [
                ("p", "좌측 메뉴 '재고현황' 클릭. 4개 탭으로 구성됩니다."),
                ("h3", "탭 1 — 지점별 상세"),
                ("p", "모든 지점·양식의 현재 재고를 테이블로 표시. 수량 옆 입력창에 숫자 입력 후 Enter를 누르면 해당 지점의 최소기준만 저장됩니다."),
                ("h3", "탭 2 — 지점별 보기 (관리자 전용)"),
                ("p", "각 지점을 아코디언 패널로 펼쳐서 확인. 지점별로 최소기준을 독립 설정 가능."),
                ("h3", "탭 3 — 양식별 누적 합계"),
                ("p", "양식 종류별 전체 지점 합계 재고. 지점 수 버튼 클릭 시 지점별 상세 팝업."),
                ("h3", "탭 4 — 월별 소비 현황"),
                ("p", "최근 6개월간 양식별 출고 수량 추이."),
                ("tip", "최소기준은 지점별로 독립 관리됩니다. 한 지점의 최소기준을 변경해도 다른 지점에 영향을 주지 않습니다."),
             ]},
            {"title": "5. 입고 처리",
             "body": [
                ("step", (1, "좌측 메뉴 '입고 처리' 클릭")),
                ("step", (2, "지점, 양식 종류, 수량, 날짜, 비고 입력")),
                ("step", (3, "입고 처리 버튼 클릭")),
                ("info", "입고 후에도 최소기준 이하이면 자동으로 알림이 생성됩니다."),
             ]},
            {"title": "6. 출고 처리",
             "body": [
                ("step", (1, "좌측 메뉴 '출고 처리' 클릭")),
                ("step", (2, "지점, 양식 종류, 수량, 날짜, 비고 입력")),
                ("step", (3, "출고 처리 버튼 클릭")),
                ("warn", "현재 재고보다 많은 수량은 출고할 수 없습니다. 출고 후 최소기준 이하이면 자동 알림이 발송됩니다."),
                ("h3", "출고 취소"),
                ("p", "잘못 처리된 출고 건을 취소하면 재고가 자동으로 복구됩니다."),
                ("step", (1, "좌측 메뉴 '거래 내역' 클릭")),
                ("step", (2, "취소할 출고 건의 되돌리기(↩) 버튼 클릭")),
                ("step", (3, "확인 팝업에서 '취소 확정' 클릭 → 재고 즉시 복구")),
                ("info", "취소된 기록은 회색+취소선+취소됨 배지로 표시. 재취소·수정 불가. 본인 지점 또는 관리자만 취소 가능."),
             ]},
            {"title": "7. 일괄 출고",
             "body": [
                ("step", (1, "좌측 메뉴 '일괄 출고' 클릭")),
                ("step", (2, "지점 선택 후 해당 지점의 양식 목록 표시")),
                ("step", (3, "각 양식의 출고 수량 입력 (0이면 건너뜀)")),
                ("step", (4, "날짜, 공통 비고 입력 후 '일괄 출고' 버튼 클릭")),
                ("tip", "여러 양식을 한 번에 출고 처리하여 시간을 절약할 수 있습니다."),
             ]},
            {"title": "8. 지점 간 이전",
             "body": [
                ("p", "재고가 부족한 지점에서 여유가 있는 지점에 이전을 요청할 때 사용합니다."),
                ("step", (1, "좌측 메뉴 '지점 간 이전' 클릭")),
                ("step", (2, "도착 지점, 양식, 수량, 날짜 입력")),
                ("step", (3, "이전 신청 버튼 클릭 → 승인 대기 상태")),
                ("info", "이전 신청을 받은 지점에는 메일로 안내됩니다. ZEGO 사이트에 접속하여 처리해주세요. 처리가 완료되면 이전 신청을 요청한 지점에 처리 결과가 메일로 안내됩니다."),
                ("h3", "이전 상태 흐름"),
                ("table", (["상태","의미"],[
                    ["대기 (Pending)","관리자 승인 전"],
                    ["승인 (Approved)","관리자가 승인, 받는 지점의 확인 대기"],
                    ["확인 (Confirmed)","이전 완료, 재고 반영"],
                    ["거부 (Rejected)","관리자가 거부"],
                ])),
             ]},
            {"title": "9. 거래 내역 & 보고서",
             "body": [
                ("p", "모든 입고·출고·이전 내역 조회. 지점, 양식, 유형, 기간으로 필터링 가능."),
                ("p", "보고서: 지점별 재고 현황, 월별 입출고 비교, 양식별 상위 출고, 지점별 재고 차트."),
                ("p", "수요예측: 과거 출고 패턴 분석으로 향후 재고 소진 예상일 계산."),
             ]},
            {"title": "10. 카탈로그(운송아이템) 신청",
             "body": [
                ("step", (1, "좌측 메뉴 '카탈로그' 클릭 — 품목 목록 확인")),
                ("step", (2, "원하는 품목에서 수량 입력 후 '담기' 버튼 클릭")),
                ("step", (3, "'신청 장바구니'에서 최종 확인 후 신청 완료")),
                ("step", (4, "'신청 내역'에서 진행 상황 확인")),
                ("info", "관리자가 설정한 신청 기간 내에만 신청 가능합니다."),
             ]},
            {"title": "11. 양식 비품 신청",
             "body": [
                ("step", (1, "좌측 메뉴 '양식 신청' 클릭")),
                ("step", (2, "양식 종류, 신청 수량, 비고 입력")),
                ("step", (3, "신청 버튼 클릭")),
                ("step", (4, "'나의 신청 내역'에서 처리 상황 확인")),
                ("info", "관리자가 설정한 신청 기간 내에만 신청 가능합니다."),
             ]},
            {"title": "12. 관리자 기능",
             "body": [
                ("h3", "사용자 관리"),
                ("p", "계정 생성·수정·삭제, 역할(관리자/지점) 설정, 지점 배정, 비밀번호 초기화."),
                ("h3", "신청 기간 설정"),
                ("p", "'운송양식' 탭: 양식 비품 신청 기간 설정. '운송아이템' 탭: 카탈로그 신청 기간 설정."),
                ("h3", "알림 배지"),
                ("table", (["메뉴","배지 의미"],[
                    ["지점 간 이전","승인 대기 중인 이전 신청 건수"],
                    ["양식 신청 관리 (관리자)","미처리 양식 신청 건수"],
                    ["신청 내역 (관리자)","미처리 카탈로그 신청 건수"],
                ])),
             ]},
        ],
    },

    "eng": {
        "font": "Malgun", "fontb": "MalgunB",
        "cover_color": colors.HexColor("#1a56db"),
        "title": "User Manual",
        "sub": "Eastar Jet Logistics & Supply Management System",
        "meta": "For: All Branch Staff & Administrators  |  Language: English  |  June 2026",
        "footer": "For inquiries, please contact the Airport Services Team.  |  hdqapo@eastarjet.com",
        "toc_title": "Table of Contents",
        "sections": [
            {"title": "1. System Overview",
             "body": [
                ("p", "ZEGO is a web-based system for integrated management of transport forms and supplies across all Eastar Jet branches."),
                ("h3", "Managed Form Types"),
                ("table", (["Category","Form Name"],[
                    ["Passenger Docs","Declaration of Indemnity (Musical Instrument)"],
                    ["Passenger Docs","Declaration of Parent/Guardian"],
                    ["Cargo Docs","Firearm Handover Form"],
                    ["Special Cargo","NOTOC (Notification to Captain)"],
                    ["General Supplies","X-banners, stanchions, notices, boarding passes, etc."],
                ])),
                ("h3", "Stock Status"),
                ("table", (["Color","Status","Meaning"],[
                    ["Green","Normal","Above minimum threshold"],
                    ["Yellow","Low","Qty > 0, at/below minimum"],
                    ["Red","Empty","Qty = 0"],
                ])),
             ]},
            {"title": "2. Login & Password",
             "body": [
                ("step", (1, "Open the ZEGO website in your browser")),
                ("step", (2, "Enter your ID and password, then click Login")),
                ("warn", "Initial password: $eastar12 — You must change it on first login."),
                ("info", "A warning banner appears 7 days before your password expires."),
             ]},
            {"title": "3. Dashboard",
             "body": [
                ("p", "First screen after login. Shows stat cards, stock alert tables, recent transactions, and a 6-month outbound trend chart."),
             ]},
            {"title": "4. Inventory Status",
             "body": [
                ("p", "Click Inventory in the left menu. There are 4 tabs."),
                ("h3", "Tab 1 — Branch Detail"),
                ("p", "Shows current stock for all branch-form combinations. Edit minimum threshold by entering a number and pressing Enter — saves for that branch only."),
                ("h3", "Tab 2 — Branch View (Admin only)"),
                ("p", "Accordion panels per branch. Minimum thresholds can be set per-branch independently."),
                ("h3", "Tab 3 — Form Summary"),
                ("p", "Total stock grouped by form type. Click branch count for per-branch detail popup."),
                ("h3", "Tab 4 — Monthly Consumption"),
                ("p", "Monthly outbound quantity trend for the past 6 months."),
                ("tip", "Minimum thresholds are managed per branch. Changing one branch's threshold does not affect others."),
             ]},
            {"title": "5. Inbound Processing",
             "body": [
                ("step", (1, "Click Inbound in the left menu")),
                ("step", (2, "Enter branch, form type, quantity, date, notes")),
                ("step", (3, "Click Process Inbound")),
                ("info", "If stock remains below minimum threshold after inbound, an alert is automatically generated."),
             ]},
            {"title": "6. Outbound Processing",
             "body": [
                ("step", (1, "Click Outbound in the left menu")),
                ("step", (2, "Enter branch, form type, quantity, date, notes")),
                ("step", (3, "Click Process Outbound")),
                ("warn", "Cannot outbound more than current stock. An alert is sent automatically if stock falls below minimum threshold."),
                ("h3", "Cancelling an Outbound"),
                ("p", "If an outbound was processed by mistake, cancel it to automatically restore inventory."),
                ("step", (1, "Click Transaction History in the left menu")),
                ("step", (2, "Click the Undo (undo) button next to the outbound record")),
                ("step", (3, "Click Confirm Cancel in the popup -> inventory is restored immediately")),
                ("info", "Cancelled records appear greyed out with strikethrough and Cancelled badge. Cannot be re-cancelled or edited. Only the processing branch or admin can cancel."),
             ]},
            {"title": "7. Bulk Outbound",
             "body": [
                ("step", (1, "Click Bulk Outbound in the left menu")),
                ("step", (2, "Select branch — form list appears")),
                ("step", (3, "Enter outbound quantity for each form (0 to skip)")),
                ("step", (4, "Enter date and notes, then click Bulk Outbound")),
                ("tip", "Process multiple forms simultaneously in a single submission."),
             ]},
            {"title": "8. Branch-to-Branch Transfer",
             "body": [
                ("p", "Used when a branch with insufficient stock requests forms from a branch that has surplus."),
                ("step", (1, "Click Branch Transfer in the left menu")),
                ("step", (2, "Select destination branch, form, quantity, date")),
                ("step", (3, "Click Request Transfer → status becomes Pending")),
                ("info", "The receiving branch will be notified by email. Please log in to the ZEGO site and process the request. Once processed, the requesting branch will receive the result by email."),
                ("h3", "Transfer Status Flow"),
                ("table", (["Status","Meaning"],[
                    ["Pending","Awaiting admin approval"],
                    ["Approved","Admin approved; awaiting receiving branch confirmation"],
                    ["Confirmed","Transfer complete; inventory updated"],
                    ["Rejected","Rejected by admin"],
                ])),
             ]},
            {"title": "9. Transaction History & Reports",
             "body": [
                ("p", "View all inbound, outbound, and transfer records. Filter by branch, form type, transaction type, and date range."),
                ("p", "Reports: branch stock status charts, monthly comparison, top outbound forms, branch quantity chart."),
                ("p", "Demand Forecasting: estimates future depletion dates based on historical outbound patterns."),
             ]},
            {"title": "10. Catalog (Transport Items) Request",
             "body": [
                ("step", (1, "Click Catalog in the left menu")),
                ("step", (2, "Enter quantity and click Add to Cart")),
                ("step", (3, "Review in Request Cart and click Submit Request")),
                ("step", (4, "Check progress in Request History")),
                ("info", "Requests are only possible during the request period set by the administrator."),
             ]},
            {"title": "11. Form Supply Request",
             "body": [
                ("step", (1, "Click Form Request in the left menu")),
                ("step", (2, "Select form type, enter quantity and notes")),
                ("step", (3, "Click Submit")),
                ("step", (4, "Track progress in My Requests")),
                ("info", "Form supply requests are also subject to the request period configured by the administrator."),
             ]},
            {"title": "12. Admin Functions",
             "body": [
                ("h3", "User Management"),
                ("p", "Create, edit, delete accounts; set roles (Admin/Branch); assign branches; reset passwords."),
                ("h3", "Request Period Settings"),
                ("p", "Transport Forms tab: set form supply request period. Transport Items tab: set catalog request period."),
                ("h3", "Notification Badges"),
                ("table", (["Menu","Badge Meaning"],[
                    ["Branch Transfer","Number of pending transfer requests"],
                    ["Form Request Mgmt (Admin)","Number of unprocessed form requests"],
                    ["Request History (Admin)","Number of unprocessed catalog requests"],
                ])),
             ]},
        ],
    },

    "jpn": {
        "font": "Gothic", "fontb": "Gothic",
        "cover_color": colors.HexColor("#e84393"),
        "title": "ユーザーマニュアル",
        "sub": "イースター航空 物流・備品管理システム",
        "meta": "対象：全支店スタッフ・管理者  |  言語：日本語  |  2026年6月現在",
        "footer": "お問い合わせは空港サービスチームまで。  |  hdqapo@eastarjet.com",
        "toc_title": "目次",
        "sections": [
            {"title": "1. システム概要",
             "body": [
                ("p", "ZEGOはイースター航空の全支店における輸送様式・備品在庫を一元管理するWebシステムです。"),
                ("h3", "管理様式の種類"),
                ("table", (["区分","様式名"],[
                    ["旅客書類","楽器誓約書（DECLARATION OF INDEMNITY, Musical Instrument）"],
                    ["旅客書類","保護者誓約書（DECLARATION OF PARENT GUARDIAN）"],
                    ["貨物書類","銃器引渡書（Firearm Handover Form）"],
                    ["特殊貨物","NOTOC（Notification to Captain）"],
                    ["一般備品","Xバナー、スタンション、案内文、搭乗券等"],
                ])),
             ]},
            {"title": "2. ログイン・パスワード",
             "body": [
                ("step", (1, "ブラウザでZEGOサイトにアクセス")),
                ("step", (2, "IDとパスワードを入力しログインボタンをクリック")),
                ("warn", "初期パスワード：$eastar12 — 初回ログイン時に必ず変更してください。"),
                ("info", "パスワード期限の7日前から上部に警告が表示されます。"),
             ]},
            {"title": "3. ダッシュボード",
             "body": [
                ("p", "ログイン後の最初の画面。統計カード、在庫警告テーブル、最近の入出庫履歴、6ヶ月出庫トレンドチャートで構成されます。"),
             ]},
            {"title": "4. 在庫状況",
             "body": [
                ("p", "左メニュー「在庫状況」をクリック。4つのタブで構成されています。"),
                ("h3", "タブ1 — 支店別詳細"),
                ("p", "全支店・様式の現在庫をテーブル表示。数値入力後Enterで該当支店のみ最低基準を保存。"),
                ("h3", "タブ2 — 支店別表示（管理者専用）"),
                ("p", "各支店をアコーディオンパネルで展開。支店ごとに最低基準を独立設定可能。"),
                ("h3", "タブ3 — 様式別集計"),
                ("p", "様式種類別に全支店合計在庫を確認。"),
                ("h3", "タブ4 — 月別消費状況"),
                ("p", "直近6ヶ月間の様式別出庫数量の推移。"),
                ("tip", "最低基準は支店ごとに独立して管理されます。一支店の変更は他支店に影響しません。"),
             ]},
            {"title": "5. 入庫処理",
             "body": [
                ("step", (1, "左メニュー「入庫処理」をクリック")),
                ("step", (2, "支店、様式種類、数量、日付、備考を入力")),
                ("step", (3, "「入庫処理」ボタンをクリック")),
                ("info", "入庫後も最低基準以下の場合、自動的にアラートが生成されます。"),
             ]},
            {"title": "6. 出庫処理",
             "body": [
                ("step", (1, "左メニュー「出庫処理」をクリック")),
                ("step", (2, "支店、様式種類、数量、日付、備考を入力")),
                ("step", (3, "「出庫処理」ボタンをクリック")),
                ("warn", "現在庫より多い数量は出庫できません。出庫後に最低基準以下になると自動通知が送信されます。"),
                ("h3", "出庫取消"),
                ("p", "誤って処理した出庫を取り消すと、在庫が自動的に復元されます。"),
                ("step", (1, "左メニュー「取引履歴」をクリック")),
                ("step", (2, "取消したい出庫レコードの元に戻す(↩)ボタンをクリック")),
                ("step", (3, "確認ポップアップで「取消確定」をクリック → 在庫が即時復元")),
                ("info", "取消済みレコードはグレーアウト・取消線・取消済みバッジで表示。再取消・編集不可。処理支店または管理者のみ操作可能。"),
             ]},
            {"title": "7. 一括出庫",
             "body": [
                ("step", (1, "左メニュー「一括出庫」をクリック")),
                ("step", (2, "支店を選択すると様式リストが表示")),
                ("step", (3, "各様式の出庫数量を入力（0は無視）")),
                ("step", (4, "日付と備考を入力し「一括出庫」をクリック")),
                ("tip", "複数の様式を一度に出庫処理でき、時間を節約できます。"),
             ]},
            {"title": "8. 支店間移動",
             "body": [
                ("p", "在庫が不足している支店から、余裕のある支店へ不足分の補充を依頼する際に使用します。"),
                ("step", (1, "左メニュー「支店間移動」をクリック")),
                ("step", (2, "送り先支店、様式、数量、日付を入力")),
                ("step", (3, "「移動申請」ボタンをクリック → 承認待ち状態")),
                ("info", "移動申請を受けた支店にはメールで通知されます。ZEGOサイトにログインして処理してください。処理が完了すると、申請した支店に処理結果がメールで通知されます。"),
                ("h3", "移動ステータス"),
                ("table", (["ステータス","意味"],[
                    ["待機（Pending）","管理者承認前"],
                    ["承認（Approved）","管理者が承認、受け取り支店の確認待ち"],
                    ["確認（Confirmed）","移動完了、在庫反映"],
                    ["却下（Rejected）","管理者が却下"],
                ])),
             ]},
            {"title": "9. 取引履歴・レポート",
             "body": [
                ("p", "全ての入庫・出庫・移動履歴を照会。支店、様式、取引タイプ、期間でフィルタリング可能。"),
                ("p", "レポート：支店別在庫状況、月別比較、様式別出庫ランキング、需要予測。"),
             ]},
            {"title": "10. カタログ申請",
             "body": [
                ("step", (1, "左メニュー「カタログ」をクリック")),
                ("step", (2, "数量を入力し「カートに追加」")),
                ("step", (3, "「申請カート」で確認し「申請完了」")),
                ("step", (4, "「申請履歴」で進捗確認")),
                ("info", "管理者が設定した申請期間内のみ申請可能です。"),
             ]},
            {"title": "11. 様式備品申請",
             "body": [
                ("step", (1, "左メニュー「様式申請」をクリック")),
                ("step", (2, "様式種類、申請数量、備考を入力")),
                ("step", (3, "「申請」ボタンをクリック")),
                ("step", (4, "「私の申請履歴」で処理状況を確認")),
                ("info", "管理者が設定した申請期間内のみ申請可能です。"),
             ]},
            {"title": "12. 管理者機能",
             "body": [
                ("h3", "ユーザー管理"),
                ("p", "アカウントの作成・編集・削除、役割設定、支店割り当て、パスワードリセット。"),
                ("h3", "申請期間設定"),
                ("p", "輸送様式タブ：様式備品申請期間の設定。輸送アイテムタブ：カタログ申請期間の設定。"),
             ]},
        ],
    },

    "chn": {
        "font": "SimSun", "fontb": "SimSun",
        "cover_color": colors.HexColor("#d97706"),
        "title": "用户手册",
        "sub": "易斯达航空 物流·备品管理系统",
        "meta": "适用对象：全部网点工作人员·管理员  |  语言：中文（简体）|  2026年6月版",
        "footer": "如有疑问，请联系机场服务团队。  |  hdqapo@eastarjet.com",
        "toc_title": "目录",
        "sections": [
            {"title": "1. 系统概述",
             "body": [
                ("p", "ZEGO是一个基于Web的系统，用于统一管理易斯达航空各网点的运输表单及备品库存。"),
                ("h3", "管理表单种类"),
                ("table", (["分类","表单名称"],[
                    ["旅客文件","乐器免责声明书（DECLARATION OF INDEMNITY, Musical Instrument）"],
                    ["旅客文件","监护人声明书（DECLARATION OF PARENT GUARDIAN）"],
                    ["货物文件","枪支交接书（Firearm Handover Form）"],
                    ["特种货物","NOTOC（危险货物通知书）"],
                    ["一般备品","X形展架、隔离带、告示文、登机牌等"],
                ])),
             ]},
            {"title": "2. 登录与密码",
             "body": [
                ("step", (1, "在浏览器中打开ZEGO网站")),
                ("step", (2, "输入账号和密码，点击登录按钮")),
                ("warn", "初始密码：$eastar12 — 首次登录时必须修改密码。"),
                ("info", "密码到期前7天上方将显示警告提示，请及时修改。"),
             ]},
            {"title": "3. 仪表盘",
             "body": [
                ("p", "登录后的第一个页面，显示统计卡片、库存预警表格、近期交易记录及6个月出库趋势图。"),
             ]},
            {"title": "4. 库存状况",
             "body": [
                ("p", "点击左侧菜单的「库存状况」，共有4个标签页。"),
                ("h3", "标签1 — 网点详细"),
                ("p", "显示所有网点·表单的当前库存。在输入框输入数字后按Enter，仅保存该网点的最低标准。"),
                ("h3", "标签2 — 按网点查看（管理员）"),
                ("p", "以折叠面板展开各网点。最低标准可按网点分别独立设置。"),
                ("h3", "标签3 — 按表单汇总"),
                ("p", "按表单种类查看全部网点合计库存。"),
                ("h3", "标签4 — 月度消耗状况"),
                ("p", "查看近6个月各表单出库数量变化趋势。"),
                ("tip", "最低标准按网点独立管理。修改某一网点的最低标准，不会影响其他网点。"),
             ]},
            {"title": "5. 入库处理",
             "body": [
                ("step", (1, "点击左侧菜单「入库处理」")),
                ("step", (2, "选择网点，输入表单种类、数量、日期、备注")),
                ("step", (3, "点击「入库处理」按钮")),
                ("info", "入库后若仍低于最低标准，系统将自动生成警报通知。"),
             ]},
            {"title": "6. 出库处理",
             "body": [
                ("step", (1, "点击左侧菜单「出库处理」")),
                ("step", (2, "输入网点、表单种类、数量、日期、备注")),
                ("step", (3, "点击「出库处理」按钮")),
                ("warn", "出库数量不能超过当前库存。出库后若库存降至最低标准以下，将自动发送通知。"),
                ("h3", "出库取消"),
                ("p", "如果出库处理有误，可以取消，库存将自动恢复。"),
                ("step", (1, "点击左侧菜单「交易记录」")),
                ("step", (2, "点击要取消的出库记录旁的撤销(↩)按钮")),
                ("step", (3, "在确认弹窗中点击「确认取消」，库存立即恢复")),
                ("info", "已取消记录以灰色+删除线+已取消徽标显示。无法再次取消或编辑。仅处理出库的网点或管理员可操作。"),
             ]},
            {"title": "7. 批量出库",
             "body": [
                ("step", (1, "点击左侧菜单「批量出库」")),
                ("step", (2, "选择网点后显示表单列表")),
                ("step", (3, "输入各表单出库数量（0则跳过）")),
                ("step", (4, "输入日期和备注，点击「批量出库」")),
                ("tip", "一次操作即可同时处理多个表单的出库，节省时间。"),
             ]},
            {"title": "8. 网点间调拨",
             "body": [
                ("p", "当本网点库存不足时，向有余量的网点发起调拨请求时使用。"),
                ("step", (1, "点击左侧菜单「网点间调拨」")),
                ("step", (2, "选择目标网点、表单、数量、日期")),
                ("step", (3, "点击「申请调拨」→ 状态变为【待审批】")),
                ("info", "接收调拨申请的网点将收到邮件通知。请登录ZEGO网站进行处理。处理完成后，发起申请的网点将收到处理结果的邮件通知。"),
                ("h3", "调拨状态流程"),
                ("table", (["状态","含义"],[
                    ["待审批（Pending）","等待管理员审批"],
                    ["已审批（Approved）","管理员已审批，等待接收网点确认"],
                    ["已确认（Confirmed）","调拨完成，库存已更新"],
                    ["已拒绝（Rejected）","管理员拒绝"],
                ])),
             ]},
            {"title": "9. 交易记录与报表",
             "body": [
                ("p", "查询所有入库、出库、调拨记录。可按网点、表单、交易类型、日期范围筛选。"),
                ("p", "报表：各网点库存状况、月度比较、表单别出库排名、需求预测。"),
             ]},
            {"title": "10. 目录（运输物品）申请",
             "body": [
                ("step", (1, "点击左侧菜单「目录」")),
                ("step", (2, "输入数量后点击「加入购物车」")),
                ("step", (3, "在「申请购物车」中确认后点击「完成申请」")),
                ("step", (4, "在「申请记录」中查看进度")),
                ("info", "仅在管理员设置的申请期间内可进行申请。"),
             ]},
            {"title": "11. 表单耗材申请",
             "body": [
                ("step", (1, "点击左侧菜单「表单申请」")),
                ("step", (2, "选择表单种类，输入申请数量和备注")),
                ("step", (3, "点击「申请」按钮")),
                ("step", (4, "在「我的申请记录」中查看处理状况")),
                ("info", "仅在管理员设置的申请期间内可进行申请。"),
             ]},
            {"title": "12. 管理员功能",
             "body": [
                ("h3", "用户管理"),
                ("p", "创建、编辑、删除用户账户；设置角色；分配网点；重置密码。"),
                ("h3", "申请期间设置"),
                ("p", "运输表单标签：设置表单耗材申请期间。运输物品标签：设置目录申请期间。"),
             ]},
        ],
    },

    "twn": {
        "font": "JhengHei", "fontb": "JhengHei",
        "cover_color": colors.HexColor("#059669"),
        "title": "用戶手冊",
        "sub": "易斯達航空 物流·備品管理系統",
        "meta": "適用對象：全部據點工作人員·管理員  |  語言：中文（繁體）|  2026年6月版",
        "footer": "如有疑問，請聯繫機場服務團隊。  |  hdqapo@eastarjet.com",
        "toc_title": "目錄",
        "sections": [
            {"title": "1. 系統概述",
             "body": [
                ("p", "ZEGO是一個基於Web的系統，用於統一管理易斯達航空各據點的運輸表單及備品庫存。"),
                ("h3", "管理表單種類"),
                ("table", (["分類","表單名稱"],[
                    ["旅客文件","樂器免責聲明書（DECLARATION OF INDEMNITY, Musical Instrument）"],
                    ["旅客文件","監護人聲明書（DECLARATION OF PARENT GUARDIAN）"],
                    ["貨物文件","槍枝交接書（Firearm Handover Form）"],
                    ["特殊貨物","NOTOC（危險品通知書）"],
                    ["一般備品","X形展架、隔離欄、告示文、登機證等"],
                ])),
             ]},
            {"title": "2. 登入與密碼",
             "body": [
                ("step", (1, "在瀏覽器中開啟ZEGO網站")),
                ("step", (2, "輸入帳號和密碼，點擊登入按鈕")),
                ("warn", "初始密碼：$eastar12 — 首次登入時必須變更密碼。"),
                ("info", "密碼到期前7天上方將顯示警告提示，請及時變更。"),
             ]},
            {"title": "3. 儀表板",
             "body": [
                ("p", "登入後的第一個頁面，顯示統計卡片、庫存預警表格、近期交易紀錄及6個月出庫趨勢圖。"),
             ]},
            {"title": "4. 庫存狀況",
             "body": [
                ("p", "點擊左側選單的「庫存狀況」，共有4個標籤頁。"),
                ("h3", "標籤1 — 據點詳細"),
                ("p", "顯示所有據點·表單的當前庫存。在輸入框輸入數字後按Enter，僅儲存該據點的最低標準。"),
                ("h3", "標籤2 — 按據點查看（管理員）"),
                ("p", "以折疊面板展開各據點。最低標準可按據點分別獨立設定。"),
                ("h3", "標籤3 — 按表單彙總"),
                ("p", "按表單種類查看全部據點合計庫存。"),
                ("h3", "標籤4 — 月度消耗狀況"),
                ("p", "查看近6個月各表單出庫數量變化趨勢。"),
                ("tip", "最低標準按據點獨立管理。修改某一據點的最低標準，不會影響其他據點。"),
             ]},
            {"title": "5. 入庫處理",
             "body": [
                ("step", (1, "點擊左側選單「入庫處理」")),
                ("step", (2, "選擇據點，輸入表單種類、數量、日期、備註")),
                ("step", (3, "點擊「入庫處理」按鈕")),
                ("info", "入庫後若仍低於最低標準，系統將自動產生警報通知。"),
             ]},
            {"title": "6. 出庫處理",
             "body": [
                ("step", (1, "點擊左側選單「出庫處理」")),
                ("step", (2, "輸入據點、表單種類、數量、日期、備註")),
                ("step", (3, "點擊「出庫處理」按鈕")),
                ("warn", "出庫數量不能超過當前庫存。出庫後若庫存降至最低標準以下，將自動發送通知。"),
                ("h3", "出庫取消"),
                ("p", "如果出庫處理有誤，可以取消，庫存將自動恢復。"),
                ("step", (1, "點擊左側選單「交易紀錄」")),
                ("step", (2, "點擊要取消的出庫紀錄旁的撤銷(↩)按鈕")),
                ("step", (3, "在確認彈窗中點擊「確認取消」，庫存立即恢復")),
                ("info", "已取消紀錄以灰色+刪除線+已取消徽標顯示。無法再次取消或編輯。僅處理出庫的據點或管理員可操作。"),
             ]},
            {"title": "7. 批量出庫",
             "body": [
                ("step", (1, "點擊左側選單「批量出庫」")),
                ("step", (2, "選擇據點後顯示表單列表")),
                ("step", (3, "輸入各表單出庫數量（0則略過）")),
                ("step", (4, "輸入日期和備註，點擊「批量出庫」")),
                ("tip", "一次操作即可同時處理多個表單的出庫，節省時間。"),
             ]},
            {"title": "8. 據點間調撥",
             "body": [
                ("p", "當本據點庫存不足時，向有餘量的據點發起調撥請求時使用。"),
                ("step", (1, "點擊左側選單「據點間調撥」")),
                ("step", (2, "選擇目標據點、表單、數量、日期")),
                ("step", (3, "點擊「申請調撥」→ 狀態變為「待審核」")),
                ("info", "接收調撥申請的據點將收到郵件通知。請登入ZEGO網站進行處理。處理完成後，發起申請的據點將收到處理結果的郵件通知。"),
                ("h3", "調撥狀態流程"),
                ("table", (["狀態","含義"],[
                    ["待審核（Pending）","等待管理員審核"],
                    ["已審核（Approved）","管理員已審核，等待接收據點確認"],
                    ["已確認（Confirmed）","調撥完成，庫存已更新"],
                    ["已拒絕（Rejected）","管理員拒絕"],
                ])),
             ]},
            {"title": "9. 交易紀錄與報表",
             "body": [
                ("p", "查詢所有入庫、出庫、調撥紀錄。可按據點、表單、交易類型、日期範圍篩選。"),
                ("p", "報表：各據點庫存狀況、月度比較、表單別出庫排名、需求預測。"),
             ]},
            {"title": "10. 目錄（運輸物品）申請",
             "body": [
                ("step", (1, "點擊左側選單「目錄」")),
                ("step", (2, "輸入數量後點擊「加入購物車」")),
                ("step", (3, "在「申請購物車」中確認後點擊「完成申請」")),
                ("step", (4, "在「申請紀錄」中查看進度")),
                ("info", "僅在管理員設定的申請期間內可進行申請。"),
             ]},
            {"title": "11. 表單耗材申請",
             "body": [
                ("step", (1, "點擊左側選單「表單申請」")),
                ("step", (2, "選擇表單種類，輸入申請數量和備註")),
                ("step", (3, "點擊「申請」按鈕")),
                ("step", (4, "在「我的申請紀錄」中查看處理狀況")),
                ("info", "僅在管理員設定的申請期間內可進行申請。"),
             ]},
            {"title": "12. 管理員功能",
             "body": [
                ("h3", "用戶管理"),
                ("p", "建立、編輯、刪除用戶帳戶；設定角色；分配據點；重設密碼。"),
                ("h3", "申請期間設定"),
                ("p", "運輸表單標籤：設定表單耗材申請期間。運輸物品標籤：設定目錄申請期間。"),
             ]},
        ],
    },
}

# ─── BUILD PDF ────────────────────────────────────────────────────────────────

OUT_DIR = "static/manuals"

def build_pdf(lang_key, data):
    font  = data["font"]
    fontb = data["fontb"]
    s = make_styles(font, fontb)

    out_path = f"{OUT_DIR}/{lang_key}.pdf"
    doc = SimpleDocTemplate(
        out_path, pagesize=A4,
        leftMargin=25*mm, rightMargin=25*mm,
        topMargin=18*mm, bottomMargin=18*mm,
    )

    story = []

    # Cover
    story.append(cover_table(data["title"], data["sub"], data["meta"], data["cover_color"]))
    story.append(Spacer(1, 8*mm))

    # Sections
    for sec in data["sections"]:
        story.append(section_title(sec["title"], s))
        story.append(Spacer(1, 2*mm))
        for item in sec["body"]:
            kind = item[0]
            val  = item[1] if len(item) > 1 else None
            if kind == "p":
                story.append(Paragraph(val, s["body"]))
                story.append(Spacer(1, 1.5*mm))
            elif kind == "h3":
                story.append(Paragraph(val, s["h3"]))
            elif kind == "info":
                story.append(info_box(val, s, "info"))
                story.append(Spacer(1, 2*mm))
            elif kind == "warn":
                story.append(info_box(val, s, "warn"))
                story.append(Spacer(1, 2*mm))
            elif kind == "tip":
                story.append(info_box(val, s, "tip"))
                story.append(Spacer(1, 2*mm))
            elif kind == "step":
                story.append(step(val[0], val[1], s))
                story.append(Spacer(1, 1*mm))
            elif kind == "table":
                headers, rows = val
                story.append(data_table(headers, rows, s, font, fontb))
                story.append(Spacer(1, 2*mm))
        story.append(Spacer(1, 4*mm))
        story.append(HRFlowable(width="100%", thickness=0.5, color=BORDER))
        story.append(Spacer(1, 4*mm))

    # Footer
    story.append(Spacer(1, 4*mm))
    story.append(footer_box(data["footer"], s))

    doc.build(story)
    size = os.path.getsize(out_path)
    print(f"  {lang_key}.pdf  ({size/1024:.0f} KB)")

print("Generating PDFs...")
for lang_key, data in LANGS.items():
    print(f"  Building {lang_key}...")
    try:
        build_pdf(lang_key, data)
    except Exception as e:
        print(f"  ERROR: {e}")
        import traceback; traceback.print_exc()

print("Done!")

