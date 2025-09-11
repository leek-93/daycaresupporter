# notice_app.py
import os, json, datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
try:
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except Exception:
    DOCX_AVAILABLE = False

APP_NAME = "Childcare Notice Maker"
OUT_DIR = "output"
STATE_FILE = "rotation_state.json"

# ---------- 시즌/로테이션 문구 ----------
SEASON_LINES = {
    "spring": [
        "포근한 봄바람 속에서 아이들의 마음도 살포시 피어나는 계절입니다.",
        "꽃내음 가득한 길을 걸으며 작은 발견 하나하나가 배움이 되길 바랍니다.",
        "새로운 시작을 응원하며, 하루하루의 설렘을 조심스레 품어 보겠습니다."
    ],
    "summer": [
        "반짝이는 햇살처럼 아이들의 웃음도 환한 계절입니다. 건강과 안전을 먼저 챙기겠습니다.",
        "시원한 쉼과 알찬 놀이가 균형을 이루도록 일정과 환경을 세심히 준비했습니다.",
        "더운 날씨에도 마음은 가볍게, 물 한 모금과 그늘 한 자리까지 꼼꼼히 살피겠습니다."
    ],
    "autumn": [
        "높고 맑은 하늘 아래 아이들의 하루가 고운 빛으로 물들어 갑니다.",
        "바스락거리는 낙엽처럼 작은 변화도 소중히 담아 따뜻한 배움으로 이어가겠습니다.",
        "수확의 기쁨처럼 노력의 열매가 맺어지도록 차분히 동행하겠습니다."
    ],
    "winter": [
        "차가운 바람 속에서도 교실은 늘 포근하도록, 안전과 건강을 먼저 살피겠습니다.",
        "따뜻한 마음과 정돈된 일과로 아이들의 하루를 든든하게 채우겠습니다.",
        "서늘한 계절에도 작은 성취를 꼼꼼히 기록하며 잔잔한 기쁨을 나누겠습니다."
    ]
}

INTRO_VARIANTS = {
    "Field Trip": [
        "아이들이 직접 보고 듣고 만지며 배우는 시간이 될 수 있도록 차분히 준비했습니다. 안전을 가장 먼저 생각하며, 아이 한 명 한 명의 속도에 맞춰 살피겠습니다.",
        "자연과 사회를 가까이에서 경험하는 하루입니다. 친구와 배려하고 질서를 지키는 작은 습관까지 함께 익힐 수 있도록 교사들이 곁에서 세심히 도우겠습니다.",
        "설렘 가득한 마음이 배움으로 이어지도록 일정과 동선을 꼼꼼히 구성했습니다. 무사히 다녀올 수 있도록 가정의 격려와 협조를 부탁드립니다."
    ],
    "Class Observation": [
        "평소 교실에서의 배움과 놀이가 어떻게 흐르는지, 아이들이 어떤 표정으로 하루를 보내는지 천천히 살펴보실 수 있도록 준비했습니다.",
        "가정과 원이 같은 마음으로 아이를 바라볼 때 성장의 결이 고와집니다. 편안한 마음으로 오셔서 작은 변화도 함께 기뻐해 주세요.",
        "교사와 친구들 사이에서 나누는 따뜻한 상호작용을 가까이에서 보시며 가정과의 연계에 도움이 되길 바랍니다."
    ],
    "Picnic": [
        "계절의 숨결을 온몸으로 느끼며 자연 속에서 마음껏 뛰놀 수 있도록 소풍을 마련했습니다. 물 한 모금, 모자 하나까지 정성껏 챙기겠습니다.",
        "친구와 함께 걷고 나누는 시간이 예절과 배려를 단단히 세웁니다. 안전한 이동과 충분한 휴식으로 편안한 하루가 되게 하겠습니다.",
        "작은 잎 하나, 바람 한 줄기에도 놀라고 배우는 때입니다. 아이들의 기쁨이 오래 기억으로 남도록 세심히 동행하겠습니다."
    ],
    "Health/Immunization": [
        "건강은 하루의 배움과 생활을 지탱하는 첫걸음입니다. 필요한 확인을 차분히 진행하고자 하오니 몇 가지 안내에 협조 부탁드립니다.",
        "알레르기·복용 약 등 중요한 정보는 교실에서 곧바로 돌봄에 반영됩니다. 아이에게 가장 안전한 환경이 될 수 있도록 함께 살펴 주세요.",
        "작은 증상도 소중히 듣고 살피겠습니다. 궁금하신 점은 언제든 편히 말씀해 주시면 가정과 상의하며 좋은 선택을 하겠습니다."
    ],
    "General Notice": [
        "늘 아이들의 하루를 믿고 응원해 주셔서 고맙습니다. 필요한 소식만 차분히 정리해 드리오니 천천히 확인 부탁드립니다.",
        "가정과 원이 한마음으로 아이를 바라볼 때 하루가 더 따뜻해집니다. 안내를 살펴보시고 의견 주시면 더 세심히 반영하겠습니다.",
        "작은 약속을 지키는 힘이 큰 성장을 만듭니다. 투명하게 소통하며 믿음을 이어가겠습니다."
    ]
}

# ---------- 로테이션 상태 관리 ----------
def load_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

def pick_rotating(key, candidates, state):
    if not candidates:
        return ""
    idx = state.get(key, 0)
    text = candidates[idx % len(candidates)]
    state[key] = (idx + 1) % len(candidates)
    return text

def season_of(d):
    m = d.month
    if m in (3,4,5): return "spring"
    if m in (6,7,8): return "summer"
    if m in (9,10,11): return "autumn"
    return "winter"

# ---------- 문안 생성 ----------
def build_notice(params):
    """
    params: dict {
      center, classname, contact_name, contact_phone,
      event_type, event_name, date, start, end, location,
      materials, transport, cost, rsvp_deadline, rain_plan, body_summary
    }
    """
    today = datetime.date.today().strftime("%Y-%m-%d")
    d = datetime.datetime.strptime(params["date"], "%Y-%m-%d").date()
    t_range = ""
    if params.get("start") and params.get("end"):
        t_range = f"{params['date']} {params['start']}–{params['end']}"
    elif params.get("start") or params.get("end"):
        t_range = f"{params['date']} {params.get('start') or params.get('end')}"

    # 시즌/로테이션
    state = load_state()
    season = season_of(d)
    season_line = pick_rotating(f"SEASON_{season}", SEASON_LINES[season], state)
    intro_line = pick_rotating(params["event_type"], INTRO_VARIANTS.get(params["event_type"], INTRO_VARIANTS["General Notice"]), state)
    save_state(state)

    long_intro = (
        f"사랑하는 {params['center']} {params['classname']} 학부모님께,\n\n"
        "안녕하십니까. 늘 저희 교육 활동에 따뜻한 관심을 보내주시는 모든 가정께 깊이 감사드립니다.\n"
        f"{season_line}\n{intro_line}\n"
    )

    # 공통 헬퍼
    def opt(label, v):
        return f"{label}{v}\n" if v and str(v).strip() else ""
    def fee(v):
        if not v or str(v).strip() in ("0 KRW","0원","0"): return ""
        return f"• 참가비: {v}\n"

    # 본문
    et = params["event_type"]
    head = f"[가정통신문] {params['event_name'] or et}"
    foot = (
        "\n※ 위 일정 및 내용은 원 사정과 안전 상황에 따라 변경될 수 있습니다.\n"
        f"{today}\n{params['center']} 원장 {params['contact_name']} 드림"
    )

    if et == "Class Observation":
        body = (
            f"{head}\n\n{long_intro}\n"
            "■ 참관 수업 개요\n"
            f"• 일시: {t_range}\n"
            f"• 장소: {params['location']}\n"
            f"{opt('• 준비물: ', params['materials'])}"
            "\n■ 안내 및 협조 요청\n"
            "• 원내 안전을 위해 가정당 보호자 1인 참석을 권장드립니다.\n"
            "• 편안한 복장을 권하며, 외부 음식 반입은 자제 부탁드립니다.\n"
            "• 수업 중 임의 촬영은 제한되며, 별도 촬영본은 추후 공유드리겠습니다.\n"
            "\n■ 참석 회신\n"
            f"• 마감: {params['rsvp_deadline']}\n"
            f"• 링크/회신 방법: (원에서 안내한 방식으로 회신)\n"
            f"{foot}"
        )
    elif et == "Field Trip":
        body = (
            f"{head}\n\n{long_intro}\n"
            "■ 체험학습 개요\n"
            f"• 일시: {t_range}\n"
            f"• 장소/프로그램: {params['location']}{opt(' / ', params['body_summary']).strip()}\n"
            f"{opt('• 이동수단: ', params['transport'])}{fee(params['cost'])}{opt('• 준비물/복장: ', params['materials'])}"
            "\n■ 안전 및 동의\n"
            "• 사진·영상 촬영 동의 및 안전수칙 확인이 필요합니다.\n"
            "• 알레르기/복용 약 등 특이사항은 반드시 기재해 주세요.\n"
            "\n■ 전자 동의서\n"
            f"• 마감: {params['rsvp_deadline']}\n"
            f"• 링크/제출 방법: (원에서 안내한 방식으로 제출)\n"
            f"{foot}"
        )
    elif et == "Picnic":
        body = (
            f"{head}\n\n{long_intro}\n"
            "■ 소풍 개요\n"
            f"• 일시: {t_range}\n"
            f"• 장소: {params['location']}\n"
            f"{opt('• 준비물/복장: ', params['materials'])}{opt('• 우천 시 대체: ', params['rain_plan'])}{opt('• 이동수단: ', params['transport'])}"
            "\n■ 안내 및 협조 요청\n"
            "• 사진·영상 촬영 동의 및 안전수칙 확인을 부탁드립니다.\n"
            "• 알레르기/식단 제한이 있는 경우 간식 제공에 반영할 수 있도록 회신에 남겨 주세요.\n"
            "\n■ 전자 동의서\n"
            f"• 마감: {params['rsvp_deadline']}\n"
            f"• 링크/제출 방법: (원에서 안내한 방식으로 제출)\n"
            f"{foot}"
        )
    elif et == "Health/Immunization":
        body = (
            f"{head}\n\n{long_intro}\n"
            "■ 진행 안내\n"
            f"• 일시/장소: {t_range} / {params['location']}\n"
            f"{opt('• 준비물: ', params['materials'])}"
            "\n■ 확인·동의 사항\n"
            "• 아이의 건강 상태, 알레르기, 복용 약 등을 정확히 기재해 주세요.\n"
            "• 필요한 경우 개별 상담을 진행하겠습니다.\n"
            "\n■ 확인서 제출\n"
            f"• 마감: {params['rsvp_deadline']}\n"
            f"• 링크/제출 방법: (원에서 안내한 방식으로 제출)\n"
            f"{foot}"
        )
    else:
        body = (
            f"{head}\n\n{long_intro}\n"
            "■ 안내 내용\n"
            f"• {params['body_summary'] or params['materials'] or '상세 내용은 첨부를 참고해 주세요.'}\n"
            "\n■ 확인/회신(해당 시)\n"
            "• 링크/회신 방법: (원에서 안내한 방식으로 회신)\n"
            f"{foot}"
        )

    return body

# ---------- 파일 저장 ----------
def save_txt(content, path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)

def save_docx(content, path):
    if not DOCX_AVAILABLE:
        return False
    doc = Document()
    # 기본 폰트
    style = doc.styles['Normal']
    style.font.name = 'Malgun Gothic'  # 맑은 고딕
    style.font.size = Pt(11)

    lines = content.split("\n")
    # 제목 처리(첫 줄)
    if lines:
        title = lines[0].strip()
        p = doc.add_paragraph(title)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.runs[0]
        run.font.bold = True
        run.font.size = Pt(15)
        doc.add_paragraph().add_run("------------------------------")

    for i, line in enumerate(lines[1:], start=1):
        doc.add_paragraph(line)

    os.makedirs(os.path.dirname(path), exist_ok=True)
    doc.save(path)
    return True

# ---------- UI ----------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("720x520")
        self.resizable(False, False)

        # 필드
        self.center = tk.StringVar(value="아이조아 어린이집")
        self.classname = tk.StringVar(value="해바라기")
        self.contact_name = tk.StringVar(value="원장 김OO")
        self.contact_phone = tk.StringVar(value="010-1234-0000")

        self.event_type = tk.StringVar(value="Field Trip")
        self.event_name = tk.StringVar(value="체험학습 안내")
        self.date = tk.StringVar(value=datetime.date.today().strftime("%Y-%m-%d"))
        self.start = tk.StringVar(value="09:30")
        self.end = tk.StringVar(value="14:00")
        self.location = tk.StringVar(value="시립 자연학습원")
        self.materials = tk.StringVar(value="물통, 모자")
        self.transport = tk.StringVar(value="버스")
        self.cost = tk.StringVar(value="0원")
        self.rsvp_deadline = tk.StringVar(value=(datetime.date.today()+datetime.timedelta(days=5)).strftime("%Y-%m-%d 23:59"))
        self.rain_plan = tk.StringVar(value="우천 시 실내 활동")
        self.body_summary = tk.StringVar(value="세부 일정은 가정통신문 뒷면 참고")

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)

        r=0
        def row(label, var, width=30):
            nonlocal r
            ttk.Label(frm, text=label, width=18, anchor="e").grid(row=r, column=0, padx=6, pady=4, sticky="e")
            e = ttk.Entry(frm, textvariable=var, width=width)
            e.grid(row=r, column=1, columnspan=3, sticky="w")
            r+=1

        ttk.Label(frm, text="행사유형", width=18, anchor="e").grid(row=r, column=0, padx=6, pady=4, sticky="e")
        cb = ttk.Combobox(frm, textvariable=self.event_type, values=list(INTRO_VARIANTS.keys()), width=28, state="readonly")
        cb.grid(row=r, column=1, sticky="w"); r+=1

        row("행사명/제목", self.event_name)
        row("어린이집 이름", self.center)
        row("반 이름", self.classname)
        row("일자 (YYYY-MM-DD)", self.date)
        row("시작/종료 (HH:MM-HH:MM)", self.start)
        row("종료시간", self.end)
        row("장소", self.location)
        row("준비물", self.materials)
        row("이동수단", self.transport)
        row("참가비", self.cost)
        row("회신마감 (YYYY-MM-DD HH:MM)", self.rsvp_deadline)
        row("우천대응", self.rain_plan)
        row("요약(선택)", self.body_summary)
        row("원장 성함", self.contact_name)
        row("연락처", self.contact_phone)

        btns = ttk.Frame(frm); btns.grid(row=r, column=0, columnspan=4, pady=12)
        ttk.Button(btns, text="문안 생성 (TXT)", command=self.make_txt).grid(row=0, column=0, padx=6)
        ttk.Button(btns, text="문안 + DOCX 생성", command=self.make_docx).grid(row=0, column=1, padx=6)

        self.status = tk.StringVar(value="준비됨")
        ttk.Label(self, textvariable=self.status, relief="sunken", anchor="w").pack(fill="x", padx=12, pady=(0,8))

    def collect(self):
        return dict(
            center=self.center.get().strip(),
            classname=self.classname.get().strip(),
            contact_name=self.contact_name.get().strip(),
            contact_phone=self.contact_phone.get().strip(),
            event_type=self.event_type.get().strip(),
            event_name=self.event_name.get().strip(),
            date=self.date.get().strip(),
            start=self.start.get().strip(),
            end=self.end.get().strip(),
            location=self.location.get().strip(),
            materials=self.materials.get().strip(),
            transport=self.transport.get().strip(),
            cost=self.cost.get().strip(),
            rsvp_deadline=self.rsvp_deadline.get().strip(),
            rain_plan=self.rain_plan.get().strip(),
            body_summary=self.body_summary.get().strip()
        )

    def _save_paths(self, title_base):
        os.makedirs(OUT_DIR, exist_ok=True)
        safe = "".join(c for c in title_base if c.isalnum() or c in ("_","-"," ")).rstrip()
        txt_path = os.path.join(OUT_DIR, f"{safe}.txt")
        docx_path = os.path.join(OUT_DIR, f"{safe}.docx")
        return txt_path, docx_path

    def make_txt(self):
        params = self.collect()
        try:
            content = build_notice(params)
            base = f"공지_{params['event_type']}__{params['classname']}_{params['date']}"
            txt_path, _ = self._save_paths(base)
            save_txt(content, txt_path)
            self.status.set(f"TXT 생성 완료 → {txt_path}")
            messagebox.showinfo(APP_NAME, f"문안(TXT) 생성 완료\n\n{txt_path}")
        except Exception as e:
            self.status.set("오류")
            messagebox.showerror(APP_NAME, str(e))

    def make_docx(self):
        if not DOCX_AVAILABLE:
            messagebox.showwarning(APP_NAME, "python-docx가 설치되지 않았습니다.\n먼저 'pip install python-docx'를 실행해 주세요.")
            return
        params = self.collect()
        try:
            content = build_notice(params)
            base = f"공지_{params['event_type']}__{params['classname']}_{params['date']}"
            txt_path, docx_path = self._save_paths(base)
            save_txt(content, txt_path)
            ok = save_docx(content, docx_path)
            self.status.set(f"DOCX 생성 완료 → {docx_path}")
            messagebox.showinfo(APP_NAME, f"문안(TXT/DOCX) 생성 완료\n\n{docx_path}")
        except Exception as e:
            self.status.set("오류")
            messagebox.showerror(APP_NAME, str(e))

if __name__ == "__main__":
    App().mainloop()
