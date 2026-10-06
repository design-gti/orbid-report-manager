#!/usr/bin/env python3
"""Konverter: Rekap GTI ORBID (raw score 360) -> Excel format template generator.

Pemakaian:
  python3 convert_gti.py "Rekap GTI ORBID Q2.xlsx" <leader-handle> [output.xlsx]

Pemetaan 3C:
  Commitment   <- raw_score_commitment (engagement self / Q12), 5 kelompok item
  Competency   <- raw_score_culture (360, non-self avg), 4 aspek 4P
  Contribution <- raw_score_ca (nilai leader), 5 aspek
  STAR         <- feedback 'continue' dari leader (raw_score_feedback_trisa)
"""
import sys
from collections import defaultdict
import openpyxl
from openpyxl import Workbook

RAW = sys.argv[1] if len(sys.argv) > 1 else "Rekap GTI ORBID Q2.xlsx"
LEADER = sys.argv[2] if len(sys.argv) > 2 else sys.exit("Sebutkan handle leader, mis.: python3 convert_gti.py rekap.xlsx nama.leader")
OUT = sys.argv[3] if len(sys.argv) > 3 else None
# file hasil olahan batch sebelumnya (opsional): Engagement total -> Commitment,
# Performance manager_score -> Contribution
TREN_FILE = "Nilai ORBID Q1 2026.xlsx"
# hasil Trisa batch sebelumnya (opsional): avg_raw non-self aspek 4P -> Competency
TRISA_FILE = "Hasil Trisa Q1 2026.xlsx"
TRISA_4P = ["Proaktif", "Progressive", "Prototype", "People"]
TREN_LABEL = "ORBID Q1"

# konfigurasi per tim (nama, ringkasan bukti, insight) ada di team_config.py
# (berisi data personal, di-gitignore). Tanpa file itu, konverter tetap jalan
# dengan nama dari handle email, ringkasan otomatis, dan insight kosong.
try:
    from team_config import TEAM_INFO, STAR_SUMMARY, INSIGHTS
except ImportError:
    TEAM_INFO, STAR_SUMMARY, INSIGHTS = {}, {}, {}

# 12 item Gallup Q12 (Q04 memakai recognition_1wk sesuai definisi "7 hari terakhir")
COMMIT_GROUPS = [
    ("Q01. Saya tahu apa yang diharapkan dari saya di tempat kerja", ["basic_needs_expectation"]),
    ("Q02. Saya memiliki perlengkapan untuk bekerja dengan baik", ["basic_needs_tools"]),
    ("Q03. Saya berkesempatan melakukan yang terbaik setiap hari", ["individual_support_talent"]),
    ("Q04. Dalam seminggu terakhir saya menerima pengakuan/pujian", ["recognition_1wk"]),
    ("Q05. Ada yang peduli pada saya sebagai pribadi", ["individual_support_cared"]),
    ("Q06. Ada yang mendorong perkembangan saya", ["individual_support_dev"]),
    ("Q07. Pendapat saya diperhitungkan", ["teamwork_opinion"]),
    ("Q08. Misi perusahaan membuat pekerjaan saya terasa penting", ["teamwork_mission"]),
    ("Q09. Rekan kerja saya berkomitmen pada kualitas", ["teamwork_commit"]),
    ("Q10. Saya punya sahabat baik di tempat kerja", ["teamwork_bestfriend"]),
    ("Q11. Dalam 6 bulan terakhir ada yang membicarakan kemajuan saya", ["growth_discussion"]),
    ("Q12. Dalam setahun terakhir saya berkesempatan belajar dan bertumbuh", ["growth_learning"]),
]
CULTURE_ASPECTS = [
    ("Proactive", "proactive_"), ("Progressive", "progressive_"),
    ("Prototype", "prototype_"), ("People", "people_"),
]
CA_ASPECTS = [
    ("Ketepatan Waktu", "timeliness_deadline"), ("Kualitas Output", "quality_akurat"),
    ("Inisiatif", "initiative_proaktif"), ("Komunikasi & Kolab.", "communication_kolaborasi"),
    ("Pemanfaatan AI", "ai_leverage"),
]
RATER_MAP = {"self": "Self", "leader": "Atasan", "peer": "Rekan", "subordinate": "Bawahan"}
UPWARD_OPEN = [
    ("open_feedback_keep", "APA YANG SEBAIKNYA DILANJUTKAN OLEH MANAGER KAMU?"),
    ("open_feedback_change", "APA YANG SEBAIKNYA DIUBAH OLEH MANAGER KAMU?"),
    ("open_feedback_team_cond", "HAL LAIN YANG PERLU DIKETAHUI MANAGER TENTANG KONDISI KAMU DAN TIM."),
    ("open_feedback_support", "DUKUNGAN YANG KAMU HARAPKAN DARI MANAGER."),
]
UPWARD_ITEMS = [
    ("career_dev_stretch", "My manager assigns stretch opportunities to help me develop new skills."),
    ("goal_clarity_team", "My manager communicates clear goals for our team."),
    ("feedback_actionable", "My manager gives me actionable feedback on a regular basis."),
    ("empathy_consideration", "My manager consistently shows consideration for me as a person."),
    ("prioritization_focus", "My manager keeps the team focused on our priority results."),
    ("transparency_info", "My manager regularly shares relevant information from their manager."),
    ("career_support_discussion", "My manager has had a meaningful discussion with me about career development."),
    ("technical_credibility_mgr", "My manager has the technical expertise required to effectively manage me."),
    ("inclusion_perspective", "The actions of my manager show they value the perspective I bring to the team."),
    ("decision_making_tough", "My manager makes tough decisions effectively."),
    ("cross_func_collab_mgr", "My manager effectively collaborates across boundaries."),
    ("manager_recommend", "I would recommend my manager to others."),
]
STD = 3.5


def handle(email):
    return str(email).split("@")[0]


def num(v):
    try:
        f = float(v)
        return f
    except (TypeError, ValueError):
        return None


def avg(xs):
    xs = [x for x in xs if x is not None]
    return sum(xs) / len(xs) if xs else None


def sheet_rows(wb, name):
    it = wb[name].iter_rows(values_only=True)
    hdr = list(next(it))
    out = []
    for r in it:
        r = list(r) if r else []
        if r and r[0]:
            r += [None] * (len(hdr) - len(r))
            out.append(r)
    return hdr, out


def main():
    wb = openpyxl.load_workbook(RAW, read_only=True, data_only=True)

    # anggota tim = ratee yang leader-nya LEADER (dari relasi rater_as='leader')
    h, d = sheet_rows(wb, "raw_score_culture")
    i_as = h.index("rater_as")
    members = sorted({handle(r[0]) for r in d if r[i_as] == "leader" and handle(r[1]) == LEADER})
    if not members:
        sys.exit(f"Tidak ada anggota untuk leader '{LEADER}'")
    cfg = TEAM_INFO.get(LEADER, {})
    nama = lambda m: cfg.get("nama", {}).get(m, m.replace(".", " ").title())
    print(f"Tim {cfg.get('tim', LEADER)} — {len(members)} anggota: {members}")

    # ── Commitment (Q12 self) ──
    h, d = sheet_rows(wb, "raw_score_commitment")
    commit_items = [g[0] for g in COMMIT_GROUPS]
    commitment, commit_agg = {}, {}
    for r in d:
        m = handle(r[0])
        if m not in members or r[h.index("rater_as")] != "self":
            continue
        vals = [avg([num(r[h.index(c)]) for c in cols]) for _, cols in COMMIT_GROUPS]
        commitment[m] = vals
        commit_agg[m] = avg(vals)

    # ── Competency (culture 360) ──
    h, d = sheet_rows(wb, "raw_score_culture")
    i_as = h.index("rater_as")
    asp_cols = {a: [i for i, c in enumerate(h) if str(c).startswith(p)] for a, p in CULTURE_ASPECTS}
    by_member = defaultdict(lambda: defaultdict(list))  # m -> rater_as -> [row]
    for r in d:
        m = handle(r[0])
        if m in members:
            by_member[m][r[i_as]].append(r)
    competency, comp_agg, blindspot, raters = {}, {}, {}, {}
    for m in members:
        rowsets = by_member[m]
        nonself = [r for k, rs in rowsets.items() if k != "self" for r in rs]
        selfr = rowsets.get("self", [])
        asp_ns, asp_self = [], []
        for a, _ in CULTURE_ASPECTS:
            cols = asp_cols[a]
            asp_ns.append(avg([num(r[i]) for r in nonself for i in cols]))
            asp_self.append(avg([num(r[i]) for r in selfr for i in cols]))
        competency[m] = asp_ns
        comp_agg[m] = avg(asp_ns)
        blindspot[m] = sum(1 for s, ns in zip(asp_self, asp_ns)
                           if s is not None and ns is not None and s - ns >= 0.8)
        all_cols = [i for cols in asp_cols.values() for i in cols]
        raters[m] = {RATER_MAP[k]: avg([num(r[i]) for r in rowsets.get(k, []) for i in all_cols])
                     for k in RATER_MAP}

    # ── Contribution (CA leader) ──
    h, d = sheet_rows(wb, "raw_score_ca")
    i_as = h.index("rater_as")
    contribution, contrib_agg = {}, {}
    for r in d:
        m = handle(r[0])
        if m in members and r[i_as] == "leader":
            vals = [num(r[h.index(c)]) for _, c in CA_ASPECTS]
            contribution[m] = vals
            contrib_agg[m] = avg(vals)

    # ── STAR: feedback 'continue' dari leader ──
    h, d = sheet_rows(wb, "raw_score_feedback_trisa")
    i_as = h.index("rater_as")
    i_cont = h.index("feedback_continue")
    star = {}
    overrides = STAR_SUMMARY.get(LEADER, {})
    for r in d:
        m = handle(r[0])
        if m in members and handle(r[1]) == LEADER and r[i_as] == "leader":
            pts = [p.strip().lstrip("-").strip() for p in str(r[i_cont] or "").split("\n")]
            pts = [p for p in pts if p and p != "-"]
            star[m] = {
                "jumlah": len(pts),
                "validasi": "Tervalidasi" if m in contribution else "Menunggu",
                "ringkasan": overrides.get(m) or "; ".join(pts)[:220],
            }
    for m in members:
        star.setdefault(m, {"jumlah": 0, "validasi": "Menunggu" if m not in contribution else "Tervalidasi", "ringkasan": ""})

    # ── Upward feedback: subordinates menilai LEADER ──
    h, d = sheet_rows(wb, "raw_score_upward_feedback")
    i_as = h.index("rater_as")
    upward = []
    urows = [r for r in d if handle(r[0]) == LEADER and r[i_as] == "subordinate"]
    if urows:
        for col, text in UPWARD_ITEMS:
            ci = h.index(col)
            vals = [num(r[ci]) for r in urows]
            vals = [v for v in vals if v is not None]
            dist = [sum(1 for v in vals if round(v) == s) for s in (1, 2, 3, 4, 5)]
            upward.append({"item": text, "avg": avg(vals), "dist": dist})
        print(f"Upward: {len(urows)} penilai subordinate untuk {cfg.get('manager', LEADER)}")
    else:
        print(f"(tidak ada upward feedback untuk {LEADER} — halaman 6 dilewati)")

    # uraian open feedback (anonim): pecah jadi poin per baris / leading dash
    upward_notes = []
    if urows:
        for col, title in UPWARD_OPEN:
            ci = h.index(col)
            points = []
            for r in urows:
                txt = str(r[ci] or "").strip()
                if txt in ("", "-", "--"):
                    continue
                for p in txt.split("\n"):
                    p = p.strip().lstrip("-").strip()
                    if p and p not in ("-", "--"):
                        points.append(p)
            upward_notes.append({"title": title, "points": points})

    # ── Tren: nilai batch sebelumnya dari file olahan Q1 ──
    tren = {}
    if TREN_FILE:
        try:
            twb = openpyxl.load_workbook(TREN_FILE, read_only=True, data_only=True)
            th, td = sheet_rows(twb, "users")
            email2h = {handle(r[2]): handle(r[2]) for r in td if r[2]}
            name2h = {str(r[1]).strip(): handle(r[2]) for r in td if r[1] and r[2]}
            th, td = sheet_rows(twb, "Engagement")
            for r in td:
                m = name2h.get(str(r[1]).strip())
                if m in members:
                    tren.setdefault(m, {})["commitment"] = num(r[th.index("total_score")])
            th, td = sheet_rows(twb, "Performance")
            for r in td:
                m = name2h.get(str(r[1]).strip())
                if m in members:
                    tren.setdefault(m, {})["contribution"] = num(r[th.index("manager_score")])
        except FileNotFoundError:
            print(f"\n(file tren '{TREN_FILE}' tidak ada — commitment/contribution Q1 dilewati)")
    if TRISA_FILE:
        try:
            swb = openpyxl.load_workbook(TRISA_FILE, read_only=True, data_only=True)
            sh, sd = sheet_rows(swb, "list_users")
            tid2h = {str(int(float(r[0]))): handle(r[2]) for r in sd if r[0] is not None and r[2]}
            sh, sd = sheet_rows(swb, "aspect_others")
            acc = defaultdict(list)
            for r in sd:
                m = tid2h.get(str(int(float(r[0]))))
                if m in members and r[1] in TRISA_4P:
                    acc[m].append(num(r[sh.index("avg_raw")]))
            for m, vals in acc.items():
                tren.setdefault(m, {})["competency"] = avg(vals)
        except FileNotFoundError:
            print(f"\n(file Trisa '{TRISA_FILE}' tidak ada — competency Q1 dilewati)")
    if tren:
        print(f"\nTren {TREN_LABEL}:", {m: {k: round(v, 2) for k, v in tren[m].items()} for m in members if m in tren})

    # ── ringkasan angka (bahan insight) ──
    print("\n%-16s %6s %6s %6s %3s" % ("anggota", "commit", "compet", "contri", "bs"))
    for m in members:
        print("%-16s %6.2f %6.2f %6.2f %3d" % (nama(m), commit_agg[m], comp_agg[m], contrib_agg[m], blindspot[m]))
    print("%-16s %6.2f %6.2f %6.2f" % ("TIM", avg(list(commit_agg.values())),
          avg(list(comp_agg.values())), avg(list(contrib_agg.values()))))
    print("\nCommitment per item (tim):", [f"{g[0]}={avg([commitment[m][i] for m in members]):.2f}" for i, g in enumerate(COMMIT_GROUPS)])
    print("Competency per aspek (tim):", [f"{a}={avg([competency[m][i] for m in members]):.2f}" for i, (a, _) in enumerate(CULTURE_ASPECTS)])
    print("Contribution per aspek (tim):", [f"{a}={avg([contribution[m][i] for m in members]):.2f}" for i, (a, _) in enumerate(CA_ASPECTS)])
    print("Penilai:", {nama(m): {k: (f"{v:.2f}" if v else "-") for k, v in raters[m].items()} for m in members})

    # ── tulis workbook format template ──
    out = Workbook()
    out.remove(out.active)
    def aoa(name, rows):
        ws = out.create_sheet(name)
        for row in rows:
            ws.append(row)
    r1 = lambda v: round(v, 2) if v is not None else ""

    aoa("Info", [
        ["Tim", cfg.get("tim", LEADER)],
        ["Manager", cfg.get("manager", nama(LEADER))],
        ["Batch", "ORBID Q2"],
        ["Penerima", cfg.get("penerima", nama(LEADER))],
        ["Label", ""],
        ["Standar Commitment", STD],
        ["Standar Competency", STD],
        ["Standar Contribution", STD],
    ])
    aoa("Anggota", [["Nama", "Commitment", "Competency", "Contribution", "Blindspot"]] + [
        [nama(m), r1(commit_agg[m]), r1(comp_agg[m]), r1(contrib_agg[m]), blindspot[m]] for m in members])
    aoa("Komitmen", [["Nama"] + commit_items] + [
        [nama(m)] + [r1(v) for v in commitment[m]] for m in members])
    aoa("Kompetensi", [["Nama"] + [a for a, _ in CULTURE_ASPECTS], ["STANDAR"] + [STD] * len(CULTURE_ASPECTS)] + [
        [nama(m)] + [r1(v) for v in competency[m]] for m in members])
    aoa("Penilai", [["Nama", "Self", "Atasan", "Rekan", "Bawahan"]] + [
        [nama(m)] + [r1(raters[m][k]) for k in ("Self", "Atasan", "Rekan", "Bawahan")] for m in members])
    aoa("Kontribusi", [["Nama"] + [a for a, _ in CA_ASPECTS]] + [
        [nama(m)] + [r1(v) for v in contribution[m]] for m in members])
    aoa("STAR", [["Nama", "Jumlah Bukti", "Validasi", "Ringkasan"]] + [
        [nama(m), star[m]["jumlah"], star[m]["validasi"], star[m]["ringkasan"]] for m in members])
    if tren:
        aoa("Tren", [["Batch", "Nama", "Commitment", "Competency", "Contribution"]] + [
            [TREN_LABEL, nama(m), r1(tren.get(m, {}).get("commitment")),
             r1(tren.get(m, {}).get("competency")), r1(tren.get(m, {}).get("contribution"))]
            for m in members])
    if upward:
        aoa("Upward", [["Item", "Average", "N1", "N2", "N3", "N4", "N5"]] + [
            [u["item"], r1(u["avg"])] + u["dist"] for u in upward])
    if upward_notes:
        aoa("Upward Uraian", [["Bagian", "Poin"]] + [
            row for sec in upward_notes
            for row in ([[sec["title"], p] for p in sec["points"]] or [[sec["title"], ""]])])
    ins = INSIGHTS.get(LEADER, {})
    aoa("Insight", [["Halaman", "Teks"]] + [[p, t] for p in sorted(ins) for t in ins[p]])

    fname = OUT or f"ORBID {cfg.get('tim', LEADER)} Q2.xlsx"
    out.save(fname)
    print(f"\nTersimpan: {fname}")


if __name__ == "__main__":
    main()
