import os
import pandas as pd
from datetime import date, timedelta

# Default input/output filenames (safe for GitHub)
INPUT_XLSX = os.getenv("INPUT_XLSX", "sample_final.xlsx")
OUTPUT_XLSX = os.getenv("OUTPUT_XLSX", "scheduled_exams.xlsx")
OUTPUT_CSV = os.getenv("OUTPUT_CSV", "scheduled_exams.csv")

EXAM_START_DATE = date.today()
TIME_SLOTS = ["09:00", "11:00", "13:00", "15:00", "17:00", "19:00", "21:00"]
MAX_EXAMS_PER_CLASS_PER_DAY = 1

RA_BY_DEPT = {
    "BİLGİSAYAR MÜHENDİSLİĞİ": [
        "Kadriye KARADENİZ", "Kadriye ÖZ", "Melek Santurk", "Öznur DİNCEL",
        "Elif DORUKBAŞI", "Hasan Demir"
    ],
    "MEKATRONİK MÜHENDİSLİĞİ": [
        "Ömer KARTALTEPE", "Semih PAK", "Ali KAFALI", "Mehmet İzzeddin GÜLER",
        "Mustafa Feyzi TEMEL"
    ],
    "ENDÜSTRİ MÜHENDİSLİĞİ": ["Muammer DOLMACI", "Muhammed Zahid KOÇ"],
    "MAKİNE MÜHENDİSLİĞİ": [
        "Mehmet Tayyip ÖZDEMİR", "Abdullah DAĞDEVİREN", "Bahaddin TOPAK",
        "Efe KARAAVCI", "Mesut YILMAZ", "Mücahid CAN", "Süheyl Bilal SUNGUR",
        "Turan DAŞ", "Bilgehan KONDUL UĞUR", "Seyit Ali KARA", "Yakup DAŞDEMİRLİ"
    ],
    "METALURJİ VE MALZEME MÜHENDİSLİĞİ": ["Atakan Oğuz OCAK"],
    "YAZILIM MÜHENDİSLİĞİ": ["Saliha ÖZGÜNGÖR"],
    "BİYOMEDİKAL MÜHENDİSLİĞİ": ["Halil İbrahim ŞAHİN", "Sena AKSOY"],
    "ELEKTRİK - ELEKTRONİK MÜHENDİSLİĞİ": [
        "İbrahim Ethem YILMAZ", "Tahsin HAKTANIR", "Cemil ZEYVELİ",
        "Leyla ZORLU", "Ekrem DEMİR", "Ali ART", "Betül KARAOĞLAN"
    ],
    "İNŞAAT MÜHENDİSLİĞİ": [
        "Selman KAHRAMAN", "Yusuf BAHÇACI", "İsmail TOZLU",
        "Mahfuz PEKGÖZ", "İbrahim TORLAK"
    ]
}


def safe_int(value, default=0):
    try:
        val = str(value).strip()
        return int(val) if val.isdigit() else default
    except Exception:
        return default


def schedule_exams(df):
    exams = []
    for _, r in df.iterrows():
        exam = {
            "Bölüm": r.get("Bölüm"),
            "Ders Kodu": r.get("Ders Kodu"),
            "Ders Adı": r.get("Ders Adı"),
            "Öğretim Görevlisi": r.get("Öğretim Görevlisi"),
            "Sınıf": r.get("Sınıf"),
            "Öğrenci Sayısı": safe_int(r.get("Öğrenci Sayısı"), 0),
            "Derslik": r.get("Derslik"),
            "Derslik Kapasitesi": safe_int(r.get("Derslik Kapasitesi"), 0)
        }

        exam_date = r.get("Sınav Tarihi")
        exam_time = r.get("Sınav Saati")

        if pd.notna(exam_date):
            try:
                exam["Sınav Tarihi"] = pd.to_datetime(exam_date, dayfirst=True).date()
            except Exception:
                exam["Sınav Tarihi"] = EXAM_START_DATE
        else:
            exam["Sınav Tarihi"] = None

        exam["Sınav Saati"] = str(exam_time).strip() if pd.notna(exam_time) else None
        exams.append(exam)

    scheduled = []
    current_date = EXAM_START_DATE
    slot_index = 0
    class_day_count = {}

    for exam in exams:
        if exam["Sınav Tarihi"]:
            if not exam["Sınav Saati"]:
                exam["Sınav Saati"] = TIME_SLOTS[0]
            scheduled.append(exam)
            continue

        key = (exam["Sınıf"], current_date)
        if class_day_count.get(key, 0) >= MAX_EXAMS_PER_CLASS_PER_DAY:
            slot_index += 1
            if slot_index >= len(TIME_SLOTS):
                slot_index = 0
                current_date += timedelta(days=1)

        scheduled.append({
            **exam,
            "Sınav Tarihi": current_date,
            "Sınav Saati": TIME_SLOTS[slot_index]
        })
        class_day_count[key] = class_day_count.get(key, 0) + 1

    return scheduled


def assign_ras(scheduled_exams):
    ra_tasks = {ra: 0 for ras in RA_BY_DEPT.values() for ra in ras}
    busy_slots = {ra: set() for ras in RA_BY_DEPT.values() for ra in ras}
    assignments = []

    for exam in scheduled_exams:
        exam_date = exam["Sınav Tarihi"]
        exam_time = exam["Sınav Saati"]
        dept = exam["Bölüm"]
        candidates = RA_BY_DEPT.get(dept, [])

        free_candidates = [ra for ra in candidates if (exam_date, exam_time) not in busy_slots[ra]]

        chosen = None
        if free_candidates:
            chosen = min(free_candidates, key=lambda ra: ra_tasks[ra])
        else:
            all_ras = [ra for ra in ra_tasks.keys() if (exam_date, exam_time) not in busy_slots[ra]]
            if all_ras:
                chosen = min(all_ras, key=lambda ra: ra_tasks[ra])
            else:
                instr = exam["Öğretim Görevlisi"]
                if (exam_date, exam_time) not in busy_slots.get(instr, set()):
                    chosen = instr
                    busy_slots.setdefault(instr, set())

        if chosen:
            ra_tasks[chosen] = ra_tasks.get(chosen, 0) + 1
            busy_slots.setdefault(chosen, set()).add((exam_date, exam_time))
            assignments.append({**exam, "RA": chosen})
        else:
            assignments.append({**exam, "RA": None})

    ra_stat_list = []
    for ra, total in ra_tasks.items():
        slot_counts = {}
        for exam in assignments:
            if exam["RA"] == ra:
                time_slot = exam["Sınav Saati"]
                slot_counts[time_slot] = slot_counts.get(time_slot, 0) + 1

        dept = next((d for d, ras in RA_BY_DEPT.items() if ra in ras), "")
        ra_stat_list.append({
            "RA": ra,
            "Department": dept,
            "TotalTasks": total,
            **{slot: slot_counts.get(slot, 0) for slot in TIME_SLOTS}
        })

    ra_stats_df = pd.DataFrame(ra_stat_list)
    return assignments, ra_stats_df


def main():
    if not os.path.exists(INPUT_XLSX):
        print(f"⚠️ Input file '{INPUT_XLSX}' not found. Please provide an Excel file.")
        return

    df = pd.read_excel(INPUT_XLSX)
    scheduled = schedule_exams(df)
    assigned, ra_stats_df = assign_ras(scheduled)
    out_df = pd.DataFrame(assigned)

    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name="Exams", index=False)
        ra_stats_df.to_excel(writer, sheet_name="RA Stats", index=False)

    out_df.to_csv(OUTPUT_CSV, index=False)
    print("✅ Schedule and RA assignments saved.")
    print(ra_stats_df)


if __name__ == "__main__":
    main()
