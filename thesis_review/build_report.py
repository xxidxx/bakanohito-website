# -*- coding: utf-8 -*-
"""Build per-chapter thesis review report (Bahasa Indonesia) as .docx."""
import json
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = r"C:/Users/idari/Downloads/Hermes/Laporan_Masukan_Review_Thesis_v2.docx"

C_KRITIS = "C00000"
C_TINGGI = "C55A11"
C_SEDANG = "1F6F8F"
C_RENDAH = "595959"
C_HDR = "1F3864"

doc = Document()
for s in doc.sections:
    s.left_margin = Cm(2.2); s.right_margin = Cm(2.2)
    s.top_margin = Cm(2.0); s.bottom_margin = Cm(2.0)

st = doc.styles["Normal"]
st.font.name = "Times New Roman"; st.font.size = Pt(11)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "SimSun")

def _shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)

def H(text, level):
    h = doc.add_heading("", level=level)
    r = h.add_run(text)
    r.font.name = "Times New Roman"; r.font.size = Pt({1: 15, 2: 13, 3: 11.5}[level])
    r.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
    r.font.bold = True
    return h

def P(text, bold_lead=None, size=11, italic=False):
    p = doc.add_paragraph()
    if bold_lead:
        r = p.add_run(bold_lead); r.bold = True
        r.font.size = Pt(size); r.font.color.rgb = RGBColor.from_string(C_RENDAH)
    r = p.add_run(text); r.font.size = Pt(size); r.italic = italic
    return p

def B(items):
    for it in items:
        p = doc.add_paragraph(style="List Bullet")
        if isinstance(it, tuple) and len(it) == 3:
            lead, rest, sev = it
            r = p.add_run(lead); r.bold = True
            r.font.color.rgb = RGBColor.from_string({"KRITIS": C_KRITIS, "TINGGI": C_TINGGI,
                                                      "SEDANG": C_SEDANG, "RENDAH": C_RENDAH}[sev])
            p.add_run(rest)
        elif isinstance(it, tuple) and len(it) == 2:
            lead, rest = it
            r = p.add_run(lead); r.bold = True
            p.add_run(rest)
        else:
            p.add_run(it)

def T(headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    for j, htxt in enumerate(headers):
        c = t.rows[0].cells[j]
        c.text = ""
        r = c.paragraphs[0].add_run(htxt); r.bold = True; r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        _shade(c, C_HDR)
    for row in rows:
        cells = t.add_row().cells
        for j, val in enumerate(row):
            cells[j].text = ""
            r = cells[j].paragraphs[0].add_run(str(val)); r.font.size = Pt(9.5)
    if widths:
        for j, w in enumerate(widths):
            for row in t.rows:
                row.cells[j].width = Cm(w)
    doc.add_paragraph()

# ---------------- COVER ----------------
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("LAPORAN MASUKAN (REVIEW)\nDraft Akhir Thesis – Revisi v2"); r.bold = True; r.font.size = Pt(17)
r.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
P("Research on the Marketing Strategy of Pertamina's Pertamax Green 95 in the Indonesian Market", size=12, italic=True)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Penulis: Id Arifandi (2428062-李帆)   |   Reviewer: Hermes   |   17 September 2026")
r.font.size = Pt(10); r.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
P("Dokumen diperiksa: 2428062-李帆-Research on the Marketing Strategies of Pertamina_REVISED_v2.docx "
  "(937 paragraf, 61 tabel, ~59 halaman teks). Format laporan: setiap bab mendapat sub-bagian tersendiri, "
  "setiap temuan diberi level keparahan [KRITIS / TINGGI / SEDANG / RENDAH] beserta lokasi dan perbaikan yang disarankan.")

# ---------------- 1. EXEC SUMMARY ----------------
H("1. Ringkasan Eksekutif", 1)
P("Kesimpulan umum: substansi draft ini KUAT dan layak lanjut. Alur diagnosis-masalah-strategi-"
  "implementasi-KPI koheren; bukti ditriangulasi (survey n=124, PESTEL, Porter Five Forces, SWOT-TOWS); "
  "strategi 7P dipetakan eksplisit ke lima masalah di Bab 3; ada cost-benefit analysis dan risk analysis "
  "yang jarang ditemukan di tesis MBA. Temuan inti \u201cmarketing activation failure, bukan product failure\u201d "
  "adalah argumen yang tajam dan didukung data.")
P("Masalah yang harus dibereskan justru sebagian besar bersifat teknis-konsistensi: (1) tracked changes belum "
  "di-accept sehingga banyak kalimat tampak \u201cbolong\u201d dan Abstrak Mandarin tampak kosong; (2) semua caption "
  "tabel/gambar tanpa nomor (\u201cTable .\u201d, \u201cFigure .\u201d); (3) sejumlah angka tidak konsisten antar-bab; "
  "(4) Lampiran kuesioner tidak lengkap; (5) satu referensi masih placeholder. Rincian di bawah.")
B([
 ("[KRITIS] ", "Dokumen masih mengandung ratusan tracked changes (insert/delete) dari \u201cClaude (revision "
   "assistant)\u201d yang BELUM di-accept \u2014 termasuk seluruh isi Abstrak Mandarin. Efeknya: Abstrak Mandarin tampak "
   "kosong, kalimat Abstrak Inggris patah di tengah, dan banyak kalimat di Bab 3\u20135 tampak kehilangan kata. "
   "Langkah pertama: Review \u2192 Accept All Changes, lalu scan ulang.", "KRITIS"),
 ("[TINGGI] ", "Penomoran Tabel/Gambar belum ada sama sekali di caption, padahal teks merujuk \u201cTable 3.6\u2013"
   "3.9, 4.4\u20134.6, 5.7\u201d.", "TINGGI"),
 ("[TINGGI] ", "Konsistensi angka: tanggal luncur (Juli vs 20 Juni vs Juni 2023), jumlah SPBU (160/170/177), "
   "volume per outlet (624 vs 406\u2192292 L/hari), item kuesioner (32 vs 35/38), statistik tema C2-T1 "
   "(39,5% vs 44,35%), mean P5/P6 beda antar-bab.", "TINGGI"),
 ("[TINGGI] ", "Kalimat-kalimat terpotong tersebar (kehilangan kata saat revisi): \u00a77.2 kalimat terakhir "
   "thesis tidak selesai, \u00a77.1 \u201cno competitor can .\u201d, \u00a75.2.1 \u201cshould\u201d\u00a0The current\u2026\u201d, \u00a76.1/\u00a76.2/\u00a76.3, \u00a73.3.2 "
   "(2 tempat), \u00a72.2.1, \u00a74.4.", "TINGGI"),
 ("[TINGGI] ", "Lampiran kuesioner tidak lengkap: item 7Ps (P1.1\u2013P7.4) tidak ditampilkan, bagian D berhenti "
   "di D4.", "TINGGI"),
 ("[TINGGI] ", "Referensi [58] masih placeholder (nama penulis & jurnal \u201c[to be confirmed]\u201d).", "TINGGI"),
 ("[SEDANG] ", "Kontradiksi metodologis yang pasti ditanya penguji: kriteria purposive sampling mensyaratkan "
   "responden \u201caware of Pertamax Green 95\u201d, tetapi hasil survey melaporkan 21,8% \u201cnever heard\u201d; frame "
   "sampling 4 kota vs data \u201c22 cities\u201d.", "SEDANG"),
])

# ---------------- A. TRACKED CHANGES ----------------
H("2. Kondisi Dokumen & Langkah Pertama (Wajib Sebelum Submit)", 1)
P("Dokumen sedang dalam kondisi mid-revision: semua revisi dari asisten revisi masih tercatat sebagai tracked "
  "changes. Ini menjelaskan hampir semua \u201ckata hilang\u201d yang dilaporkan di laporan ini \u2014 beberapa gap bisa "
  "hilang dengan sendirinya setelah revisi di-accept, tetapi beberapa adalah fragmen asli yang harus diperbaiki "
  "manual.")
B([
 ("[KRITIS] ", "Accept All Changes (Home \u2192 Tracking \u2192 Accept All). Jangan kirim dokumen dengan tracked "
   "changes ke pembimbing \u2014 nama \u201cClaude (revision assistant)\u201d akan terlihat dan mengurangi kredibilitas.", "KRITIS"),
 ("[KRITIS] ", "Setelah accept, periksa satu-per-satu lokasi \u201ckalimat terpotong\u201d yang didaftar di laporan ini. "
   "Beberapa akan tersambung oleh teks hasil revisi; sisanya (mis. \u201cshould\u00a0The current\u201d, \u201cno competitor can .\u201d) "
   "adalah fragmen asli yang perlu ditulis ulang.", "KRITIS"),
 ("[TINGGI] ", "Periksa ulang seluruh angka 7Ps setelah accept \u2014 sebagian edit revisi mengubah narasi "
   "(mis. \u201cstrategic paradox\u201d \u2192 \u201ccentral inconsistency\u201d) tanpa mengubah tabel statistik; pastikan narasi dan "
   "tabel tetap sinkron.", "TINGGI"),
])

# ---------------- B. FRONT MATTER ----------------
H("3. Front Matter (Halaman Judul, 独创性声明, 摘 要, Abstract, Daftar Isi)", 1)
P("Kelebihan: struktur sesuai format tesis NEU; pernyataan orisinalitas dan lisensi hak cipta lengkap dengan "
  "tanggal; kata kunci tersedia dalam dua bahasa.")
B([
 ("[KRITIS] ", "Abstrak Mandarin tampak KOSONG di teks (isi masih tersimpan sebagai tracked insertion). Setelah "
   "Accept All, pastikan paragraf isi muncul di antara judul \u201c摘 要\u201d dan baris \u201c关键词\u201d \u2014 jika tidak, tempel "
   "manual dari revisi.", "KRITIS"),
 ("[KRITIS] ", "Abstrak Inggris patah: kalimat pertama \u201c\u2026launched Pertamax Green 95 it remains available\u2026\u201d "
   "kehilangan klausa penghubung (seharusnya \u201c\u2026in June 2023\u2026 Yet after 28 months on the market, it remains\u2026\u201d); "
   "paragraf kelima dimulai langsung \u201cdigital promotion, distribution densification\u2026\u201d tanpa induk kalimat. "
   "Rapikan setelah accept revisi; baca keras seluruh abstrak.", "KRITIS"),
 ("[TINGGI] ", "Abstrak Inggris menyebut \u201c177 of 6,000+ gas stations\u201d sedangkan Bab 1 menyebut 170 \u2014 "
   "samakan satu angka (lihat Tabel Audit Konsistensi).", "TINGGI"),
 ("[SEDANG] ", "Daftar isi: angka halaman masih statis; regenerasi TOC (References \u2192 Update Table) setelah "
   "penomoran bab/halaman final.", "SEDANG"),
 ("[RENDAH] ", "Halaman judul/meta: kolom \u201c作者姓名\u201d dan \u201c分类号/密级\u201d kosong di tabel meta \u2014 pastikan diisi "
   "sesuai template resmi sebelum submit.", "RENDAH"),
])

# ---------------- C. BAB 1 ----------------
H("4. Bab 1 \u2014 Introduction", 1)
P("Kelebihan: latar belakang kuat dan terdata (regulasi Perpres 66/2018, pangsa RON 90 >84%, gap harga ~35%, "
  "Southeast Asia gap); gap literatur dirumuskan jelas (supply-side vs demand-side asymmetry); ruang lingkup "
  "dibatasi secara eksplisit; metode dua jalur dijelaskan; technical route divisualisasi.")
B([
 ("[TINGGI] ", "Tanggal luncur tidak konsisten: \u00a71.1.1 \u201claunch of Pertamax Green 95 in July 2023\u201d vs "
   "\u00a73.2.1 \u201claunched on June 20, 2023\u201d vs revisi abstrak \u201cin June 2023\u201d. Pilih satu tanggal resmi dan "
   "gunakan di seluruh dokumen; hitung ulang semua durasi (28 vs 18 bulan).", "TINGGI"),
 ("[TINGGI] ", "Angka penghematan devisa berbeda di tiga tempat tanpa penjelasan cakupan: IDR 123 triliun "
   "untuk B35 2024 (\u00a71.2.1), USD 40,7 miliar 2020\u20132025 (\u00a74.1.2/\u00a74.3.1), US$3,68 miliar H1-2025 tanpa "
   "sumber di \u00a77.1. Tulis cakupan + sumber di setiap kemunculan.", "TINGGI"),
 ("[SEDANG] ", "\u00a71.2.1: \u201cIn the petroleum marketing domain, A 7Ps-based analysis\u201d \u2014 kapital \u201cA\u201d di tengah "
   "kalimat.", "SEDANG"),
 ("[SEDANG] ", "\u00a71.1.2 (butir kedua): \u201cdetermines needs to enhance\u201d \u2192 \u201cdetermines the need to enhance\u201d.", "SEDANG"),
 ("[SEDANG] ", "\u00a71.3.2: \u201cData analysis will be using\u201d \u2192 \u201cData analysis employs / will use\u201d.", "SEDANG"),
 ("[SEDANG] ", "Belum ada Research Questions eksplisit, padahal abstrak revisi menyebut \u201cdua pertanyaan "
   "inti\u201d. Tambahkan 2 RQ eksplisit di \u00a71.3.1 agar abstrak dan isi konsisten.", "SEDANG"),
 ("[SEDANG] ", "\u00a71.2.1 kalimat pembuka tidak menyebut peneliti/atribusi sebelum \u201cTheir study projected\u2026\u201d "
   "\u2014 perbaiki kohesi (sebutkan [2] di awal kalimat).", "SEDANG"),
 ("[RENDAH] ", "Caption \u201cTable . Indonesia Energy and Biofuel Statistics\u201d tanpa nomor (isu global penomoran).", "RENDAH"),
])

# ---------------- D. BAB 2 ----------------
H("5. Bab 2 \u2014 Related Concepts and Theoretical Aspects of Marketing", 1)
P("Kelebihan: ringkas; cakupan framework lengkap (definisi marketing, STP, 7Ps, PESTEL, 5F, SWOT); justifikasi "
  "7Ps vs 4Ps untuk retail energi disampaikan.")
B([
 ("[TINGGI] ", "\u00a72.2.1 kalimat terakhir terpotong: \u201c\u2026regulatory requirements, and corporate[1].\u201d \u2014 kata "
   "hilang setelah \u201ccorporate\u201d (mungkin \u201ccorporate strategy\u201d) dan sitasi menggantung. Lengkapi.", "TINGGI"),
 ("[SEDANG] ", "\u00a72.1.3: \u201cThe product can be describe as a multi-level concept\u201d \u2192 \u201ccan be described\u201d.", "SEDANG"),
 ("[SEDANG] ", "Judul \u201c2.2.3 SWOT Analysis Model.\u201d \u2014 titik di akhir judul (juga terjadi di 5.1.1, 5.2.4, "
   "6.2, 6.3). Hapus semua titik trailing pada heading.", "SEDANG"),
 ("[RENDAH] ", "Sebagian paragraf murni ringkasan textbook. Tambahkan 1\u20132 kalimat kaitan ke konteks fuel "
   "retail Indonesia per framework agar tidak terasa generik bagi penguji.", "RENDAH"),
])

# ---------------- E. BAB 3 ----------------
H("6. Bab 3 \u2014 Current State of Marketing & Consumer Research", 1)
P("Kelebihan: analisis 7Ps kondisi eksisting terstruktur dan kaya data internal; desain survey dipetakan ke "
  "7Ps; analisis tematik rapi dengan kutipan responden (R69, R11\u2026); threshold interpretasi dijelaskan "
  "(>=3,50 Good / 3,00\u20133,49 Moderate / <3,00 Needs Improvement); NPS dan CLI dihitung transparan; lima "
  "masalah dirumuskan dengan prioritas dan interdependensi (promotion sebagai bottleneck).")
B([
 ("[KRITIS] ", "Kontradiksi metodologis: kriteria purposive sampling (\u00a73.3.1) mensyaratkan responden tahu "
   "Pertamax Green 95 (\u201caware\u2026 or as non-users with product awareness\u201d), tetapi \u00a73.3.2 melaporkan 21,8% "
   "\u201cnever heard\u201d. Longgarkan kriteria (terima responden belum tahu) atau jelaskan bahwa temuan unaware "
   "berasal dari responden yang lolos karena \u201caware\u201d diinterpretasikan longgar \u2014 lebih baik revisi kriteria.", "KRITIS"),
 ("[KRITIS] ", "Statistik tema C2-T1 bertentangan: \u00a73.3.2 \u201c39.5% (n=49)\u201d vs \u00a73.4 Problem 4 & 5 serta \u00a73.3.2 "
   "bagian C2 menyebut \u201c44.35% (n=55)\u201d. Verifikasi coding, tentukan satu angka, dan nyatakan basis "
   "multi-response jika responden boleh memilih >1 tema.", "KRITIS"),
 ("[TINGGI] ", "Jumlah item: abstrak, \u00a71.3.2, dan \u00a73.3.2 menyebut 32 item Likert; \u00a73.3.1 menyebut \u201c35 "
   "closed-ended + 3 open-ended (total 38)\u201d. Samakan; jika 32 item + 3 open = 35, tulis jujur.", "TINGGI"),
 ("[TINGGI] ", "Volume per outlet: \u00a73.2.1 \u201c624 liters per day per outlet\u201d vs \u00a73.2.2/\u00a76.1/\u00a77.1 \u201c406\u2192292 "
   "L/day (2025)\u201d. Dua angka beda makna \u2014 pastikan periode dan definisi (rata-rata sejak luncur vs Nov "
   "2025) dijelaskan, atau samakan.", "TINGGI"),
 ("[TINGGI] ", "\u00a73.4 Problem 1: \u201cmarketed for 18 months\u201d \u2014 salah; sejak Juni/Juli 2023 hingga survei 2025 "
   "berarti 24\u201329 bulan. Konsistenkan dengan \u201c28 months\u201d di \u00a77.1.", "TINGGI"),
 ("[TINGGI] ", "Kumulatif penjualan \u00a73.2.1: 11.746 KL (IT Jakarta) + 4.488 KL (IT Surabaya) = 16.234 KL \u2260 "
   "22.461 KL total. Jelaskan kontributor sisanya.", "TINGGI"),
 ("[SEDANG] ", "Kalimat duplikat beruntun di \u00a73.3.2 People: \u201c\u2026diminishing returns compared with improvements "
   "in Promotion and Price\u2026\u201d muncul dua kali (sebelum dan sesudah [40]). Hapus satu.", "SEDANG"),
 ("[SEDANG] ", "Kalimat terpotong di \u00a73.3.2: \u201cThe  is that Pertamina's strongest dimension\u2026\u201d (kata hilang); "
   "\u201ccreates a : consumers do not encounter\u2026\u201d; \u201caddress the \u201c\u201d identified\u201d (kutipan kosong); \u00a73.4 P1 \u201cThis "
   "is the .\u201d", "SEDANG"),
 ("[SEDANG] ", "\u00a73.3.2: \u201cstrong frontline` service delivery\u201d \u2014 karakter backtick nyasar; \u201call most items below "
   "3.00\u201d \u2192 \u201calmost all\u201d; \u201ccalculated as the meaning of B1 and B3\u201d \u2192 \u201cthe mean of\u201d.", "SEDANG"),
 ("[SEDANG] ", "Rujukan silang \u201cSection 3.3.2.5\u201d (\u00a73.3.2 Price) tidak ada \u2014 sub-bagian demografi tidak "
   "dinomori. Beri nomor sub-bagian atau ubah rujukan.", "SEDANG"),
 ("[SEDANG] ", "\u00a73.3.2 Loyalty: \u201cAmong current users (n=67), repurchase intention rises confirming\u2026\u201d \u2014 "
   "kalimat tidak selesai, nilai angka belum ditulis.", "SEDANG"),
 ("[SEDANG] ", "Saran metodologi: tambahkan Cronbach's alpha untuk skala 32 item (mudah, satu tabel), periode "
   "pengumpulan data, metode rekrutmen, dan pernyataan etika/konsen singkat. Tambahkan t-test user vs "
   "non-user untuk memperkuat klaim gap.", "SEDANG"),
 ("[SEDANG] ", "\u00a73.3.1: \u201cemploys a qualitative survey methodology\u201d bertentangan dengan klaim mixed-methods. "
   "Ganti jadi \u201cmixed-method survey (kuantitatif + kualitatif)\u201d.", "SEDANG"),
 ("[RENDAH] ", "Frame sampling 4 kota (Jabodetabek, Surabaya, Bandung, Medan) vs \u00a73.3.2 \u201c22 cities\u201d \u2014 "
   "jelaskan (mungkin responden di luar target frame ikut serta).", "RENDAH"),
])

# ---------------- F. BAB 4 ----------------
H("7. Bab 4 \u2014 Marketing Environment Analysis (PESTEL / Five Forces / SWOT)", 1)
P("Kelebihan: PESTEL lengkap per dimensi dengan data aktual; analisis kompetitor dengan data footprint "
  "(Shell ~200, BP ~50); pembahasan credence good menajamkan logika buyer power; internal capability "
  "assessment dan SWOT-TOWS terintegrasi; narasi \u201cbest-kept-secret syndrome\u201d kuat.")
B([
 ("[KRITIS] ", "\u00a74.3.2 dan \u00a74.3.3 TIDAK PUNYA JUDUL (kosong di daftar isi dan teks), padahal \u00a77.1 dan \u00a76 "
   "merujuk \u201cSection 4.3.3\u201d. Pulihkan judul sesuai isi: 4.3.2 \u2248 \u201cMarketing Capability Gap\u201d (paragraf soal "
   "fungsi marketing warisan monopoli) dan 4.3.3 \u2248 \u201cBrand Equity and Distribution Assets\u201d.", "KRITIS"),
 ("[TINGGI] ", "\u00a74.3.3 diakhiri \u201cTable .\u201d tanpa judul dan tanpa isi numerik yang jelas \u2014 lengkapi caption "
   "dan isi tabel.", "TINGGI"),
 ("[SEDANG] ", "\u201ci.e Shell\u201d tanpa titik \u2192 \u201ci.e., Shell\u201d (\u00a74.2.1 dua kali; juga \u00a73.4 P1/P2). "
   "\u201cNon-Pertamina brands (i.e Shell) is Pertamax\u2026\u201d \u2192 \u201care\u201d.", "SEDANG"),
 ("[SEDANG] ", "\u201c\u2026as of nov-2025\u201d (\u00a74.4.2) \u2192 \u201cNovember 2025\u201d.", "SEDANG"),
 ("[SEDANG] ", "Caption gambar \u201cFigure . Indonesia Biofuel Consumption\u2026 (2018\u20132030\u201d \u2014 tanda kurung tutup "
   "hilang.", "SEDANG"),
 ("[SEDANG] ", "\u00a74.4 pembuka: SWOT dikutip dari [10] (paper Resource-Based View) \u2014 nomor sitasi tampak "
   "keliru; cek pemetaan referensi.", "SEDANG"),
 ("[SEDANG] ", "\u00a74.1.4: keluhan MyPertamina tanpa sitasi (Bab 3 memakai [31] untuk hal serupa) \u2014 samakan.", "SEDANG"),
 ("[RENDAH] ", "\u00a74.3.1: konversi satuan perlu footnote jelas: \u201c4 miliar liter\u201d = 4 juta KL agar tidak "
   "dibaca 4.000.000 KL vs produksi 160.946 KL.", "RENDAH"),
])

# ---------------- G. BAB 5 ----------------
H("8. Bab 5 \u2014 Marketing Strategy Design (STP & 7P)", 1)
P("Kelebihan: segmentasi 4 variabel + 4 segment terintegrasi dengan ukuran dan usage rate; pemilihan target "
  "utama dijustifikasi 4 alasan; positioning statement formal lengkap dengan hierarki pesan 3 tier; semua "
  "strategi 7P punya rekomendasi-implementasi-KPI; cost-benefit analysis (titik impas 0,25\u20130,78 juta "
  "konsumen) dan risk matrix (likelihood-impact-mitigasi) menambah kredibilitas; prinsip sekuens "
  "\u201cpromotion mendahului/menyertai distribusi\u201d logis.")
B([
 ("[KRITIS] ", "Kalimat terpotong: \u00a75.2.1 \u201cSecond, Pertamina should\u00a0The current generic messaging\u2026\u201d (kata "
   "hilang setelah \u201cshould\u201d \u2014 kemungkinan \u201cdevelop vehicle-segment communication\u201d); \u00a75.1.3 \u201cShell V-Power) "
   "due to Pertamax Green 95's\u2026\u201d (predikat hilang); \u00a74.4 \u201cis the : without resolving\u201d.", "KRITIS"),
 ("[TINGGI] ", "Timeline target 30% coverage tidak sinkron: \u00a75.2.3 Phase 1 (0\u201312 bln) densifikasi + Phase 2 "
   "(12\u201336 bln) ekspansi luar Jawa, tetapi \u00a76.1 dan \u00a77.1 menyatakan \u201c30% \u2248 1.950 SPBU within 24 months\u201d. "
   "Samakan: mis. target 1.950 SPBU dalam 24 bulan (fase inti), perluasan luar Jawa 36 bulan.", "TINGGI"),
 ("[TINGGI] ", "Mean 7Ps tidak konsisten antar-bab: Process 3,869 (Bab 3/4) vs 3,891 (\u00a75.1 pembuka); People "
   "3,762 (\u00a73.3.2) vs 3,814 (\u00a75.1). Audit semua mean dimensi dari tabel statistik dan pakai satu versi.", "TINGGI"),
 ("[TINGGI] ", "Proyeksi \u201c7Ps 3,449 \u2192 3,65\u20133,70 dalam 12 bulan\u201d dan \u201cCLI 3,552 \u2192 3,80\u201d tidak disertai "
   "dasar perhitungan. Tambahkan rantai asumsi (target per-item Tabel 5.7, benchmark, atau elastisitas) "
   "\u2014 ini pertanyaan pertama penguji.", "TINGGI"),
 ("[SEDANG] ", "\u00a75.1.1: \u201cFour segmentation variables are applied: demographic\u2026 behavioral\u2026 psychographic\u201d "
   "\u2014 disebut empat tapi dirinci tiga (geographic ada di teori Bab 2 tetapi tidak dipakai) \u2014 rapikan "
   "frasa atau tambahkan variabel keempat.", "SEDANG"),
 ("[SEDANG] ", "\u00a75.2.2: \u201cMonroe's (2003)\u201d vs daftar pustaka [37] Monroe 1990 \u2014 samakan tahun.", "SEDANG"),
 ("[SEDANG] ", "\u201cThe Porter's Five Forces find that\u2026\u201d (\u00a75.1) \u2192 \u201cThe Five Forces analysis finds\u2026\u201d; "
   "konsistenkan British/American spelling (prioritise vs prioritize).", "SEDANG"),
 ("[SEDANG] ", "Trailing period pada judul: \u201c5.1.1 Market Segmentation.\u201d, \u201c5.2.4 Promotion Strategy.\u201d.", "SEDANG"),
 ("[RENDAH] ", "Segment \u201cPrice-Sensitive Young Commuters\u201d hanya 2,42% sampel \u2014 hindari over-claim; nyatakan "
   "sebagai segment pendukung dengan pilot Surabaya.", "RENDAH"),
])

# ---------------- H. BAB 6 ----------------
H("9. Bab 6 \u2014 Marketing Strategy Implementation Guarantees", 1)
P("Kelebihan: empat sistem jaminan (manajemen, SDM, finansial, TI) realistis untuk konteks BUMN; komite "
  "VP-level, budget line khusus RKAP, dan target KPI dihubungkan ke Tabel 5.7; keterkaitan dengan risk "
  "register \u00a75.3 dipertahankan.")
B([
 ("[KRITIS] ", "Fragmen kalimat: \u00a76.1 \u201cwill proceed as\u00a0rather than as isolated initiatives\u201d (kata hilang); "
   "\u00a76.2 \u201cwill generate\u00a0across all seven 7P elements\u201d; \u00a76.3 \u201cpromotional credit provisioning (,\u201d \u2014 "
   "tanda kurung sisa yang tidak selesai.", "KRITIS"),
 ("[TINGGI] ", "\u201c1.950 gas station\u201d = 30% dari 6.000+ = 1.800, bukan 1.950 \u2014 jelaskan basis (mungkin "
   "6.500 jaringan). Tambah footnote.", "TINGGI"),
 ("[SEDANG] ", "\u201c177 gas station\u201d berulang memakai singular untuk jamak \u2014 konsistenkan \u201cgas stations\u201d di "
   "seluruh dokumen (isu bahasa global: ratusan kemunculan \u201cgas station\u201d tanpa -s).", "SEDANG"),
 ("[RENDAH] ", "Lengkapi ringkasan biaya per program di \u00a76.3 agar total ekspektasi tercocokkan dengan "
   "estimasi IDR 21\u201332 miliar dari \u00a75.1.4.", "RENDAH"),
])

# ---------------- I. BAB 7 ----------------
H("10. Bab 7 \u2014 Research Conclusions and Future Prospects", 1)
P("Kelebihan: simpulan menjawab tujuan riset; sintesis lima masalah ke strategi ringkas dan akurat; arah "
  "riset masa depan (longitudinal, SEM, komparasi ASEAN) tepat.")
B([
 ("[KRITIS] ", "Kalimat terakhir thesis TIDAK SELESAI: \u201c\u2026Pertamina must continuously evaluate market "
   "conditions, adjusting and optimizing its\u201d \u2014 selesaikan (mis. \u201cits marketing mix and resource "
   "allocation\u201d) dan tutup dengan kalimat penutup yang kuat.", "KRITIS"),
 ("[KRITIS] ", "\u00a77.1: \u201cno competitor can .\u201d \u2014 kata hilang (kemungkinan \u201ccredibly replicate\u201d).", "KRITIS"),
 ("[TINGGI] ", "\u00a77.2: \u201cforeign-brand competitors cannot .\u201d \u2014 kata hilang; juga \u201cUS$3,68 miliar FX savings "
   "H1-2025\u201d tanpa sitasi dan berbeda basis dengan angka di Bab 1/4 (lihat audit).", "TINGGI"),
 ("[SEDANG] ", "\u201cafter 28 months of availability\u201d \u2014 pastikan konsisten dengan tanggal luncur yang "
   "distandarkan di Bab 1/3.", "SEDANG"),
])

# ---------------- J. REFERENCES & APPENDIX ----------------
H("11. References dan Appendix", 1)
B([
 ("[KRITIS] ", "[58] masih placeholder: \u201cNguyen [First name/initials], [Additional authors]\u2026 [Journal to be "
   "confirmed], 2021\u201d. Lengkapi data riil atau hapus; cek apakah [58] dirujuk in-text di mana pun \u2014 jika "
   "tidak, hapus agar tidak muncul di daftar.", "KRITIS"),
 ("[TINGGI] ", "Lampiran kuesioner tidak lengkap: Section A hanya memuat judul A1\u2013A7 TANPA 32 item "
   "(P1.1\u2013P7.4); Section D berhenti di D4 (D5\u2013D8 hilang). \u00a73.3.1 menyatakan \u201ccomplete questionnaire provided "
   "in Appendix A\u201d \u2014 lampirkan versi penuh.", "TINGGI"),
 ("[TINGGI] ", "Penomoran silang: teks merujuk \u201cTable 3.6\u20133.9, 4.4\u20134.6, 5.7\u201d tetapi caption tidak bernomor. "
   "Perbaiki pembungkus dengan penomoran bab (Table 4.x, Figure 4.x).", "TINGGI"),
 ("[TINGGI] ", "Format campuran: [20] dan [26] memakai karakter China \u201c15 版\u201d/\u201c12 版\u201d di referensi Inggris; "
   "[6] Bloomberg Technoz ditandai [M] (buku) padahal artikel berita media \u2192 [EB/OL]; [38] tanpa kota/"
   "penerbit lengkap \u2192 samakan semua dengan satu gaya (GB/T 7714).", "TINGGI"),
 ("[SEDANG] ", "[48] tanggal akses \u201c[2026-06-04]\u201d lebih muda dari konteks penulisan \u2014 periksa; beberapa "
   "entri URL tanpa tanggal akses \u2014 lengkapi.", "SEDANG"),
 ("[RENDAH] ", "Konsistensi bahasa judul regulasi Indonesia (Peraturan Menteri\u2026 vs ESDM Regulation\u2026) \u2014 "
   "pakai satu pola dengan terjemahan dalam kurung.", "RENDAH"),
])

# ---------------- K. AUDIT TABLE ----------------
H("12. Tabel Audit Konsistensi Angka", 1)
P("Semua konflik ini harus diselesaikan ke SATU versi sebelum submit; penguji membaca lintas-bab.", italic=True)
T(
 ["#", "Variabel", "Nilai yang bertentangan", "Lokasi", "Rekomendasi"],
 [
  ["1", "Tanggal luncur", "Juli 2023 | 20 Juni 2023 | Juni 2023", "\u00a71.1.1 | \u00a73.2.1 | abstrak", "Pilih satu tanggal resmi; pakai di mana-mana"],
  ["2", "Durasi di pasar", "28 bulan | 18 bulan", "\u00a77.1 | \u00a73.4 P1", "\u201c18\u201d salah \u2192 28\u201329"],
  ["3", "Jumlah SPBU", "160 (end-2024) | 170 (Nov-2025) | 177", "\u00a71.1.1 | \u00a71.1.1 | abstrak, \u00a73.2.3, \u00a74.4.2, \u00a76.1", "Audit data resmi; satu angka"],
  ["4", "Volume per outlet", "624 L/hari | 406\u2192292 L/hari (2025)", "\u00a73.2.1 | \u00a73.2.2, \u00a76.1, \u00a77.1", "Jelaskan periode/definisi; samakan"],
  ["5", "Kumulatif penjualan", "22.461 KL total vs 11.746+4.488 = 16.234 KL", "\u00a73.2.1", "Jelaskan kontributor sisanya"],
  ["6", "Item kuesioner", "32 Likert | 35 closed + 3 open (38)", "abstrak, \u00a71.3.2, \u00a73.3.2 | \u00a73.3.1", "Samakan; lampirkan item lengkap"],
  ["7", "Statistik C2-T1", "39,5% (n=49) | 44,35% (n=55)", "\u00a73.3.2 | \u00a73.3.2 C2 & \u00a73.4 P4/P5", "Verifikasi coding; satu angka"],
  ["8", "Mean Process (P6)", "3,869 | 3,891", "\u00a73.3.2, \u00a73.4 | \u00a75.1", "Cek tabel statistik; satu versi"],
  ["9", "Mean People (P5)", "3,762 | 3,814", "\u00a73.3.2 | \u00a75.1", "Cek tabel statistik; satu versi"],
  ["10", "Kota sampel", "Frame 4 kota | \u201c22 cities\u201d", "\u00a73.3.1 | \u00a73.3.2", "Jelaskan; atau revisi frame"],
  ["11", "Kriteria sampel", "Wajib \u201caware\u201d | 21,8% \u201cnever heard\u201d", "\u00a73.3.1 | \u00a73.3.2", "Longgarkan kriteria"],
  ["12", "FX savings", "IDR 123T (2024) | USD 40,7B (2020\u201325) | US$3,68B (H1-25)", "\u00a71.2.1 | \u00a74.1.2/\u00a74.3.1 | \u00a77.1", "Cakupan + sumber tiap angka"],
  ["13", "Monroe", "(2003) | [37] 1990", "\u00a75.2.2 | Referensi", "Samakan tahun"],
  ["14", "Penomoran tabel", "In-text 3.6\u20133.9/4.4\u20134.6/5.7 | caption \u201cTable .\u201d", "seluruh dokumen", "Isi numbering caption"],
  ["15", "Target 30%", "24 bulan | fase 0\u201312 + 12\u201336 bln", "\u00a76.1, \u00a77.1 | \u00a75.2.3", "Samakan timeline dan basis 1.950"],
 ],
 widths=[0.8, 3.2, 4.2, 3.6, 4.0],
)

# ---------------- L. ACTION PLAN ----------------
H("13. Rencana Aksi 10 Langkah (Urutan Disarankan)", 1)
B([
 "1)  Accept All Changes di Word; simpan sebagai file bersih baru (_v3_clean.docx); kirim versi itu ke pembimbing.",
 "2)  Scan ulang semua lokasi \u201ckalimat terpotong\u201d yang terdaftar (Bab 2\u20137); tulis ulang secara manual. Baca keras seluruh dokumen sekali.",
 "3)  Abstrak: pastikan versi Mandarin terisi penuh, versi Inggris utuh (perbaiki kalimat 1 dan paragraf 5), kata kunci 2 bahasa.",
 "4)  Beri nomor SEMUA Tabel & Gambar (penomoran per-bab) sampai cocok dengan rujukan in-text (3.6\u20133.9, 4.4\u20134.6, 5.7 dll).",
 "5)  Audit konsistensi angka (Tabel 12): tanggal, SPBU, volume, item, %, mean, timeline \u2014 satu versi.",
 "6)  Lengkapi Lampiran kuesioner: 32 item 7Ps lengkap + D5\u2013D8.",
 "7)  Bereskan Referensi: [58] placeholder, [20]/[26] \u201c版\u201d, [6] [M]\u2192[EB/OL], Monroe 1990, akses URL.",
 "8)  Pulihkan judul \u00a74.3.2/\u00a74.3.3 dan \u201cTable .\u201d tanpa judul.",
 "9)  (Opsional, penguatan) Cronbach's alpha, periode survei, t-test user vs non-user, RQ eksplisit di \u00a71.3.1.",
 "10) Cleaning akhir: titik trailing di heading, \u201ci.e.,\u201d, \u201cNovember 2025\u201d, backtick, \u201cgas stations\u201d jamak, spell-check.",
])

# ---------------- M. SIDANG ----------------
H("14. Persiapan Sidang \u2014 Pertanyaan yang Kemungkinan Muncul", 1)
B([
 "Mengapa n=124 dan purposive sampling? Bagaimana generalisasi dibatasi \u2014 siapkan kalimat \u201cexploratory\u201d.",
 "Bagaimana proyeksi 3,449 \u2192 3,65\u20133,70 diturunkan? (siapkan rantai asumsi per KPI).",
 "Bagaimana bisa ada 21,8% \u201cnever heard\u201d padahal kriteria sampel mensyaratkan aware?",
 "Mengapa tidak ada uji statistik inferensial (alpha, t-test, regresi)?",
 "Apa dasar penetapan target item (mis. P4.1 2,113 \u2192 3,0)?",
 "Bagaimana menjawab perbedaan angka antar-bab (bawa daftar audit ini sebagai lampiran)?",
])

# ---------------- N. CLOSING ----------------
H("15. Kesimpulan Umum", 1)
P("Draft ini secara substansi sudah pada level yang layak untuk tahap akhir: temuan inti jelas, bukti "
  "tertriangulasi, strategi lengkap dengan KPI dan analisis biaya-manfaat. Yang tersisa dominan pekerjaan "
  "\u201cpembersihan\u201d: accept revisions, konsistensi angka, penomoran tabel, kalimat yang terpotong, dan "
  "kelengkapan lampiran. Perkiraan 1\u20132 minggu kerja terfokus sudah cukup untuk menghasilkan versi bersih "
  "yang siap dikirim ke pembimbing dan dipresentasikan di sidang. Substansi tidak perlu dirombak.")

doc.save(OUT)
print("saved:", OUT)