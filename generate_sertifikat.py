#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generator Sertifikat Peringkat Umum Kelas X
SMK Negeri 1 Maluku Tengah - Semester Ganjil TP 2026/2027

Alur:
  1. Ambil data ledger dari Google Sheet (CSV export), file CSV lokal, atau web app Apps Script (code.gs)
  2. Parse + validasi (header dideteksi otomatis, baris kosong dilewati, peringkat diperiksa ulang)
  3. Pilih siswa peringkat umum 1..N (default 3, seri ikut dicetak)
  4. Hasilkan HTML sertifikat A4 landscape (satu halaman per siswa) dan, opsional, PDF

Contoh:
  python generate_sertifikat.py                                  # ambil langsung dari Google Sheet
  python generate_sertifikat.py --csv ledger.csv --top 3
  python generate_sertifikat.py --json "https://script.google.com/macros/s/XXXX/exec"
  python generate_sertifikat.py --top 3 --tanggal 2026-12-19 --pdf
"""
import argparse
import base64
import csv
import datetime as dt
import html
import io
import json
import math
import re
import sys
from pathlib import Path

# ----------------------------------------------------------------------------
# KONFIGURASI SEKOLAH
# ----------------------------------------------------------------------------
SHEET_ID = "1NOWqcfl-2nWKkJurflo6eJrr2g7Osb68oIu03mTmbcM"
GID = "705250680"

KOP = {
    "baris1": "PEMERINTAH PROVINSI MALUKU",
    "baris2": "DINAS PENDIDIKAN DAN KEBUDAYAAN",
    "baris3": "SMK NEGERI 1 MALUKU TENGAH",
    "alamat": "Jl. Sultan Hasanuddin, Telp. (0914) 21354 Kode Pos 97515",
    "email": "smkn1malukutengah@gmail.com",
}
KEPSEK = {
    "jabatan": "Kepala SMK Negeri 1 Maluku Tengah",
    "nama": "Siani La Jamaludin, S.Pd., M.Pd.",
    "nip": "197605152008011023",
}
TAHUN_PELAJARAN = "2026/2027"
SEMESTER = "Ganjil"

BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
         "Agustus", "September", "Oktober", "November", "Desember"]
ROMAWI_BULAN = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"]


# ----------------------------------------------------------------------------
# UTILITAS
# ----------------------------------------------------------------------------
def norm(s):
    """Normalisasi teks header: huruf kecil, spasi/baris baru dirapikan."""
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


def to_num(v):
    """Ubah teks angka (format Indonesia atau Inggris) menjadi float; None jika bukan angka."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("\u00a0", "")
    if not s or s in {"-", "#N/A", "#DIV/0!", "#REF!", "#VALUE!"}:
        return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):      # 1.234,56
            s = s.replace(".", "").replace(",", ".")
        else:                                 # 1,234.56
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def fmt_id(x, desimal=2):
    """Format angka gaya Indonesia: 75,58."""
    return f"{x:.{desimal}f}".replace(".", ",")


def romawi(n):
    pairs = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
             (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
    out = ""
    for val, sym in pairs:
        while n >= val:
            out += sym
            n -= val
    return out


def label_kelas(kode):
    """XAKL -> 'X AKL', XTJKT1 -> 'X TJKT 1'."""
    m = re.match(r"^X([A-Za-z]+?)\s*(\d*)$", str(kode).strip())
    if not m:
        return str(kode).strip()
    prog, no = m.groups()
    return f"X {prog.upper()}" + (f" {no}" if no else "")


def rapikan_nama(nama):
    nama = re.sub(r"\s+", " ", nama).strip()
    if nama.isupper() or nama.islower():
        nama = nama.title()
    return nama


def kelas_nama(nama):
    """Kelas CSS untuk mengecilkan huruf nama yang panjang agar tetap satu baris."""
    n = len(rapikan_nama(nama))
    return " xlong" if n > 38 else (" long" if n > 27 else "")


def tanggal_indonesia(d):
    return f"{d.day} {BULAN[d.month - 1]} {d.year}"


# ----------------------------------------------------------------------------
# PENGAMBILAN DATA
# ----------------------------------------------------------------------------
def ambil_csv_online(sheet_id, gid):
    import requests
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    r = requests.get(url, timeout=60)
    teks = r.content.decode("utf-8", errors="replace")
    if r.status_code != 200 or teks.lstrip().lower().startswith(("<!doctype", "<html")):
        raise RuntimeError(
            "Google Sheet tidak bisa diunduh sebagai CSV (kemungkinan belum dibagikan 'Siapa saja yang memiliki link').\n"
            "Solusi: (a) ubah akses berbagi menjadi 'Anyone with the link - Viewer', atau\n"
            "        (b) File > Download > CSV lalu jalankan dengan --csv namafile.csv, atau\n"
            "        (c) pasang code.gs sebagai Web App lalu jalankan dengan --json URL_WEBAPP."
        )
    return teks


def baca_baris_csv(teks):
    return list(csv.reader(io.StringIO(teks)))


def parse_ledger(rows):
    """Parse baris ledger (list of list). Header dicari otomatis lewat kolom 'Nama Siswa'."""
    h = None
    for i, r in enumerate(rows):
        if any(norm(c) == "nama siswa" for c in r):
            h = i
            break
    if h is None:
        raise ValueError("Header 'Nama Siswa' tidak ditemukan. Pastikan sheet/gid yang dipilih adalah Ledger Gabungan.")

    head = [norm(c) for c in rows[h]]

    def kolom(*kunci, mulai_dengan=False):
        for idx, c in enumerate(head):
            if mulai_dengan and any(c.startswith(k) for k in kunci):
                return idx
            if not mulai_dengan and c in kunci:
                return idx
        return None

    c_kelas = kolom("kelas")
    c_nis = kolom("nis")
    c_nama = kolom("nama siswa")
    c_jml = kolom("jumlah", mulai_dengan=True)
    c_rata = kolom("rata", mulai_dengan=True)
    c_pu = kolom("peringkat umum", mulai_dengan=True)
    c_pk = kolom("peringkat kelas", mulai_dengan=True)
    wajib = {"Kelas": c_kelas, "Nama Siswa": c_nama, "Jumlah": c_jml, "Rata-Rata": c_rata, "Peringkat Umum": c_pu}
    kurang = [k for k, v in wajib.items() if v is None]
    if kurang:
        raise ValueError(f"Kolom wajib tidak ditemukan: {', '.join(kurang)}")

    # nama mapel ada di baris di bawah header
    baris_mapel = rows[h + 1] if h + 1 < len(rows) else []
    mapel_idx = list(range(c_nama + 1, c_jml))
    mapel = [re.sub(r"\s+", " ", baris_mapel[i]).strip() if i < len(baris_mapel) else f"Mapel {i}" for i in mapel_idx]

    siswa, kosong = [], 0
    for r in rows[h + 2:]:
        def sel(idx):
            return r[idx].strip() if idx is not None and idx < len(r) else ""
        nama = sel(c_nama)
        if not nama:
            continue
        rata = to_num(sel(c_rata))
        jumlah = to_num(sel(c_jml))
        nilai = {m: to_num(r[i]) if i < len(r) else None for m, i in zip(mapel, mapel_idx)}
        siswa.append({
            "kelas": sel(c_kelas),
            "nis": sel(c_nis),
            "nama": nama,
            "nilai": nilai,
            "jumlah": jumlah,
            "rata": rata,
            "pu": to_num(sel(c_pu)),
            "pk": to_num(sel(c_pk)),
            "dinilai": rata is not None and jumlah is not None,
        })
        if rata is None:
            kosong += 1
    return {"mapel": mapel, "siswa": siswa}


def dari_json(sumber):
    """Baca JSON dari web app code.gs (URL) atau file JSON lokal."""
    if re.match(r"^https?://", sumber):
        import requests
        r = requests.get(sumber, timeout=90)
        r.raise_for_status()
        data = r.json()
    else:
        data = json.loads(Path(sumber).read_text(encoding="utf-8"))
    if not data.get("ok", True):
        raise RuntimeError(f"Web app mengembalikan galat: {data.get('error')}")
    for s in data["siswa"]:
        s["dinilai"] = s.get("rata") is not None and s.get("jumlah") is not None
    return {"mapel": data.get("mapel", []), "siswa": data["siswa"]}


# ----------------------------------------------------------------------------
# VALIDASI & PERINGKAT
# ----------------------------------------------------------------------------
def validasi(data):
    """Hitung ulang peringkat, bandingkan dengan kolom sheet, dan laporkan kelengkapan."""
    siswa = data["siswa"]
    dinilai = [s for s in siswa if s["dinilai"]]
    laporan = []

    # peringkat hitung ulang (competition ranking: seri berbagi peringkat)
    urut = sorted(dinilai, key=lambda s: -s["jumlah"])
    for i, s in enumerate(urut):
        if i > 0 and math.isclose(s["jumlah"], urut[i - 1]["jumlah"], abs_tol=1e-6):
            s["pu_hitung"] = urut[i - 1]["pu_hitung"]
        else:
            s["pu_hitung"] = i + 1
    for s in dinilai:
        if s["pu"] is None:
            s["pu_final"] = s["pu_hitung"]
        else:
            s["pu_final"] = int(s["pu"])

    # kelengkapan per kelas
    per_kelas = {}
    for s in siswa:
        k = per_kelas.setdefault(s["kelas"], [0, 0])
        k[0] += 1
        k[1] += 1 if s["dinilai"] else 0
    laporan.append("Kelengkapan nilai per kelas (terisi / total siswa):")
    for k, (tot, ok) in per_kelas.items():
        tanda = "OK " if ok == tot else "!! "
        laporan.append(f"  {tanda}{k:8s} {ok:3d} / {tot:3d}")

    # selisih peringkat sheet vs hitung ulang
    for s in dinilai:
        if s["pu"] is not None and int(s["pu"]) != s["pu_hitung"]:
            laporan.append(f"  !! Peringkat umum {s['nama']} di sheet = {int(s['pu'])}, hitung ulang = {s['pu_hitung']} "
                           f"(biasanya karena data kelas lain belum lengkap)")

    # celah peringkat
    ada = sorted({s["pu_final"] for s in dinilai})
    if ada:
        hilang = [n for n in range(1, max(ada) + 1) if n not in ada]
        laporan.append(f"Peringkat umum terbaca: {ada}")
        if hilang:
            laporan.append(f"  !! Peringkat {hilang} tidak ada pada data yang terbaca -> data ledger belum lengkap / ada baris yang terlewat.")
    return laporan


def pilih_penerima(data, top):
    dinilai = [s for s in data["siswa"] if s["dinilai"]]
    pen = [s for s in dinilai if s["pu_final"] <= top]
    pen.sort(key=lambda s: (s["pu_final"], s["kelas"], s["nama"]))
    return pen


# ----------------------------------------------------------------------------
# TEMPLATE SERTIFIKAT
# ----------------------------------------------------------------------------
CSS = r"""
@page { size: A4 landscape; margin: 0; }
:root{
  --green:#0a5f3a; --green-d:#064026; --gold:#b8922f; --gold-l:#e6cf86; --red:#c4261d;
  --ink:#1c2a23; --paper:#fffdf8; --muted:#4c5a52;
}
*{box-sizing:border-box;margin:0;padding:0}
html,body{background:#cfd6d1}
body{font-family:'Cormorant Garamond','Times New Roman',Tinos,'Liberation Serif',serif;color:var(--ink);
  -webkit-print-color-adjust:exact;print-color-adjust:exact;padding-top:58px}
.toolbar{position:fixed;top:0;left:0;right:0;z-index:50;background:#063f27;color:#fff;display:flex;gap:14px;
  align-items:center;justify-content:center;padding:10px 16px;font:14px/1.3 system-ui,Segoe UI,Arial,sans-serif}
.toolbar button{background:#e6cf86;border:0;color:#3b2b00;font-weight:700;padding:8px 18px;border-radius:6px;cursor:pointer;font-size:14px}
.toolbar small{opacity:.8}
.sheet{width:297mm;height:209.6mm;position:relative;overflow:hidden;margin:8mm auto;background:
  radial-gradient(ellipse at 50% 45%,#fffffc 0%,#fdf7e6 100%);box-shadow:0 3px 18px rgba(0,0,0,.3);
  break-after:page;page-break-after:always}
main .sheet:last-child{break-after:auto;page-break-after:auto}
@media print{
  html,body{background:none;padding:0}
  .toolbar{display:none}
  .sheet{margin:0;box-shadow:none}
}
.f1{position:absolute;inset:6mm;border:1.7mm solid var(--green)}
.f2{position:absolute;inset:8.7mm;border:.55mm solid var(--gold)}
.f3{position:absolute;inset:10.2mm;border:.25mm solid var(--green)}
.corner{position:absolute;width:7.4mm;height:7.4mm;background:var(--paper);border:.6mm solid var(--gold);transform:rotate(45deg)}
.corner::after{content:"";position:absolute;inset:1.4mm;background:var(--red)}
.c1{left:2.3mm;top:2.3mm}.c2{right:2.3mm;top:2.3mm}.c3{left:2.3mm;bottom:2.3mm}.c4{right:2.3mm;bottom:2.3mm}
.watermark{position:absolute;left:50%;top:56%;width:118mm;transform:translate(-50%,-50%);opacity:.065}
.content{position:absolute;inset:12.4mm 17mm 14mm 17mm;display:flex;flex-direction:column}

/* KOP */
.kop{display:flex;align-items:center;gap:4mm}
.kop img{height:23mm;width:auto;flex:none}
.kop .sp{width:21mm;flex:none}
.kop .tx{flex:1;text-align:center;font-family:'Times New Roman',Tinos,'Liberation Serif',serif;color:#111;line-height:1.14}
.k1,.k2{font-size:12.4pt;font-weight:700;letter-spacing:.02em}
.k3{font-size:17pt;font-weight:700;letter-spacing:.015em;margin:.3mm 0 .6mm}
.k4{font-size:8.6pt;line-height:1.35}
.kop-line{height:1.9mm;border-top:1.1mm solid #111;border-bottom:.35mm solid #111;margin:2mm 0 0}

/* ISI */
.body{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;padding-top:1mm}
.title{font-family:'Playfair Display','Cormorant Garamond',Georgia,serif;font-weight:800;font-size:33pt;letter-spacing:.3em;
  color:var(--green);line-height:1;padding-left:.3em}
.subtitle{font-family:'Montserrat','Segoe UI',Arial,sans-serif;font-size:8.6pt;letter-spacing:.42em;font-weight:600;
  color:var(--gold);margin-top:2mm;padding-left:.42em}
.orn{display:flex;align-items:center;gap:3mm;margin:1.8mm 0 1.2mm;width:88mm}
.orn i{flex:1;height:.3mm;background:linear-gradient(90deg,transparent,var(--gold),transparent)}
.orn b{width:2.6mm;height:2.6mm;background:var(--red);transform:rotate(45deg);border:.4mm solid var(--gold)}
.nomor{font-size:10pt;color:var(--muted);letter-spacing:.06em}
.lead{font-size:12.6pt;font-style:italic;margin-top:2.4mm;color:var(--muted)}
.name{font-family:'Great Vibes','Playfair Display','Segoe Script',Georgia,serif;font-size:36pt;line-height:1.1;color:var(--green-d);
  margin-top:.8mm;padding:0 6mm .6mm;border-bottom:.35mm solid var(--gold);min-width:150mm}
.name.long{font-size:32pt}.name.xlong{font-size:27pt}
.idline{font-size:12pt;font-weight:600;letter-spacing:.12em;margin-top:1.8mm;color:var(--ink)}
.desc{font-size:12.3pt;margin-top:2.4mm;color:var(--muted)}
.badge{margin-top:2.6mm;display:inline-flex;align-items:center;gap:4mm;padding:1.7mm 9mm;color:#fff;
  background:linear-gradient(180deg,var(--green) 0%,var(--green-d) 100%);border:.6mm solid var(--gold);border-radius:12mm;
  box-shadow:0 0 0 .8mm var(--paper),0 0 0 1.2mm var(--gold-l)}
.badge .rk{font-family:'Playfair Display',Georgia,serif;font-weight:800;font-size:19pt;letter-spacing:.1em;color:var(--gold-l)}
.badge .rk2{font-family:'Playfair Display',Georgia,serif;font-weight:700;font-size:15pt;letter-spacing:.14em}
.desc2{font-size:12.4pt;line-height:1.32;margin-top:3mm;max-width:205mm}
.desc2 b{color:var(--green-d);font-size:14pt}
.quote{font-size:11.2pt;font-style:italic;color:var(--muted);margin-top:2mm;max-width:190mm;line-height:1.3}

/* TANDA TANGAN */
.sign-row{display:flex;align-items:flex-end;justify-content:space-between;height:38mm;padding:0 2mm 0 4mm}
.seal{width:33mm;height:38mm;position:relative}
.seal svg{width:100%;height:100%}
.sign{width:92mm;text-align:center;font-size:11.8pt;line-height:1.28}
.sign .sp{height:14.5mm}
.sign .nm{font-weight:700;font-size:12.6pt;text-decoration:underline;text-underline-offset:.7mm}
.sign .tmp{margin-bottom:.6mm}
"""

FONTS = ('<link rel="preconnect" href="https://fonts.googleapis.com">'
         '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>'
         '<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,500;0,600;0,700;1,500;1,600'
         '&family=Great+Vibes&family=Montserrat:wght@600&family=Playfair+Display:wght@700;800&display=swap" rel="stylesheet">')

GRADIEN = """
<svg width="0" height="0" style="position:absolute" aria-hidden="true">
  <defs>
    <linearGradient id="gGold" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#f6e7a8"/><stop offset=".5" stop-color="#cfa941"/><stop offset="1" stop-color="#8f6d17"/></linearGradient>
    <linearGradient id="gSilver" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#fafafa"/><stop offset=".5" stop-color="#b9bec3"/><stop offset="1" stop-color="#7c8289"/></linearGradient>
    <linearGradient id="gBronze" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#f3cfa5"/><stop offset=".5" stop-color="#c58443"/><stop offset="1" stop-color="#7d4a1b"/></linearGradient>
    <linearGradient id="gGreen" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#8fd3ae"/><stop offset=".5" stop-color="#2b8a5c"/><stop offset="1" stop-color="#0a5f3a"/></linearGradient>
  </defs>
</svg>"""


def seal_svg(peringkat):
    """Rosette bernomor peringkat (emas/perak/perunggu untuk 1-3, hijau untuk lainnya)."""
    grad = {1: "gGold", 2: "gSilver", 3: "gBronze"}.get(peringkat, "gGreen")
    cx, cy, teeth = 50, 46, 36
    pts = []
    for i in range(teeth * 2):
        r = 44 if i % 2 == 0 else 40
        a = math.pi * i / teeth
        pts.append(f"{cx + r * math.cos(a):.2f},{cy + r * math.sin(a):.2f}")
    poly = " ".join(pts)
    return f"""
<svg viewBox="0 0 100 114" xmlns="http://www.w3.org/2000/svg">
  <path d="M30 80 L22 112 L36 104 L44 114 L50 84 Z" fill="#c4261d" stroke="#8d1912" stroke-width=".8"/>
  <path d="M70 80 L78 112 L64 104 L56 114 L50 84 Z" fill="#0a5f3a" stroke="#064026" stroke-width=".8"/>
  <polygon points="{poly}" fill="url(#{grad})" stroke="#7a5b10" stroke-width=".7"/>
  <circle cx="{cx}" cy="{cy}" r="34" fill="none" stroke="#fff" stroke-opacity=".75" stroke-width="1.2"/>
  <circle cx="{cx}" cy="{cy}" r="30.5" fill="#fffdf4" fill-opacity=".92" stroke="#7a5b10" stroke-width=".8"/>
  <text x="{cx}" y="{cy + 8}" text-anchor="middle" font-family="Playfair Display, Georgia, serif" font-weight="800"
        font-size="{30 if peringkat < 10 else 24}" fill="#064026">{peringkat}</text>
  <text x="{cx}" y="{cy - 15}" text-anchor="middle" font-family="Montserrat, Arial, sans-serif" font-weight="700"
        font-size="5.2" letter-spacing="1" fill="#7a5b10">PERINGKAT</text>
  <text x="{cx}" y="{cy + 21}" text-anchor="middle" font-family="Montserrat, Arial, sans-serif" font-weight="700"
        font-size="5.2" letter-spacing="1" fill="#7a5b10">UMUM</text>
</svg>"""


def sertifikat_html(s, ctx):
    e = html.escape
    pu = s["pu_final"]
    return f"""
<section class="sheet">
  <div class="f1"></div><div class="f2"></div><div class="f3"></div>
  <span class="corner c1"></span><span class="corner c2"></span><span class="corner c3"></span><span class="corner c4"></span>
  <img class="watermark" src="{ctx['logo']}" alt="">
  <div class="content">
    <header class="kop">
      <img src="{ctx['logo']}" alt="Logo Provinsi Maluku">
      <div class="tx">
        <div class="k1">{e(KOP['baris1'])}</div>
        <div class="k2">{e(KOP['baris2'])}</div>
        <div class="k3">{e(KOP['baris3'])}</div>
        <div class="k4">{e(KOP['alamat'])}</div>
        <div class="k4">e-mail: {e(KOP['email'])}</div>
      </div>
      <div class="sp"></div>
    </header>
    <div class="kop-line"></div>

    <section class="body">
      <div class="title">SERTIFIKAT</div>
      <div class="subtitle">PENGHARGAAN PRESTASI AKADEMIK</div>
      <div class="orn"><i></i><b></b><i></i></div>
      <div class="nomor">Nomor: {e(s['nomor'])}</div>
      <div class="lead">Dengan bangga diberikan kepada:</div>
      <div class="name{kelas_nama(s['nama'])}">{e(rapikan_nama(s['nama']))}</div>
      <div class="idline">NIS {e(s['nis'] or '-')} &nbsp;&bull;&nbsp; KELAS {e(label_kelas(s['kelas']))}</div>
      <div class="desc">atas prestasinya yang membanggakan dan telah ditetapkan sebagai</div>
      <div class="badge"><span class="rk">PERINGKAT {romawi(pu)}</span><span class="rk2">UMUM KELAS X</span></div>
      <div class="desc2">dengan perolehan nilai rata-rata <b>{fmt_id(s['rata'])}</b> pada Semester {SEMESTER}<br>
        Tahun Pelajaran {TAHUN_PELAJARAN} di SMK Negeri 1 Maluku Tengah.</div>
      <div class="quote">&ldquo;Semoga prestasi ini menjadi motivasi untuk terus belajar, berkarya, berakhlak mulia,
        serta mengharumkan nama sekolah, daerah, dan bangsa.&rdquo;</div>
    </section>

    <footer class="sign-row">
      <div class="seal">{seal_svg(pu)}</div>
      <div class="sign">
        <div class="tmp">{e(ctx['tempat'])}, {e(ctx['tanggal_teks'])}</div>
        <div>{e(KEPSEK['jabatan'])},</div>
        <div class="sp"></div>
        <div class="nm">{e(KEPSEK['nama'])}</div>
        <div>NIP. {e(KEPSEK['nip'])}</div>
      </div>
    </footer>
  </div>
</section>"""


def bangun_html(penerima, ctx):
    halaman = "\n".join(sertifikat_html(s, ctx) for s in penerima)
    return f"""<!DOCTYPE html>
<html lang="id"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Sertifikat Peringkat Umum Kelas X - SMK Negeri 1 Maluku Tengah</title>
{FONTS}
<style>{CSS}</style></head>
<body>
{GRADIEN}
<div class="toolbar"><button onclick="window.print()">Cetak / Simpan PDF</button>
<small>{len(penerima)} sertifikat &bull; Kertas A4 Landscape, Margin: None, aktifkan &ldquo;Background graphics&rdquo;</small></div>
<main>
{halaman}
</main>
</body></html>"""


# ----------------------------------------------------------------------------
# UTAMA
# ----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Generator sertifikat peringkat umum kelas X SMKN 1 Maluku Tengah")
    sumber = ap.add_mutually_exclusive_group()
    sumber.add_argument("--csv", help="File CSV hasil unduhan sheet (File > Download > CSV)")
    sumber.add_argument("--json", help="URL Web App code.gs atau file JSON hasil ekspor")
    ap.add_argument("--sheet-id", default=SHEET_ID)
    ap.add_argument("--gid", default=GID)
    ap.add_argument("--top", type=int, default=3, help="Cetak peringkat umum 1..N (default 3; seri ikut dicetak)")
    ap.add_argument("--logo", default="logo_provinsi_maluku.png")
    ap.add_argument("--tempat", default="Masohi", help="Tempat penerbitan (default Masohi)")
    ap.add_argument("--tanggal", default=dt.date.today().isoformat(), help="Tanggal sertifikat YYYY-MM-DD (default hari ini)")
    ap.add_argument("--nomor-awal", type=int, default=1)
    ap.add_argument("--format-nomor", default="{n:03d}/SERT-PU/SMKN1MT/{bulan_romawi}/{tahun}",
                    help="Pola nomor; variabel: n, bulan_romawi, tahun")
    ap.add_argument("--out", default="sertifikat_peringkat_umum_kelas_X.html")
    ap.add_argument("--data-out", default="data_ledger_terbaca.json", help="Simpan hasil parse penuh untuk diperiksa")
    ap.add_argument("--pdf", action="store_true", help="Buat juga PDF (butuh: pip install playwright && playwright install chromium)")
    a = ap.parse_args()

    # 1. ambil data
    if a.json:
        data = dari_json(a.json)
        print(f"[sumber] JSON: {a.json}")
    else:
        if a.csv:
            teks = Path(a.csv).read_text(encoding="utf-8-sig")
            print(f"[sumber] CSV lokal: {a.csv}")
        else:
            teks = ambil_csv_online(a.sheet_id, a.gid)
            print(f"[sumber] Google Sheet {a.sheet_id} (gid={a.gid})")
        data = parse_ledger(baca_baris_csv(teks))

    print(f"[parse] {len(data['siswa'])} siswa terbaca, {sum(1 for s in data['siswa'] if s['dinilai'])} sudah memiliki nilai")

    # 2. validasi
    for baris in validasi(data):
        print(baris)
    Path(a.data_out).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[simpan] data lengkap -> {a.data_out}")

    # 3. penerima
    penerima = pilih_penerima(data, a.top)
    if not penerima:
        sys.exit("Tidak ada siswa dengan peringkat umum yang memenuhi syarat. Periksa data ledger.")
    tgl = dt.date.fromisoformat(a.tanggal)
    for i, s in enumerate(penerima):
        s["nomor"] = a.format_nomor.format(n=a.nomor_awal + i, bulan_romawi=ROMAWI_BULAN[tgl.month - 1], tahun=tgl.year)
    print("[penerima]")
    for s in penerima:
        print(f"  Peringkat {s['pu_final']}: {s['nama']} ({label_kelas(s['kelas'])}) rata-rata {fmt_id(s['rata'])}  no. {s['nomor']}")

    # 4. HTML
    logo_path = Path(a.logo)
    if not logo_path.exists():
        sys.exit(f"Logo tidak ditemukan: {logo_path}")
    logo = "data:image/png;base64," + base64.b64encode(logo_path.read_bytes()).decode()
    ctx = {"logo": logo, "tempat": a.tempat, "tanggal_teks": tanggal_indonesia(tgl)}
    Path(a.out).write_text(bangun_html(penerima, ctx), encoding="utf-8")
    print(f"[selesai] {a.out}")

    # 5. PDF opsional
    if a.pdf:
        from playwright.sync_api import sync_playwright
        pdf_path = str(Path(a.out).with_suffix(".pdf"))
        with sync_playwright() as p:
            b = p.chromium.launch()
            pg = b.new_page()
            pg.goto(Path(a.out).resolve().as_uri(), wait_until="networkidle")
            pg.pdf(path=pdf_path, format="A4", landscape=True, print_background=True, prefer_css_page_size=True)
            b.close()
        print(f"[selesai] {pdf_path}")


if __name__ == "__main__":
    main()
