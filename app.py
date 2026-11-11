"""
Aplikasi Perencanaan Geometrik Jalan
  1. Denah trase tikungan (FC / S-C-S / S-S)
  2. Diagram superelevasi
  3. Profil memanjang (garis rencana + lengkung vertikal + Animasi Simulasi Mobil)
  4. Ekspor Excel Ber-Formula (Perhitungan Tikungan, Lengkung Vertikal, & Galian/Timbunan)

Acuan: TPGJAK Bina Marga 1997 (Tata Cara Perencanaan Geometrik Jalan Antar Kota)
Jalankan: streamlit run app.py
"""
import io
import math
import time

import matplotlib.pyplot as plt
from matplotlib.transforms import Affine2D
import matplotlib.patches as patches
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st
from matplotlib.ticker import FuncFormatter

st.set_page_config(page_title="Perencanaan Geometrik Jalan", layout="wide")

# ----------------------------------------------------------------------------
# Tabel acuan (TPGJAK 1997): kecepatan rencana -> R min tanpa lengkung peralihan,
# kelandaian maksimum
# ----------------------------------------------------------------------------
V_TAB = [20, 30, 40, 50, 60, 80, 100, 120]
R_FC_TAB = [60, 130, 250, 350, 500, 900, 1500, 2500]
G_MAX_TAB = [10, 10, 10, 9, 8, 6, 5, 3]


def sta_str(s, d=2):
    """Format stasioning km+meter, contoh 0+250.00"""
    return f"{int(s // 1000)}+{s % 1000:0{4 + d}.{d}f}"


def f_max(V):
    return -0.00065 * V + 0.192 if V < 80 else -0.00125 * V + 0.24


def r_min(V, emax):
    return V**2 / (127 * (emax + f_max(V)))


# ----------------------------------------------------------------------------
# Perhitungan tikungan horizontal
# ----------------------------------------------------------------------------
def hitung_tikungan(V, delta, emax, en, Rc, sta_pi):
    fm = f_max(V)
    Rmin = r_min(V, emax)
    Dmax = 181913.53 * (emax + fm) / V**2
    D = 1432.39 / Rc
    e = -emax * D**2 / Dmax**2 + 2 * emax * D / Dmax  # Bina Marga
    e = max(e, en)

    # Panjang lengkung peralihan minimum
    re = 0.035 if V <= 70 else 0.025  # m/m/detik
    Ls_waktu = V * 3 / 3.6
    Ls_shortt = 0.022 * V**3 / (Rc * 0.4) - 2.727 * V * e / 0.4
    Ls_re = (e + en) * V / (3.6 * re)
    Ls_min = max(Ls_waktu, Ls_shortt, Ls_re)
    Ls_pakai = math.ceil(Ls_min / 5) * 5

    R_fc = float(np.interp(V, V_TAB, R_FC_TAB))
    d2 = math.radians(delta / 2)
    r = dict(V=V, delta=delta, emax=emax, en=en, Rc=Rc, fm=fm, Rmin=Rmin,
             Dmax=Dmax, D=D, e=e, Ls_waktu=Ls_waktu, Ls_shortt=Ls_shortt,
             Ls_re=Ls_re, Ls_min=Ls_min, R_fc=R_fc, sta_pi=sta_pi, Ls_pakai=Ls_pakai)

    if Rc >= R_fc:
        Tc = Rc * math.tan(d2)
        Ec = Tc * math.tan(d2 / 2)
        Lc = math.radians(delta) * Rc
        r.update(jenis="Full Circle (FC)", Ls=0.0, Ls_sup=Ls_pakai, ths=0.0,
                 dc=delta, Lc=Lc, T=Tc, E=Ec, Xs=0, Ys=0, p=0, k=0)
    else:
        ths = math.degrees(Ls_pakai / (2 * Rc))
        dc = delta - 2 * ths
        Lc = math.radians(dc) * Rc
        Ls = float(Ls_pakai)
        jenis = "Spiral-Circle-Spiral (S-C-S)"
        if Lc < 25:  # Lc terlalu pendek -> Spiral-Spiral
            jenis = "Spiral-Spiral (S-S)"
            ths = delta / 2
            Ls = math.radians(ths) * 2 * Rc
            Lc, dc = 0.0, 0.0
        Xs = Ls * (1 - Ls**2 / (40 * Rc**2))
        Ys = Ls**2 / (6 * Rc)
        p = Ys - Rc * (1 - math.cos(math.radians(ths)))
        k = Xs - Rc * math.sin(math.radians(ths))
        Ts = (Rc + p) * math.tan(d2) + k
        Es = (Rc + p) / math.cos(d2) - Rc
        r.update(jenis=jenis, Ls=Ls, Ls_sup=Ls, ths=ths, dc=dc, Lc=Lc, T=Ts,
                 E=Es, Xs=Xs, Ys=Ys, p=p, k=k)

    r["Ltot"] = r["Lc"] + 2 * r["Ls"]
    r["sta_awal"] = sta_pi - r["T"]
    r["sta_akhir"] = r["sta_awal"] + r["Ltot"]
    return r


def tabel_hasil(r):
    fc = r["jenis"].startswith("Full")
    a, b = ("TC", "CT") if fc else ("TS", "ST")
    rows = [
        ("Jenis tikungan", r["jenis"], ""),
        ("f maks", f"{r['fm']:.4f}", ""),
        ("R minimum", f"{r['Rmin']:.2f}", "m"),
        ("R rencana (Rc)", f"{r['Rc']:.2f}", "m"),
        ("D maks", f"{r['Dmax']:.3f}", "°"),
        ("D rencana", f"{r['D']:.3f}", "°"),
        ("e rencana", f"{r['e'] * 100:.2f}", "%"),
        ("Ls (waktu 3 detik)", f"{r['Ls_waktu']:.2f}", "m"),
        ("Ls (Shortt)", f"{r['Ls_shortt']:.2f}", "m"),
        ("Ls (tingkat perubahan kemiringan)", f"{r['Ls_re']:.2f}", "m"),
        ("Ls dipakai", f"{r['Ls_sup']:.2f}", "m"),
    ]
    if not fc:
        rows += [("θs", f"{r['ths']:.3f}", "°"), ("Δc", f"{r['dc']:.3f}", "°"),
                 ("Xs", f"{r['Xs']:.3f}", "m"), ("Ys", f"{r['Ys']:.3f}", "m"),
                 ("p", f"{r['p']:.3f}", "m"), ("k", f"{r['k']:.3f}", "m")]
    rows += [
        ("Tangen (T)", f"{r['T']:.3f}", "m"),
        ("Eksternal (E)", f"{r['E']:.3f}", "m"),
        ("Lc", f"{r['Lc']:.3f}", "m"),
        ("Panjang tikungan", f"{r['Ltot']:.3f}", "m"),
        (f"STA {a}", sta_str(r["sta_awal"]), ""),
        ("STA PI", sta_str(r["sta_pi"]), ""),
        (f"STA {b}", sta_str(r["sta_akhir"]), ""),
    ]
    return pd.DataFrame(rows, columns=["Parameter", "Nilai", "Satuan"])


# ----------------------------------------------------------------------------
# Pembuat File Excel Ber-Formula (openpyxl)
# ----------------------------------------------------------------------------
def generate_excel_with_formulas(r, P=None, df_earthwork=None):
    wb = openpyxl.Workbook()
    
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_sub = Font(name="Calibri", size=11, bold=True)
    font_title = Font(name="Calibri", size=14, bold=True, color="1F4E78")
    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_input = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    thin_border = Border(left=Side(style='thin', color='D9D9D9'),
                         right=Side(style='thin', color='D9D9D9'),
                         top=Side(style='thin', color='D9D9D9'),
                         bottom=Side(style='thin', color='D9D9D9'))
    
    ws1 = wb.active
    ws1.title = "Tikungan Horizontal"
    ws1.views.sheetView[0].showGridLines = True

    ws1["A1"] = "PERHITUNGAN GEOMETRIK TIKUNGAN HORIZONTAL (TPGJAK 1997)"
    ws1["A1"].font = font_title
    
    ws1.cell(row=3, column=1, value="1. DATA INPUT RENCANA").font = font_sub
    input_data = [
        ("Kecepatan Rencana (Vr)", r["V"], "km/jam", "B4"),
        ("Sudut Tikungan (Delta)", r["delta"], "derajat", "B5"),
        ("Superelevasi Maksimum (e_max)", r["emax"], "m/m", "B6"),
        ("Kemiringan Normal (e_n)", r["en"], "m/m", "B7"),
        ("Jari-jari Rencana (Rc)", r["Rc"], "m", "B8"),
        ("Stasioning PI (STA PI)", r["sta_pi"], "m", "B9"),
    ]
    
    ws1.cell(row=4, column=1, value="Parameter").font = font_header
    ws1.cell(row=4, column=1).fill = fill_header
    ws1.cell(row=4, column=2, value="Nilai").font = font_header
    ws1.cell(row=4, column=2).fill = fill_header
    ws1.cell(row=4, column=3, value="Satuan").font = font_header
    ws1.cell(row=4, column=3).fill = fill_header
    
    for idx, (param, val, unit, cell_ref) in enumerate(input_data, start=5):
        ws1.cell(row=idx, column=1, value=param)
        cell_v = ws1.cell(row=idx, column=2, value=val)
        cell_v.fill = fill_input
        ws1.cell(row=idx, column=3, value=unit)
        if "m/m" in unit:
            cell_v.number_format = '0.00%'
            
    ws1.cell(row=12, column=1, value="2. HASIL PERHITUNGAN BER-FORMULA").font = font_sub
    
    headers_calc = ["Deskripsi Parameter", "Formula Excel / Nilai", "Satuan", "Keterangan / Persamaan"]
    for col_num, h in enumerate(headers_calc, start=1):
        c = ws1.cell(row=13, column=col_num, value=h)
        c.font = font_header
        c.fill = fill_header

    calc_rows = [
        ("Faktor Gesekan Maksimum (f_max)", "=IF(B5<80, -0.00065*B5+0.192, -0.00125*B5+0.24)", "-", "TPGJAK 1997"),
        ("Jari-Jari Minimum (R_min)", "=B5^2 / (127 * (B7 + B14))", "m", "Vr^2 / (127*(emax+fm))"),
        ("Derajat Lengkung Maksimum (D_max)", "=181913.53 * (B7 + B14) / (B5^2)", "deg", "181913.53*(emax+fm)/Vr^2"),
        ("Derajat Lengkung Rencana (D)", "=1432.39 / B9", "deg", "1432.39 / Rc"),
        ("Superelevasi Rencana (e)", "=MAX(B8, -B7*(B17^2)/(B16^2) + 2*B7*B17/B16)", "%", "Metode Bina Marga"),
        ("Ls (Waktu 3 detik)", "=B5 * 3 / 3.6", "m", "Vr * 3 / 3.6"),
        ("Ls (Shortt)", "=0.022 * (B5^3) / (B9 * 0.4) - 2.727 * B5 * B18 / 0.4", "m", "Formula Shortt"),
        ("Ls (Kelandaian Relatif)", "=(B18 + B8) * B5 / (3.6 * IF(B5<=70, 0.035, 0.025))", "m", "Modulus kelandaian relatif"),
        ("Ls Minimum", "=MAX(B19, B20, B21)", "m", "Max dari 3 kriteria Ls"),
        ("Ls Dipakai (Pembulatan 5m)", "=CEILING(B22, 5)", "m", "Pembulatan ke atas kelipatan 5"),
        ("Sudut Lengkung Peralihan (Theta_s)", "=DEGREES(B23 / (2 * B9))", "deg", "Ls / (2 * Rc) dalam derajat"),
        ("Sudut Lengkung Lingkaran (Delta_c)", "=B6 - (2 * B24)", "deg", "Delta - 2*Theta_s"),
        ("Panjang Busur Lingkaran (Lc)", "=RADIANS(B25) * B9", "m", "Delta_c * Rc (rad)"),
        ("Jenis Tikungan", '=IF(B9>=1500, "Full Circle (FC)", IF(B26>=25, "Spiral-Circle-Spiral (S-C-S)", "Spiral-Spiral (S-S)"))', "-", "Evaluasi Kriteria Trase"),
        ("Pergeseran Tangen (Xs)", "=B23 * (1 - (B23^2) / (40 * (B9^2)))", "m", "Ls*(1 - Ls^2/(40*Rc^2))"),
        ("Pergeseran Absis (Ys)", "=(B23^2) / (6 * B9)", "m", "Ls^2 / (6*Rc)"),
        ("Penyimpangan Ordinat (p)", "=B29 - B9 * (1 - COS(RADIANS(B24)))", "m", "Ys - Rc*(1 - cos(Theta_s))"),
        ("Penyimpangan Absis (k)", "=B28 - B9 * SIN(RADIANS(B24))", "m", "Xs - Rc*sin(Theta_s)"),
        ("Panjang Tangen (T)", '=IF(ISNUMBER(SEARCH("Full", B27)), B9*TAN(RADIANS(B6/2)), (B9+B30)*TAN(RADIANS(B6/2)) + B31)', "m", "Jarak Tangen Utama"),
        ("Eksternal (E)", '=IF(ISNUMBER(SEARCH("Full", B27)), B32*TAN(RADIANS(B6/4)), (B9+B30)/COS(RADIANS(B6/2)) - B9)', "m", "Jarak Eksternal Utama"),
        ("Panjang Total Lengkung (L_tot)", "=B26 + (2 * B23)", "m", "Lc + 2*Ls"),
        ("STA Awal Tikungan (TC / TS)", "=B10 - B32", "m", "STA PI - T"),
        ("STA Akhir Tikungan (CT / ST)", "=B35 + B34", "m", "STA Awal + Ltot"),
    ]

    for idx, (desc, formula, unit, ket) in enumerate(calc_rows, start=14):
        ws1.cell(row=idx, column=1, value=desc)
        c_val = ws1.cell(row=idx, column=2, value=formula)
        ws1.cell(row=idx, column=3, value=unit)
        ws1.cell(row=idx, column=4, value=ket)
        if unit == "%":
            c_val.number_format = '0.00%'
        elif unit in ["m", "deg"]:
            c_val.number_format = '#,##0.00'
            
    if P and len(P["kurva"]) > 0:
        ws2 = wb.create_sheet(title="Lengkung Vertikal")
        ws2.views.sheetView[0].showGridLines = True
        ws2["A1"] = "PERHITUNGAN LENGKUNG VERTIKAL"
        ws2["A1"].font = font_title

        headers_v = ["STA PPV", "Elevasi PPV (m)", "g1 (%)", "g2 (%)", "A (%)", "Jh (m)", "Lv Rencana (m)", "Ev (m)", "STA PLV", "STA PTV", "Jenis Lengkung"]
        for col_num, h in enumerate(headers_v, start=1):
            c = ws2.cell(row=3, column=col_num, value=h)
            c.font = font_header
            c.fill = fill_header

        for r_idx, k in enumerate(P["kurva"], start=4):
            ws2.cell(row=r_idx, column=1, value=k["sta"]).number_format = '#,##0.00'
            ws2.cell(row=r_idx, column=2, value=k["el"]).number_format = '#,##0.00'
            ws2.cell(row=r_idx, column=3, value=k["g1"]).number_format = '0.00%'
            ws2.cell(row=r_idx, column=4, value=k["g2"]).number_format = '0.00%'
            ws2.cell(row=r_idx, column=5, value=f"=ABS(D{r_idx}-C{r_idx})").number_format = '0.00%'
            ws2.cell(row=r_idx, column=6, value=P["Jh"]).number_format = '#,##0.00'
            ws2.cell(row=r_idx, column=7, value=k["Lv"]).number_format = '#,##0.00'
            ws2.cell(row=r_idx, column=8, value=f"=E{r_idx}*100*G{r_idx}/800").number_format = '#,##0.000'
            ws2.cell(row=r_idx, column=9, value=f"=A{r_idx}-(G{r_idx}/2)").number_format = '#,##0.00'
            ws2.cell(row=r_idx, column=10, value=f"=A{r_idx}+(G{r_idx}/2)").number_format = '#,##0.00'
            ws2.cell(row=r_idx, column=11, value=f'=IF(D{r_idx}<C{r_idx},"Cembung","Cekung")')

    if df_earthwork is not None and not df_earthwork.empty:
        ws3 = wb.create_sheet(title="Galian Timbunan")
        ws3.views.sheetView[0].showGridLines = True
        ws3["A1"] = "EVALUASI GALIAN DAN TIMBUNAN TIAP 50 METER"
        ws3["A1"].font = font_title

        headers_ew = ["No", "Stasioning (STA)", "Elevasi Rencana (m)", "Elevasi Tanah Asli (m)", "Selisih Elevasi (m)", "Status Pekerjaan"]
        for col_num, h in enumerate(headers_ew, start=1):
            c = ws3.cell(row=3, column=col_num, value=h)
            c.font = font_header
            c.fill = fill_header

        for idx, row in df_earthwork.reset_index(drop=True).iterrows():
            r_i = idx + 4
            ws3.cell(row=r_i, column=1, value=idx + 1)
            ws3.cell(row=r_i, column=2, value=row["STA_val"]).number_format = '#,##0.00'
            ws3.cell(row=r_i, column=3, value=row["Elevasi rencana (m)"]).number_format = '#,##0.000'
            ws3.cell(row=r_i, column=4, value=row["Elevasi tanah (m)"]).number_format = '#,##0.000'
            ws3.cell(row=r_i, column=5, value=f"=C{r_i}-D{r_i}").number_format = '#,##0.000'
            ws3.cell(row=r_i, column=6, value=f'=IF(E{r_i}>0, "Timbunan", "Galian")')

    for sheet in wb.worksheets:
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
            for cell in row:
                if cell.value is not None:
                    cell.border = thin_border
        for col in sheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            sheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ----------------------------------------------------------------------------
# Gambar 1: denah trase tikungan
# ----------------------------------------------------------------------------
def heading(s, Ls, Lc, Rc, dl):
    Ltot = 2 * Ls + Lc
    th = np.zeros_like(s)
    if Ls > 0:
        ths = Ls / (2 * Rc)
        m = (s > 0) & (s <= Ls)
        th[m] = s[m] ** 2 / (2 * Rc * Ls)
        m = (s > Ls) & (s < Ls + Lc)
        th[m] = ths + (s[m] - Ls) / Rc
        m = (s >= Ls + Lc) & (s <= Ltot)
        th[m] = dl - (Ltot - s[m]) ** 2 / (2 * Rc * Ls)
    else:
        m = (s > 0) & (s <= Ltot)
        th[m] = s[m] / Rc
    th[s > Ltot] = dl
    return th


def fig_denah(r, arah, B):
    sg = 1 if arah == "Kiri" else -1
    Ls, Lc, Rc = r["Ls"], r["Lc"], r["Rc"]
    dl = math.radians(r["delta"])
    Ltot = r["Ltot"]
    ext = math.ceil(max(100, 1.3 * r["T"]) / 10) * 10

    s = np.unique(np.concatenate([np.arange(-ext, Ltot + ext, 0.5),
                                  [0, Ls, Ls + Lc, Ltot]]))
    th = heading(s, Ls, Lc, Rc, dl)
    dS = np.diff(s)
    x = np.concatenate([[0], np.cumsum(dS * (np.cos(th[:-1]) + np.cos(th[1:])) / 2)])
    y = np.concatenate([[0], np.cumsum(dS * (np.sin(th[:-1]) + np.sin(th[1:])) / 2)])
    i0 = np.argmin(np.abs(s))
    x, y = x - x[i0], (y - y[i0]) * sg
    ths_ = th * sg
    nx, ny = -np.sin(ths_), np.cos(ths_)

    fig, ax = plt.subplots(figsize=(11, 6.5))
    ax.plot(x + B / 2 * nx, y + B / 2 * ny, color="gray", lw=1)
    ax.plot(x - B / 2 * nx, y - B / 2 * ny, color="gray", lw=1)
    ax.plot(x, y, "k", lw=1.6, label="As jalan")

    pi = np.array([r["T"], 0.0])
    p2 = pi + ext * np.array([math.cos(dl), sg * math.sin(dl)])
    ax.plot([-ext, pi[0], p2[0]], [0, pi[1], p2[1]], "b--", lw=0.9,
            label="Garis singgung")
    ax.plot(*pi, "ks")
    ax.annotate(f"PI\nΔ = {r['delta']:.1f}°\nSTA {sta_str(r['sta_pi'])}", pi,
                textcoords="offset points", xytext=(8, -34 * sg), fontsize=8)

    fc = r["jenis"].startswith("Full")
    if fc:
        keys = [("TC", 0), ("CT", Lc)]
    elif Lc == 0:
        keys = [("TS", 0), ("SC=CS", Ls), ("ST", Ltot)]
    else:
        keys = [("TS", 0), ("SC", Ls), ("CS", Ls + Lc), ("ST", Ltot)]
    pts = {}
    for nm, sk in keys:
        px, py = np.interp(sk, s, x), np.interp(sk, s, y)
        pts[nm] = (px, py, sk)
        ax.plot(px, py, "ro")
        ax.annotate(f"{nm}\n{sta_str(r['sta_awal'] + sk)}", (px, py),
                    textcoords="offset points", xytext=(-14, 10 * sg + (-24 if sg < 0 else 0)),
                    fontsize=8, color="darkred")

    s_ref = Ls if not fc else 0
    cx = np.interp(s_ref, s, x) + Rc * np.interp(s_ref, s, nx)
    cy = np.interp(s_ref, s, y) + Rc * np.interp(s_ref, s, ny)
    ax.plot(cx, cy, "g+", ms=10)
    for nm, (px, py, _) in pts.items():
        if nm in ("SC", "CS", "SC=CS", "TC", "CT"):
            ax.plot([cx, px], [cy, py], "g:", lw=0.8)
    ax.annotate(f"R = {Rc:.0f} m", (cx, cy), textcoords="offset points",
                xytext=(6, 4), fontsize=8, color="green")

    sta0 = r["sta_awal"]
    for st_ in np.arange(math.ceil((sta0 - ext) / 50) * 50, sta0 + Ltot + ext, 50):
        sv = st_ - sta0
        if sv < s[0] or sv > s[-1]:
            continue
        px, py = np.interp(sv, s, x), np.interp(sv, s, y)
        qx, qy = np.interp(sv, s, nx), np.interp(sv, s, ny)
        ax.plot([px - (B / 2 + 2) * qx, px + (B / 2 + 2) * qx],
                [py - (B / 2 + 2) * qy, py + (B / 2 + 2) * qy], color="gray", lw=0.6)
        ax.text(px - (B / 2 + 4) * qx, py - (B / 2 + 4) * qy,
                f"{int(st_ // 1000)}+{st_ % 1000:03.0f}", fontsize=6, ha="center",
                va="center", color="gray")

    ax.set_aspect("equal")
    ax.grid(alpha=0.25)
    ax.set_xlabel("x (m)")
    ax.set_ylabel("y (m)")
    ax.set_title(f"Denah Trase Tikungan - {r['jenis']} (tikungan ke {arah.lower()})")
    ax.legend(loc="best", fontsize=8)
    return fig


# ----------------------------------------------------------------------------
# Gambar 2: diagram superelevasi
# ----------------------------------------------------------------------------
def fig_super(r):
    e, n = r["e"] * 100, r["en"] * 100
    Ls = r["Ls_sup"]
    fc = r["jenis"].startswith("Full")
    if fc:
        a, b = r["sta_awal"] - 0.75 * Ls, r["sta_awal"] + 0.25 * Ls
        c, d = r["sta_akhir"] - 0.25 * Ls, r["sta_akhir"] + 0.75 * Ls
        names = ["Awal peralihan", "TC", "CT", "Akhir peralihan"]
    else:
        a, b = r["sta_awal"], r["sta_awal"] + r["Ls"]
        c, d = b + r["Lc"], b + r["Lc"] + r["Ls"]
        names = ["TS", "SC", "CS", "ST"]
    pad = 0.3 * Ls
    xr = (b - a) * 2 * n / (e + n)
    x0 = (b - a) * n / (e + n)
    outer = [(a - pad, -n), (a, -n), (b, e), (c, e), (d, -n), (d + pad, -n)]
    inner = [(a - pad, -n), (a + xr, -n), (b, -e), (c, -e), (d - xr, -n), (d + pad, -n)]

    fig, ax = plt.subplots(figsize=(11, 5))
    ax.plot(*zip(*outer), "r-", lw=2, label="Tepi perkerasan luar")
    ax.plot(*zip(*inner), "b-", lw=2, label="Tepi perkerasan dalam")
    ax.axhline(0, color="k", lw=0.8)
    ax.fill_between(*zip(*outer), 0, alpha=0.05, color="r")
    for nm, xv in zip(names, [a, b, c, d]):
        ax.axvline(xv, color="gray", ls=":", lw=0.9)
        ax.text(xv, ax.get_ylim()[1], f"{nm}\n{sta_str(xv)}", ha="center",
                va="bottom", fontsize=8)
    ax.plot([a + x0, d - x0], [0, 0], "ko", ms=4)
    ax.annotate(f"+e = {e:.2f}%", (b, e), textcoords="offset points", xytext=(5, 6))
    ax.annotate(f"-e = {-e:.2f}%", (b, -e), textcoords="offset points", xytext=(5, -14))
    ax.annotate(f"en = -{n:.1f}%", (a - pad, -n), textcoords="offset points",
                xytext=(0, -14))
    ax.set_ylabel("Kemiringan melintang (%)")
    ax.set_xlabel("Stasioning")
    ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: sta_str(v, 0)))
    ax.set_ylim(-e * 1.35, e * 1.45)
    ax.grid(alpha=0.25)
    ax.legend(loc="lower center", ncol=2, fontsize=8)
    ax.set_title("Diagram Superelevasi (rotasi pada sumbu jalan)", pad=34)
    return fig


# ----------------------------------------------------------------------------
# Profil memanjang & Animasi Mobil
# ----------------------------------------------------------------------------
def hitung_profil(pvi, V):
    d = pvi.dropna().sort_values(pvi.columns[0]).reset_index(drop=True)
    sta = d.iloc[:, 0].to_numpy(float)
    el = d.iloc[:, 1].to_numpy(float)
    lvm = d.iloc[:, 2].to_numpy(float)
    g = np.diff(el) / np.diff(sta)
    Jh = 0.278 * V * 2.5 + V**2 / (254 * 0.40)

    kurva = []
    for i in range(1, len(sta) - 1):
        A = (g[i] - g[i - 1]) * 100
        if abs(A) < 1e-9:
            continue
        a = abs(A)
        if A < 0:   # cembung
            L = a * Jh**2 / 399
            if L < Jh:
                L = 2 * Jh - 399 / a
        else:       # cekung
            L = a * Jh**2 / (120 + 3.5 * Jh)
            if L < Jh:
                L = 2 * Jh - (120 + 3.5 * Jh) / a
        L = math.ceil(max(L, 0.6 * V, V * 3 / 3.6) / 10) * 10
        if lvm[i] > 0:
            L = lvm[i]
        kurva.append(dict(i=i, sta=sta[i], el=el[i], A=A, Lv=L, Ev=a * L / 800,
                          bvc=sta[i] - L / 2, evc=sta[i] + L / 2,
                          jenis="Cembung" if A < 0 else "Cekung",
                          g1=g[i - 1], g2=g[i]))

    x = np.arange(sta[0], sta[-1] + 0.01, 1.0)
    y = np.interp(x, sta, el)
    for k in kurva:
        m = (x >= k["bvc"]) & (x <= k["evc"])
        dx = x[m] - k["bvc"]
        yb = k["el"] - k["g1"] * k["Lv"] / 2
        y[m] = yb + k["g1"] * dx + (k["g2"] - k["g1"]) / (2 * k["Lv"]) * dx**2
    return dict(sta=sta, el=el, g=g * 100, kurva=kurva, x=x, y=y, Jh=Jh)


def draw_car(ax, x_pos, y_pos, slope_angle_deg, car_width=25.0, car_height=5.0):
    """Menggambar patch mobil dengan posisi dan sudut kemiringan sesuai jalan"""
    transform = Affine2D().rotate_deg_around(x_pos, y_pos, slope_angle_deg) + ax.transData
    
    # Body Utama Mobil
    car_body = patches.Rectangle(
        (x_pos - car_width / 2, y_pos + 0.8),
        car_width, car_height,
        facecolor='crimson', edgecolor='black', linewidth=1.2, transform=transform, zorder=10
    )
    # Kabin Atas
    car_roof = patches.Polygon(
        [
            [x_pos - car_width * 0.25, y_pos + 0.8 + car_height],
            [x_pos - car_width * 0.1, y_pos + 0.8 + car_height * 1.7],
            [x_pos + car_width * 0.2, y_pos + 0.8 + car_height * 1.7],
            [x_pos + car_width * 0.35, y_pos + 0.8 + car_height]
        ],
        facecolor='darkred', edgecolor='black', linewidth=1.0, transform=transform, zorder=10
    )
    # Roda Belakang & Depan
    r_radius = car_height * 0.4
    wheel1 = patches.Circle((x_pos - car_width * 0.3, y_pos + r_radius), r_radius, facecolor='black', transform=transform, zorder=11)
    wheel2 = patches.Circle((x_pos + car_width * 0.3, y_pos + r_radius), r_radius, facecolor='black', transform=transform, zorder=11)
    
    ax.add_patch(car_body)
    ax.add_patch(car_roof)
    ax.add_patch(wheel1)
    ax.add_patch(wheel2)


def fig_profil(P, tanah, tik, car_x=None):
    fig, ax = plt.subplots(figsize=(11, 5.5))
    t = tanah.dropna().sort_values(tanah.columns[0])
    ax.plot(t.iloc[:, 0], t.iloc[:, 1], color="saddlebrown", ls="--", marker="o",
            ms=3, label="Tanah asli")
    ax.plot(P["x"], P["y"], "k", lw=2, label="Garis rencana")
    ax.plot(P["sta"], P["el"], "r^", label="PPV")
    
    for k in P["kurva"]:
        ax.axvline(k["bvc"], color="gray", ls=":", lw=0.8)
        ax.axvline(k["evc"], color="gray", ls=":", lw=0.8)
        ax.annotate(f"Lv={k['Lv']:.0f} m\nEv={k['Ev']:.2f} m", (k["sta"], k["el"]),
                    textcoords="offset points",
                    xytext=(0, 12 if k["A"] < 0 else -30), ha="center", fontsize=7)
        
    for i, gi in enumerate(P["g"]):
        xm = (P["sta"][i] + P["sta"][i + 1]) / 2
        ym = (P["el"][i] + P["el"][i + 1]) / 2
        ax.text(xm, ym, f"{gi:+.2f}%", fontsize=8, color="navy", ha="center",
                bbox=dict(fc="white", ec="none", alpha=0.7, pad=1))
        
    if tik:
        ax.axvspan(tik[0], tik[1], color="orange", alpha=0.15, label="Tikungan")

    # Render Posisi Mobil jika ada parameter car_x
    if car_x is not None:
        car_y = float(np.interp(car_x, P["x"], P["y"]))
        # Hitung sudut kemiringan (gradient slope) jalan pada posisi mobil
        dx = 0.5
        y_before = float(np.interp(car_x - dx, P["x"], P["y"]))
        y_after = float(np.interp(car_x + dx, P["x"], P["y"]))
        slope_rad = math.atan2(y_after - y_before, 2 * dx)
        slope_deg = math.degrees(slope_rad)
        
        # Gambar Mobil dengan skala fleksibel terhadap panjang rute
        total_len = P["x"][-1] - P["x"][0]
        c_width = max(15.0, total_len * 0.035)
        c_height = c_width * 0.22
        
        draw_car(ax, car_x, car_y, slope_deg, car_width=c_width, car_height=c_height)
        
        # Label Informasi Mobil Real-time
        ax.annotate(f"🚘 Mobil\nSTA: {sta_str(car_x, 1)}\nElevasi: {car_y:.2f}m\nSlope: {math.tan(slope_rad)*100:+.2f}%", 
                    (car_x, car_y),
                    textcoords="offset points", xytext=(0, 25), ha="center", fontsize=8,
                    bbox=dict(boxstyle="round,pad=0.3", fc="yellow", alpha=0.8, ec="black"))

    ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: sta_str(v, 0)))
    ax.set_xlabel("Stasioning")
    ax.set_ylabel("Elevasi (m)")
    ax.grid(alpha=0.25)
    ax.legend(fontsize=8, loc="best")
    ax.set_title("Profil Memanjang & Simulasi Pergerakan Mobil")
    return fig


def png(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight")
    return buf.getvalue()


# ----------------------------------------------------------------------------
# Antarmuka Utama Streamlit
# ----------------------------------------------------------------------------
st.title("Perencanaan Geometrik Jalan")
st.caption("Tikungan horizontal, superelevasi, dan profil memanjang - acuan TPGJAK Bina Marga 1997")

with st.sidebar:
    st.header("Data input")
    V = st.number_input("Kecepatan rencana Vr (km/jam)", 20, 120, 60, 10)
    delta = st.number_input("Sudut tikungan Δ (°)", 1.0, 179.0, 40.0, 0.5)
    emax = st.number_input("e maks (%)", 4.0, 12.0, 10.0, 0.5) / 100
    en = st.number_input("Kemiringan normal en (%)", 1.0, 4.0, 2.0, 0.5) / 100
    Rmin = r_min(V, emax)
    st.caption(f"R minimum = {Rmin:.1f} m")
    mode = st.radio("Jari-jari tikungan Rc", ["Otomatis (Rmin dibulatkan)", "Manual"])
    if mode.startswith("Oto"):
        Rc = float(math.ceil(Rmin / 10) * 10)
        st.caption(f"Rc = {Rc:.0f} m")
    else:
        Rc = st.number_input("Rc (m)", 10.0, 5000.0, 150.0, 5.0)
    sta_pi = st.number_input("STA PI (m)", 0.0, 100000.0, 250.0, 10.0)
    arah = st.radio("Arah tikungan", ["Kanan", "Kiri"], horizontal=True)
    B = st.number_input("Lebar badan jalan (m)", 3.0, 30.0, 7.0, 0.5)

r = hitung_tikungan(V, delta, emax, en, Rc, sta_pi)
if Rc < Rmin:
    st.error(f"Rc = {Rc:.1f} m lebih kecil dari R minimum ({Rmin:.1f} m). Perbesar Rc.")
if r["jenis"].endswith("(S-S)") and r["Ls"] < r["Ls_min"]:
    st.warning("Ls pada tikungan S-S lebih pendek dari Ls minimum. Perbesar Rc.")

P_global = None
df_earthwork_global = None

tab1, tab2, tab3, tab4 = st.tabs(["Denah tikungan", "Diagram superelevasi",
                                  "Profil memanjang & Animasi", "Hasil hitungan"])

with tab1:
    f1 = fig_denah(r, arah, B)
    st.pyplot(f1)
    st.download_button("Unduh gambar (PNG)", png(f1), "denah_tikungan.png", "image/png")

with tab2:
    f2 = fig_super(r)
    st.pyplot(f2)
    st.download_button("Unduh gambar (PNG)", png(f2), "diagram_superelevasi.png", "image/png")

with tab3:
    st.write("Isi tabel PPV (titik perpotongan garis kelandaian) dan tanah asli. "
             "Kolom Lv = 0 berarti panjang lengkung vertikal dihitung otomatis.")
    c1, c2 = st.columns(2)
    with c1:
        pvi = st.data_editor(
            pd.DataFrame({"STA (m)": [0.0, 200.0, 400.0, 600.0],
                          "Elevasi rencana (m)": [100.0, 104.0, 102.0, 106.0],
                          "Lv manual (m)": [0.0, 0.0, 0.0, 0.0]}),
            num_rows="dynamic", key="pvi", hide_index=True)
    with c2:
        tanah = st.data_editor(
            pd.DataFrame({"STA (m)": [0.0, 100.0, 200.0, 300.0, 400.0, 500.0, 600.0],
                          "Elevasi tanah asli (m)": [100.5, 102.3, 103.0, 103.4, 102.6, 103.9, 105.2]}),
            num_rows="dynamic", key="tanah", hide_index=True)

    pvi = pvi.fillna(0.0)
    if len(pvi.dropna()) >= 2 and pvi.iloc[:, 0].is_unique:
        P = hitung_profil(pvi, V)
        P_global = P
        gmax = float(np.interp(V, V_TAB, G_MAX_TAB))
        for i, gi in enumerate(P["g"]):
            if abs(gi) > gmax:
                st.warning(f"Kelandaian segmen {i + 1} = {gi:.2f}% melebihi batas {gmax:.0f}% untuk Vr {V} km/jam.")
        for k in P["kurva"]:
            lm = min(P["sta"][k["i"]] - P["sta"][k["i"] - 1], P["sta"][k["i"] + 1] - P["sta"][k["i"]])
            if k["Lv"] / 2 > lm:
                st.warning(f"Lengkung vertikal di STA {sta_str(k['sta'], 0)} saling bertumpuk dengan PPV tetangga.")

        # --------------------------------------------------------------------
        # Bagian Animasi Mobil Interaktif
        # --------------------------------------------------------------------
        st.subheader("🚘 Animasi & Visualisasi Mobil Berjalan")
        
        col_ctrl1, col_ctrl2 = st.columns([1, 3])
        with col_ctrl1:
            run_animation = st.button("▶️ Jalankan Animasi Mobil", use_container_width=True)
            kecepatan_animasi = st.slider("Kecepatan Simulasi", 1, 10, 5)
        
        with col_ctrl2:
            car_pos_slider = st.slider(
                "Atur Posisi Mobil Manual (STA)", 
                float(P["x"][0]), float(P["x"][-1]), float(P["x"][0]), step=5.0
            )

        plot_container = st.empty()

        if run_animation:
            # Loop pergerakan frame mobil
            step_size = (P["x"][-1] - P["x"][0]) / (50 * (11 - kecepatan_animasi))
            car_x_curr = P["x"][0]
            
            while car_x_curr <= P["x"][-1]:
                fig_anim = fig_profil(P, tanah, (r["sta_awal"], r["sta_akhir"]), car_x=car_x_curr)
                plot_container.pyplot(fig_anim)
                plt.close(fig_anim)
                car_x_curr += max(2.0, step_size)
                time.sleep(0.05)
        else:
            fig_static = fig_profil(P, tanah, (r["sta_awal"], r["sta_akhir"]), car_x=car_pos_slider)
            plot_container.pyplot(fig_static)
            st.download_button("Unduh gambar (PNG)", png(fig_static), "profil_memanjang.png", "image/png")

        st.caption(f"Jarak pandang henti Jh = {P['Jh']:.1f} m")
        if P["kurva"]:
            st.subheader("Lengkung vertikal")
            st.dataframe(pd.DataFrame([{
                "STA PPV": sta_str(k["sta"], 0), "Jenis": k["jenis"],
                "g1 (%)": round(k["g1"] * 100, 2), "g2 (%)": round(k["g2"] * 100, 2),
                "A (%)": round(abs(k["A"]), 2), "Lv (m)": k["Lv"], "Ev (m)": round(k["Ev"], 3),
                "STA PLV": sta_str(k["bvc"], 0), "STA PTV": sta_str(k["evc"], 0)}
                for k in P["kurva"]]), hide_index=True)

        t = tanah.dropna().sort_values(tanah.columns[0])
        xs = np.arange(math.ceil(P["x"][0] / 50) * 50, P["x"][-1] + 0.1, 50)
        rn = np.interp(xs, P["x"], P["y"])
        ta = np.interp(xs, t.iloc[:, 0], t.iloc[:, 1])
        st.subheader("Galian / timbunan tiap 50 m")
        df_earthwork_global = pd.DataFrame({
            "STA_val": xs,
            "STA": [sta_str(v, 0) for v in xs], "Elevasi rencana (m)": rn.round(3),
            "Elevasi tanah (m)": ta.round(3), "Selisih (m)": (rn - ta).round(3),
            "Keterangan": np.where(rn - ta > 0, "Timbunan", "Galian")})
        st.dataframe(df_earthwork_global.drop(columns=["STA_val"]), hide_index=True)
    else:
        st.info("Isi minimal 2 PPV dengan STA yang berbeda.")

with tab4:
    st.subheader("Tabel Perhitungan Geometrik Jalan")
    st.dataframe(tabel_hasil(r), hide_index=True, height=600)
    
    st.markdown("---")
    st.subheader("📊 Unduh Laporan Perhitungan Excel")
    st.caption("File Excel ini dilengkapi **formula aktif** (seperti `MAX()`, `CEILING()`, `IF()`, `RADIANS()`, dll.) sehingga bisa di-edit dan otomatis ter-recalculating.")
    
    excel_bytes = generate_excel_with_formulas(r, P_global, df_earthwork_global)
    st.download_button(
        label="📥 Unduh Perhitungan Lengkap (.xlsx)",
        data=excel_bytes,
        file_name="Perencanaan_Geometrik_Jalan.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )