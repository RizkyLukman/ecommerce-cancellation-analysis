"""
Bangun dashboard Excel otomatis (versi pembanding untuk dashboard manual).

Input : Clean_data/analysis_dataset.xlsx  (output 02_analysis.ipynb)
Output: Dashboard/dashboard_cancel_return_auto.xlsx

Alur: Power Query -> tbl_orders -> 1 pivot cache -> 6 pivot (sheet `pivot`)
      -> tabel bantu GETPIVOTDATA (sheet `chart_data`) -> chart + slicer (sheet `dashboard`).
Butuh Excel 2016+ (Windows) dan pywin32.
"""
import os
import sys
import win32com.client as win32

PROJ = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(PROJ, "Clean_data", "analysis_dataset.xlsx")
OUT_DIR = os.path.join(PROJ, "Dashboard")
# argumen opsional: path output lain (mis. untuk pratinjau saat file utama sedang dibuka)
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(OUT_DIR, "dashboard_cancel_return_auto.xlsx")

# ---- konstanta Excel ----
xlDatabase, xlSrcExternal, xlCmdSql = 1, 0, 2
xlRowField, xlColumnField, xlPageField, xlDataField = 1, 2, 3, 4
xlSum, xlCount = -4157, -4112
xlDescending, xlTopCount = 2, 1
xlTabular = 1
xlColumnClustered, xlBarClustered, xlLine, xlDoughnut = 51, 57, 4, -4120
xlCategory, xlValue = 1, 2
xlLegendTop, xlLegendRight = -4160, -4152
msoShapeRectangle, msoShapeRoundedRectangle, msoTextOrientationHorizontal = 1, 5, 1
xlOpenXMLWorkbook = 51


def rgb(hex_):
    h = hex_.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return r + g * 256 + b * 65536


NAVY, NAVY2 = rgb("#1F2430"), rgb("#2C3240")
PINK, PINK_DARK, TEAL = rgb("#E8426B"), rgb("#B5123A"), rgb("#2BB3A0")
# latar abu kebiruan (senada navy) agar kartu putih & teks lebih kontras
GREY_BG, GREY_TXT, GREY_LINE = rgb("#D5DCE6"), rgb("#4B5563"), rgb("#9CA3AF")
PINK_SOFT, PINK_LINE = rgb("#FDECEF"), rgb("#F4B6C4")
WHITE, INK = rgb("#FFFFFF"), rgb("#1F2430")
FONT = "Segoe UI"

M_QUERY = f'''let
    Source = Excel.Workbook(File.Contents("{SRC}"), null, true),
    data = Source{{[Item="data",Kind="Sheet"]}}[Data],
    Promoted = Table.PromoteHeaders(data, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted, {{
        {{"order_id", type text}}, {{"bulan_file", type text}}, {{"tahun", Int64.Type}},
        {{"bulan", type date}}, {{"status_efektif", type text}}, {{"is_settled", Int64.Type}},
        {{"is_gagal_revenue", Int64.Type}}, {{"is_cancelled", Int64.Type}},
        {{"cancel_source", type text}}, {{"alasan_standar", type text}},
        {{"is_returned", Int64.Type}}, {{"return_type", type text}},
        {{"metode_pembayaran", type text}}, {{"metode_analisis", type text}},
        {{"jumlah_kategori_berbeda", Int64.Type}}, {{"daftar_kategori", type text}},
        {{"kategori_analisis", type text}}, {{"is_artefak_kirim", Int64.Type}}}})
in
    Typed'''


# ------------------------------------------------------------------ pivot helpers
def make_pivot(cache, ws, dest, name):
    pt = cache.CreatePivotTable(TableDestination=ws.Range(dest), TableName=name)
    pt.RowAxisLayout(xlTabular)
    pt.ShowDrillIndicators = False
    return pt


def add_rate(pt, settled=True, gagal=True):
    if settled:
        pt.AddDataField(pt.PivotFields("is_settled"), "Order Settled", xlSum)
    if gagal:
        pt.AddDataField(pt.PivotFields("is_gagal_revenue"), "Order Gagal", xlSum)
    f = pt.AddDataField(pt.PivotFields("rate_gagal"), "Gagal Rate", xlSum)
    f.NumberFormat = "0.00%"


def add_row(pt, field):
    pf = pt.PivotFields(field)
    pf.Orientation = xlRowField
    return pf


def hide_blank(pf):
    for it in pf.PivotItems():
        if it.Name in ("(blank)", ""):
            it.Visible = False


def page_filter(pt, field, value):
    pf = pt.PivotFields(field)
    pf.Orientation = xlPageField
    pf.ClearAllFilters()
    pf.CurrentPage = str(value)


# ------------------------------------------------------------------ shape helpers
def box(ws, left, top, w, h, fill, rounded=False, line=None):
    s = ws.Shapes.AddShape(msoShapeRoundedRectangle if rounded else msoShapeRectangle, left, top, w, h)
    s.Fill.ForeColor.RGB = fill
    if line is None:
        s.Line.Visible = False
    else:
        s.Line.ForeColor.RGB = line
        s.Line.Weight = 0.75
    if rounded:
        s.Adjustments.SetItem(1, 0.06)
    s.Shadow.Visible = False
    return s


def text(ws, left, top, w, h, txt, size=10, color=INK, bold=False, italic=False, align=1, link=None):
    t = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, left, top, w, h)
    t.Fill.Visible = False
    t.Line.Visible = False
    tf = t.TextFrame2
    tf.MarginLeft = tf.MarginRight = tf.MarginTop = tf.MarginBottom = 0
    tf.WordWrap = True
    if link:
        t.DrawingObject.Formula = link
    else:
        tf.TextRange.Text = txt
    r = tf.TextRange
    r.Font.Name = FONT
    r.Font.Size = size
    r.Font.Bold = bold
    r.Font.Italic = italic
    r.Font.Fill.ForeColor.RGB = color
    r.ParagraphFormat.Alignment = align  # 1 kiri, 2 tengah
    return t


def style_chart(ch, title, subtitle=None):
    # font dasar dulu; kalau diset setelah judul, ukuran judul ikut tertimpa
    ch.ChartArea.Format.Line.Visible = False
    ch.ChartArea.Format.Fill.ForeColor.RGB = WHITE
    ch.ChartArea.Font.Name = FONT
    ch.ChartArea.Font.Size = 9
    ch.ChartArea.Font.Color = INK
    ch.HasTitle = True
    ch.ChartTitle.Text = title if not subtitle else f"{title}\n{subtitle}"
    t = ch.ChartTitle
    # early-binding (makepy) mengekspos property berparameter sebagai GetCharacters(...)
    chars = t.GetCharacters if hasattr(t, "GetCharacters") else t.Characters
    f = chars().Font
    f.Size, f.Bold, f.Color = 12, True, INK
    if subtitle:
        f = chars(len(title) + 2, len(subtitle)).Font
        f.Size, f.Bold, f.Color = 9, False, GREY_TXT
    t.HorizontalAlignment = -4131  # xlLeft: judul & subjudul rata kiri di semua chart
    t.Left = 8


def tidy_value_axis(ch, show=False, gridlines=False):
    ax = ch.Axes(xlValue)
    ax.HasMajorGridlines = gridlines
    if gridlines:
        ax.MajorGridlines.Format.Line.ForeColor.RGB = rgb("#E5E7EB")
    ax.TickLabelPosition = 4 if show else -4142  # xlTickLabelPositionNextToAxis / None
    ax.Format.Line.Visible = False
    ax.MinimumScale = 0


# ================================================================== main
def build():
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    # simpan ke file sementara dulu; file lama baru diganti setelah build sukses
    tmp = OUT[:-5] + "_tmp.xlsx"
    if os.path.exists(tmp):
        os.remove(tmp)

    # DispatchEx = instans Excel baru & terpisah, supaya Quit() tidak menutup Excel yang sedang dipakai user
    xl = win32.gencache.EnsureDispatch(win32.DispatchEx("Excel.Application"))
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Add()
        while wb.Worksheets.Count < 5:
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws_dash, ws_readme, ws_cd, ws_pv, ws_data = (wb.Worksheets(i) for i in range(1, 6))
        for ws, nm in [(ws_dash, "dashboard"), (ws_readme, "README"), (ws_cd, "chart_data"),
                       (ws_pv, "pivot"), (ws_data, "data")]:
            ws.Name = nm

        # ---------------------------------------------------------- 1. data (Power Query)
        wb.Queries.Add("q_orders", M_QUERY)
        lo = ws_data.ListObjects.Add(
            SourceType=xlSrcExternal,
            Source='OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;'
                   'Location=q_orders;Extended Properties=""',
            Destination=ws_data.Range("$A$1"))
        qt = lo.QueryTable
        qt.CommandType = xlCmdSql
        qt.CommandText = "SELECT * FROM [q_orders]"
        qt.BackgroundQuery = False
        qt.RefreshStyle = 1
        qt.Refresh(False)
        lo.Name = "tbl_orders"
        lo.TableStyle = "TableStyleLight1"
        # bulan = tanggal dummy (thn 2000) -> tampil nama bulan Indonesia, slicer urut kronologis
        lo.ListColumns("bulan").DataBodyRange.NumberFormat = "[$-421]mmm"
        n_rows = lo.ListRows.Count
        print("tbl_orders rows:", n_rows)
        assert n_rows == 20404

        # ---------------------------------------------------------- 2. pivots (1 cache)
        cache = wb.PivotCaches().Create(SourceType=xlDatabase, SourceData="tbl_orders")
        ws_pv.Range("B1").Value = "Semua pivot memakai 1 pivot cache dari tbl_orders " \
                                  "(syarat agar slicer bisa mengendalikan beberapa pivot)."

        def label(cell, txt):
            ws_pv.Range(cell).Value = txt
            ws_pv.Range(cell).Font.Bold = True

        label("B3", "P1 KPI")
        p1 = make_pivot(cache, ws_pv, "B6", "P1_KPI")
        p1.CalculatedFields().Add("rate_gagal", "=is_gagal_revenue/is_settled", True)
        add_rate(p1)

        label("G3", "P2 Tren (tidak ikut slicer Tahun & Bulan)")
        p2 = make_pivot(cache, ws_pv, "G6", "P2_Tren")
        add_row(p2, "bulan_file")
        add_rate(p2)

        label("M3", "P3 Sumber pembatalan")
        p3 = make_pivot(cache, ws_pv, "M6", "P3_Sumber")
        page_filter(p3, "is_cancelled", 1)
        add_row(p3, "cancel_source")
        p3.AddDataField(p3.PivotFields("order_id"), "Order Batal", xlCount)

        label("Q3", "P4 Metode (tidak ikut slicer Bulan & Metode)")
        p4 = make_pivot(cache, ws_pv, "Q6", "P4_Metode")
        pf4 = add_row(p4, "metode_analisis")
        add_rate(p4)
        hide_blank(pf4)
        pf4.AutoSort(xlDescending, "Gagal Rate")
        p4.ColumnGrand = False

        label("W3", "P5 Kategori top 5 (tidak ikut slicer Bulan)")
        p5 = make_pivot(cache, ws_pv, "W6", "P5_Kategori")
        pf5 = add_row(p5, "kategori_analisis")
        add_rate(p5)
        hide_blank(pf5)
        pf5.PivotFilters.Add2(Type=xlTopCount, DataField=p5.PivotFields("Gagal Rate"), Value1=5)
        pf5.AutoSort(xlDescending, "Gagal Rate")
        p5.ColumnGrand = False

        label("AC3", "P6 Alasan top 5")
        p6 = make_pivot(cache, ws_pv, "AC6", "P6_Alasan")
        page_filter(p6, "is_cancelled", 1)
        pf6 = add_row(p6, "alasan_standar")
        p6.AddDataField(p6.PivotFields("order_id"), "Order Batal", xlCount)
        pf6.PivotFilters.Add2(Type=xlTopCount, DataField=p6.PivotFields("Order Batal"), Value1=5)
        pf6.AutoSort(xlDescending, "Order Batal")
        p6.ColumnGrand = False

        for c in ("B", "G", "M", "Q", "W", "AC"):
            ws_pv.Columns(c).ColumnWidth = 26
        anchor = {p.Name: f"pivot!{p.TableRange1.Cells(1, 1).Address}" for p in (p1, p2, p3, p4, p5, p6)}

        # ---------------------------------------------------------- 3. chart_data
        cd = ws_cd
        cd.Range("A1").Value = "Tabel bantu chart & KPI. Semua angka diambil dari pivot (GETPIVOTDATA/INDEX); " \
                               "jangan diketik manual."
        a1, a2, a3 = anchor["P1_KPI"], anchor["P2_Tren"], anchor["P3_Sumber"]

        # KPI (baris 3-6)
        cd.Range("A3:B3").Value = ("KPI", "Nilai")
        kpi = [("Total Order Settled", f'=GETPIVOTDATA("Order Settled",{a1})', "#,##0"),
               ("Gagal-Revenue Rate", f'=GETPIVOTDATA("Gagal Rate",{a1})', "0.00%"),
               ("Total Order Gagal", f'=GETPIVOTDATA("Order Gagal",{a1})', "#,##0")]
        for i, (nm, f, fmt) in enumerate(kpi):
            r = 4 + i
            cd.Range(f"A{r}").Value = nm
            cd.Range(f"B{r}").Formula = f
            cd.Range(f"B{r}").NumberFormat = fmt
            cd.Range(f"C{r}").Formula = f'=TEXT(B{r},"{fmt}")'

        # Tren (baris 9-31)
        cd.Range("A9:E9").Value = ("periode", "rate", "rate_bulan_lain", "rate_mei_2025", "rata_rata")
        bulan = [f"{y}-{m:02d}" for y in (2024, 2025) for m in range(1, 13)][:23]
        for i, b in enumerate(bulan):
            r = 10 + i
            cd.Range(f"A{r}").NumberFormat = "@"
            cd.Range(f"A{r}").Value = b
            cd.Range(f"B{r}").Formula = f'=IFERROR(GETPIVOTDATA("Gagal Rate",{a2},"bulan_file",A{r}),NA())'
            cd.Range(f"C{r}").Formula = f'=IF(A{r}="2025-05",NA(),B{r})'
            cd.Range(f"D{r}").Formula = f'=IF(A{r}="2025-05",B{r},NA())'
            cd.Range(f"E{r}").Formula = f'=GETPIVOTDATA("Gagal Rate",{a2})'
        cd.Range("B10:E32").NumberFormat = "0.00%"
        TR = (10, 32)

        # Sumber (baris 35-38)
        cd.Range("A35:B35").Value = ("cancel_source", "order_batal")
        for i, src in enumerate(["Pembeli", "Sistem/Penjual (Otomatis)", "Penjual (Manual)"]):
            r = 36 + i
            cd.Range(f"A{r}").Value = src
            cd.Range(f"B{r}").Formula = f'=IFERROR(GETPIVOTDATA("Order Batal",{a3},"cancel_source",A{r}),0)'
        cd.Range("A39").Value = "Total order batal"
        cd.Range("B39").Formula = "=SUM(B36:B38)"
        cd.Range("C39").Formula = '="Total: "&TEXT(B39,"#,##0")&" order batal"'

        # Metode (baris 42-50): INDEX ke kolom pivot P4, urutan sesuai sort pivot
        cd.Range("A42:F42").Value = ("label", "metode", "n_settled", "rate", "di_atas_rata2", "di_bawah_rata2")
        q = p4.TableRange1
        r0, c0 = q.Row + 1, q.Column
        col = lambda off: ws_pv.Cells(r0, c0 + off).Address.replace("$", "")
        for i in range(8):
            r = 43 + i
            rows = f"pivot!${col(0)[0]}${r0}:${col(0)[0]}${r0 + 9}"
            cd.Range(f"B{r}").Formula = f'=IFERROR(INDEX({rows},{i + 1})&"","")'
            cd.Range(f"C{r}").Formula = (f'=IF(B{r}="","",INDEX(pivot!${col(1)[0]}${r0}:${col(1)[0]}${r0 + 9},{i + 1}))')
            cd.Range(f"D{r}").Formula = (f'=IF(B{r}="",NA(),INDEX(pivot!${col(3)[0]}${r0}:${col(3)[0]}${r0 + 9},{i + 1}))')
            cd.Range(f"A{r}").Formula = f"=B{r}"
            cd.Range(f"E{r}").Formula = f'=IF(ISNA(D{r}),NA(),IF(D{r}>$B$52,D{r},NA()))'
            cd.Range(f"F{r}").Formula = f'=IF(ISNA(D{r}),NA(),IF(D{r}<=$B$52,D{r},NA()))'
        cd.Range("A52").Value = "Rata-rata (metode n>=30)"
        g_col, s_col = col(2)[0], col(1)[0]
        cd.Range("B52").Formula = (f"=SUM(pivot!${g_col}${r0}:${g_col}${r0 + 9})/"
                                   f"SUM(pivot!${s_col}${r0}:${s_col}${r0 + 9})")
        cd.Range("D43:F50,B52").NumberFormat = "0.00%"

        # Kategori top 5 (baris 55-60)
        cd.Range("A55:D55").Value = ("label", "kategori", "n_settled", "rate")
        q = p5.TableRange1
        r0, c0 = q.Row + 1, q.Column
        L = lambda off: ws_pv.Cells(r0, c0 + off).Address.split("$")[1]
        for i in range(5):
            r = 56 + i
            cd.Range(f"B{r}").Formula = f'=IFERROR(INDEX(pivot!${L(0)}${r0}:${L(0)}${r0 + 6},{i + 1})&"","")'
            cd.Range(f"C{r}").Formula = f'=IF(B{r}="","",INDEX(pivot!${L(1)}${r0}:${L(1)}${r0 + 6},{i + 1}))'
            cd.Range(f"D{r}").Formula = f'=IF(B{r}="",NA(),INDEX(pivot!${L(3)}${r0}:${L(3)}${r0 + 6},{i + 1}))'
            cd.Range(f"A{r}").Formula = f"=B{r}"
        cd.Range("D56:D60").NumberFormat = "0.00%"

        # Alasan top 5 (baris 63-68): % terhadap TOTAL order batal (dari P3),
        # bukan terhadap total top-5 (grand total pivot ikut terpotong filter Top 5)
        cd.Range("A63:D63").Value = ("alasan", "order_batal", "pct_dari_total_batal", "")
        q = p6.TableRange1
        r0, c0 = q.Row + 1, q.Column
        for i in range(5):
            r = 64 + i
            cd.Range(f"A{r}").Formula = f'=IFERROR(INDEX(pivot!${L(0)}${r0}:${L(0)}${r0 + 6},{i + 1})&"","")'
            cd.Range(f"B{r}").Formula = f'=IF(A{r}="","",INDEX(pivot!${L(1)}${r0}:${L(1)}${r0 + 6},{i + 1}))'
            cd.Range(f"C{r}").Formula = f'=IF(A{r}="",NA(),B{r}/$B$39)'
        cd.Range("C64:C68").NumberFormat = "0.00%"
        cd.Range("A69").Value = "Top 5 kumulatif"
        cd.Range("C69").Formula = "=SUM(C64:C68)"
        cd.Range("C69").NumberFormat = "0.00%"

        cd.Columns("A").ColumnWidth = 34
        cd.Columns("B:F").ColumnWidth = 16
        for hdr in ("A3:B3", "A9:E9", "A35:B35", "A42:F42", "A55:D55", "A63:C63"):
            cd.Range(hdr).Font.Bold = True

        # ---------------------------------------------------------- 4. dashboard
        d = ws_dash
        d.Activate()
        xl.ActiveWindow.DisplayGridlines = False
        xl.ActiveWindow.DisplayHeadings = False
        d.Cells.Interior.Color = GREY_BG
        d.Cells.Font.Name = FONT
        W, X0 = 1180, 10
        HALF = (W - 10) / 2

        # header
        box(d, X0, 10, W, 64, NAVY)
        text(d, X0 + 20, 20, 700, 28, "Dashboard Analisis Pembatalan & Retur", 20, WHITE, bold=True)
        text(d, X0 + 20, 50, 900, 16,
             "Periode Januari 2024 – November 2025 (23 bulan). Gagal-revenue = order batal + retur penuh, dibagi order settled.",
             9.5, rgb("#C9CDD6"))

        # slicer strip
        box(d, X0, 82, W, 96, WHITE, rounded=True)
        slicer_specs = [
            ("tahun", "Tahun", X0 + 10, 160, 1, [p1, p3, p4, p5, p6]),
            ("bulan", "Bulan", X0 + 180, 430, 6, [p1, p3, p6]),
            ("metode_pembayaran", "Metode Bayar", X0 + 620, 550, 4, [p1, p2, p3, p5, p6]),
        ]
        for field, cap, left, width, ncol, targets in slicer_specs:
            sc = wb.SlicerCaches.Add2(targets[0], field, f"Slicer_{field}")
            for pt in targets[1:]:
                sc.PivotTables.AddPivotTable(pt)
            if field == "bulan":
                # item = tanggal, jadi urut ascending = Jan..Des; custom list tidak dipakai
                sc.SortUsingCustomLists = False
                sc.SortItems = 2  # xlSlicerSortAscending
            sl = sc.Slicers.Add(SlicerDestination=d, Name=f"sl_{field}", Caption=cap,
                                Top=88, Left=left, Width=width, Height=84)
            sl.NumberOfColumns = ncol
            sl.RowHeight = 16 if field != "tahun" else 22
            sl.Style = "SlicerStyleOther2"
            sl.DisableMoveResizeUI = True

        # KPI cards
        cards = [("Total Order Settled", "C4", PINK), ("Gagal-Revenue Rate", "C5", INK),
                 ("Total Order Gagal", "C6", TEAL)]
        cw = (W - 30) / 4
        for i, (title, cell, colr) in enumerate(cards):
            left = X0 + i * (cw + 10)
            box(d, left, 186, cw, 72, WHITE, rounded=True)
            box(d, left + 6, 186, cw - 12, 4, colr)  # aksen warna di atas kartu
            text(d, left, 196, cw, 16, title, 10.5, GREY_TXT, bold=True, align=2)
            text(d, left, 214, cw, 36, "", 24, colr, bold=True, align=2, link=f"=chart_data!{cell}")
        left = X0 + 3 * (cw + 10)
        box(d, left, 186, cw, 72, PINK_SOFT, rounded=True, line=PINK_LINE)
        text(d, left + 12, 200, cw - 24, 48,
             "Nilai kerugian (Rp) tidak ditampilkan karena nilai pembayaran pada order retur tidak konsisten. "
             "Seluruh angka berbasis jumlah order.", 9, PINK_DARK, italic=True, align=2)

        def add_chart(left, top, w, h, ctype):
            co = d.ChartObjects().Add(left, top, w, h)
            ch = co.Chart
            ch.ChartType = ctype
            while ch.SeriesCollection().Count:
                ch.SeriesCollection(1).Delete()
            return co, ch

        # tren
        co, ch = add_chart(X0, 266, W, 240, xlColumnClustered)
        x_rng = f"=chart_data!$A${TR[0]}:$A${TR[1]}"
        for nm, colref, colr in [("Rate bulanan", "C", PINK), ("Mei 2025 (tertinggi)", "D", PINK_DARK)]:
            s = ch.SeriesCollection().NewSeries()
            s.Name = nm
            s.Values = f"=chart_data!${colref}${TR[0]}:${colref}${TR[1]}"
            s.XValues = x_rng
            s.Format.Fill.ForeColor.RGB = colr
        ch.ChartGroups(1).Overlap = 100
        ch.ChartGroups(1).GapWidth = 45
        s = ch.SeriesCollection().NewSeries()
        s.Name = "Rata-rata periode"
        s.Values = f"=chart_data!$E${TR[0]}:$E${TR[1]}"
        s.XValues = x_rng
        s.ChartType = xlLine
        s.Format.Line.ForeColor.RGB = GREY_LINE
        s.Format.Line.DashStyle = 4
        s.Format.Line.Weight = 1.25
        ch.SeriesCollection(2).HasDataLabels = True
        ch.SeriesCollection(2).DataLabels().NumberFormat = "0.0%"
        ch.SeriesCollection(2).DataLabels().Format.TextFrame2.TextRange.Font.Bold = True
        style_chart(ch, "Tren Gagal-Revenue Rate Bulanan",
                    "Persentase order settled yang batal atau retur penuh. Hanya dipengaruhi filter Metode Bayar.")
        ch.HasLegend = True
        ch.Legend.Position = xlLegendTop
        tidy_value_axis(ch, show=True, gridlines=True)
        ch.Axes(xlValue).TickLabels.NumberFormat = "0%"
        ch.Axes(xlCategory).TickLabels.Orientation = 45
        ch.Axes(xlCategory).Format.Line.ForeColor.RGB = rgb("#D1D5DB")

        # donut sumber
        co, ch = add_chart(X0, 514, HALF, 210, xlDoughnut)
        s = ch.SeriesCollection().NewSeries()
        s.Name = "Sumber"
        s.Values = "=chart_data!$B$36:$B$38"
        s.XValues = "=chart_data!$A$36:$A$38"
        for i, colr in enumerate([PINK, TEAL, GREY_LINE], start=1):
            s.Points(i).Format.Fill.ForeColor.RGB = colr
        ch.ChartGroups(1).DoughnutHoleSize = 58
        s.HasDataLabels = True
        dl = s.DataLabels()
        dl.ShowValue = False
        dl.ShowPercentage = True
        dl.NumberFormat = "0.0%"
        dl.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = WHITE
        dl.Format.TextFrame2.TextRange.Font.Bold = True
        style_chart(ch, "Sumber Pembatalan", "Komposisi dari total order batal (tidak termasuk retur penuh).")
        ch.HasLegend = True
        ch.Legend.Position = xlLegendRight
        ch.Legend.Format.TextFrame2.TextRange.Font.Size = 10
        text(d, X0 + HALF - 250, 700, 240, 16, "", 8.5, GREY_TXT, italic=True, align=2,
             link="=chart_data!C39")

        # metode
        co, ch = add_chart(X0 + HALF + 10, 514, HALF, 210, xlBarClustered)
        for nm, colref, colr in [("Di atas rata-rata", "E", PINK), ("Di bawah rata-rata", "F", TEAL)]:
            s = ch.SeriesCollection().NewSeries()
            s.Name = nm
            s.Values = f"=chart_data!${colref}$43:${colref}$50"
            s.XValues = "=chart_data!$A$43:$A$50"
            s.Format.Fill.ForeColor.RGB = colr
            s.HasDataLabels = True
            s.DataLabels().NumberFormat = "0.0%"
        ch.ChartGroups(1).Overlap = 100
        ch.ChartGroups(1).GapWidth = 40
        style_chart(ch, "Gagal-Revenue Rate per Metode Pembayaran",
                    "Metode dengan minimal 30 order. Tidak dipengaruhi filter Bulan dan Metode Bayar.")
        ch.HasLegend = True
        ch.Legend.Position = xlLegendTop
        ch.Axes(xlCategory).ReversePlotOrder = True
        ch.Axes(xlCategory).Format.Line.ForeColor.RGB = rgb("#D1D5DB")
        tidy_value_axis(ch)

        # kategori
        co, ch = add_chart(X0, 732, HALF, 200, xlBarClustered)
        s = ch.SeriesCollection().NewSeries()
        s.Name = "Gagal-Revenue Rate"
        s.Values = "=chart_data!$D$56:$D$60"
        s.XValues = "=chart_data!$A$56:$A$60"
        s.Format.Fill.ForeColor.RGB = TEAL
        s.HasDataLabels = True
        s.DataLabels().NumberFormat = "0.0%"
        ch.ChartGroups(1).GapWidth = 45
        style_chart(ch, "Top 5 Kategori Rawan Gagal",
                    "Order dengan satu kategori produk, minimal 30 order. Tidak dipengaruhi filter Bulan.")
        ch.HasLegend = False
        ch.Axes(xlCategory).ReversePlotOrder = True
        ch.Axes(xlCategory).Format.Line.ForeColor.RGB = rgb("#D1D5DB")
        tidy_value_axis(ch)

        # alasan
        co, ch = add_chart(X0 + HALF + 10, 732, HALF, 200, xlBarClustered)
        s = ch.SeriesCollection().NewSeries()
        s.Name = "% dari total order batal"
        s.Values = "=chart_data!$C$64:$C$68"
        s.XValues = "=chart_data!$A$64:$A$68"
        s.Format.Fill.ForeColor.RGB = PINK
        s.HasDataLabels = True
        s.DataLabels().NumberFormat = "0.0%"
        ch.ChartGroups(1).GapWidth = 45
        style_chart(ch, "Top 5 Alasan Pembatalan", "Persentase dari total order batal. Label alasan dengan arti yang sama sudah digabung.")
        ch.HasLegend = False
        ch.Axes(xlCategory).ReversePlotOrder = True
        ch.Axes(xlCategory).Format.Line.ForeColor.RGB = rgb("#D1D5DB")
        tidy_value_axis(ch)

        # panel insight & keputusan (teks statis: berlaku untuk seluruh periode, tidak ikut slicer).
        # Angka memakai format yang sama dengan angka lain di dashboard (locale Excel: 14.07%, 1,056).
        INSIGHTS = [
            ("38.6% pembatalan (1,056 order) terjadi karena pembeli ingin mengubah pesanan, alamat, atau voucher.",
             "Kirim konfirmasi alamat dan varian setelah order masuk, lalu pantau perubahan jumlah pembatalan 'ubah pesanan'."),
            ("COD menyumbang 54.4% dari total order gagal meskipun rate-nya di bawah rata-rata (13.8%). "
             "Hampir semua kasus pengiriman gagal (243 dari 247) berasal dari COD.",
             "Lakukan konfirmasi sebelum pengiriman untuk order COD. Evaluasi COD dari jumlah order gagal, tidak hanya dari rate."),
            ("Online Payment menyumbang 243 dari 431 order yang batal karena belum dibayar. "
             "Tanpa kasus ini, rate Online Payment turun dari 18.9% menjadi 11.0%.",
             "Fokuskan pengingat pembayaran pada Online Payment. Pembayaran di gerai (Indomaret/Alfamart) hanya 25 kasus."),
            ("223 pembatalan (8.1%) berkaitan dengan sisi penjual: terlambat kirim, terlambat atur pengiriman, "
             "chat tidak direspons, dan stok habis. Sebagian besar tercatat sebagai pembatalan sistem.",
             "Tetapkan target waktu proses pengiriman, perbarui stok secara rutin, dan percepat respons chat."),
            ("Rak / Rak Serbaguna memiliki rate 31.9% dari 568 order, sekitar 101 order gagal lebih banyak "
             "dibanding jika rate-nya sama dengan rata-rata.",
             "Tinjau listing Rak Serbaguna (deskripsi, ukuran, ongkir) sebagai prioritas perbaikan kategori."),
            ("Rate naik pada Mar–Jun 2025 dengan puncak 22.1% di bulan Mei, lalu turun ke kisaran 9.5–12.4% "
             "sejak Jul 2025. Penyebab kenaikan belum diidentifikasi.",
             "Lakukan analisis lanjutan periode Mar–Jun 2025 berdasarkan metode bayar, kategori, dan keterlambatan penjual."),
        ]
        PT, ROW_H, HEAD_H = 940, 36, 58
        PH = HEAD_H + ROW_H * len(INSIGHTS) + 26
        box(d, X0, PT, W, PH, WHITE, rounded=True)
        box(d, X0 + 8, PT + 14, 4, PH - 28, PINK)  # aksen kiri, di dalam sudut membulat
        text(d, X0 + 20, PT + 10, 600, 18, "Insight & Keputusan", 12, INK, bold=True)
        text(d, X0 + 20, PT + 28, 900, 14,
             "Ringkasan untuk periode Jan 2024 – Nov 2025. Bagian ini tidak berubah mengikuti filter.",
             9, GREY_TXT)
        C_NUM, C_INS, C_DEC = X0 + 20, X0 + 52, X0 + 640
        W_INS, W_DEC = 570, W - (C_DEC - X0) - 24
        for x, w, lbl in [(C_INS, W_INS, "INSIGHT"), (C_DEC, W_DEC, "KEPUTUSAN")]:
            text(d, x, PT + HEAD_H - 12, w, 12, lbl, 8, GREY_TXT, bold=True)
        for i, (ins, dec) in enumerate(INSIGHTS):
            top = PT + HEAD_H + 4 + i * ROW_H
            ln = d.Shapes.AddLine(C_NUM, top - 3, X0 + W - 20, top - 3)
            ln.Line.ForeColor.RGB = rgb("#E5E7EB")
            ln.Line.Weight = 0.75
            badge = d.Shapes.AddShape(9, C_NUM, top + 4, 20, 20)  # msoShapeOval
            badge.Fill.ForeColor.RGB = PINK
            badge.Line.Visible = False
            tf = badge.TextFrame2
            tf.MarginLeft = tf.MarginRight = tf.MarginTop = tf.MarginBottom = 0
            tf.VerticalAnchor = 3  # msoAnchorMiddle
            tf.TextRange.Text = str(i + 1)
            tf.TextRange.Font.Name = FONT
            tf.TextRange.Font.Size = 9
            tf.TextRange.Font.Bold = True
            tf.TextRange.Font.Fill.ForeColor.RGB = WHITE
            tf.TextRange.ParagraphFormat.Alignment = 2
            text(d, C_INS, top + 2, W_INS, ROW_H - 4, ins, 9.5, INK)
            text(d, C_DEC, top + 2, W_DEC, ROW_H - 4, dec, 9.5, rgb("#13786B"))
        text(d, X0 + 20, PT + PH - 18, W - 40, 12,
             "Angka dihitung dari data transaksi dan menunjukkan keterkaitan, belum tentu sebab-akibat. "
             "Efek setiap keputusan perlu diukur setelah diterapkan.", 8, GREY_TXT, italic=True)

        # footnote
        FT = PT + PH + 8
        box(d, X0, FT, W, 70, NAVY, rounded=True)
        text(d, X0 + 14, FT + 10, W - 28, 56,
             "Catatan:  (1) Nilai kerugian Rupiah tidak dihitung karena total pembayaran tidak konsisten setelah retur.  "
             "(2) Rate dihitung dari order settled; 250 order berstatus Dalam Proses (Nov 2025) tidak disertakan.  "
             "(3) Beberapa label alasan dengan arti sama digabung, misalnya 'Need to change delivery address' "
             "dan 'Perlu mengubah alamat pengiriman'.\n"
             "(4) Filter Bulan hanya berlaku untuk KPI, Sumber Pembatalan, dan Alasan Pembatalan, karena rate per kelompok "
             "kurang stabil jika jumlah order kecil. Definisi lengkap ada di sheet README.",
             9, rgb("#E5E7EB"))

        # ---------------------------------------------------------- 5. README
        rd = ws_readme
        readme = [
            ("Dashboard Analisis Pembatalan & Retur: README", True),
            ("", False),
            ("Tujuan", True),
            ("Mendeskripsikan pola order yang gagal menjadi revenue (Batal + Retur Penuh): tren bulanan, sumber & alasan "
             "pembatalan, serta keterkaitannya dengan metode pembayaran dan kategori produk. Jenis analisis: deskriptif dan diagnostik.", False),
            ("", False),
            ("Sumber data & alur", True),
            ("RAW_PUBLIC_Dataset → data_cleaning.ipynb → Clean_data/master_clean_orders_FIXED.xlsx → 02_analysis.ipynb → "
             "Clean_data/analysis_dataset.xlsx → Power Query (q_orders) → tbl_orders → pivot → chart_data → dashboard.", False),
            ("Refresh: Data > Refresh All. Path Power Query bersifat absolut; jika folder dipindah, ubah lewat Data > Queries & Connections > q_orders.", False),
            ("", False),
            ("Definisi", True),
            ("Order settled: is_settled = 1 (status akhir sudah diketahui). 250 order 'Dalam Proses' (Nov 2025) dikeluarkan dari denominator.", False),
            ("Gagal-Revenue Rate = (Batal + Retur Penuh) / order settled. Calculated field rate_gagal = is_gagal_revenue / is_settled.", False),
            ("Sumber & Alasan memakai basis order batal (is_cancelled = 1); retur penuh tidak termasuk karena tidak punya alasan pembatalan.", False),
            ("Metode & kategori hanya ditampilkan bila n ≥ 30 order settled (dihitung atas seluruh 23 bulan). Kategori: hanya order 1 kategori.", False),
            ("", False),
            ("Hubungan slicer", True),
            ("Filter Tahun: KPI, Sumber, Metode, Kategori, Alasan. Filter Bulan: KPI, Sumber, Alasan. Filter Metode Bayar: KPI, Tren, Sumber, Kategori, Alasan.", False),
            ("Tren tidak ikut Tahun/Bulan agar konteks 23 bulan tetap utuh. Chart Metode tidak ikut slicer Metode karena fungsinya membandingkan semua metode.", False),
            ("", False),
            ("Batasan metodologi", True),
            ("1. Kerugian Rupiah tidak dihitung: total_pembayaran_idr tidak konsisten setelah retur (retur penuh kadang 0, kadang tetap nilai awal).", False),
            ("2. waktu_pesanan sebagian berupa proxy (jadwal kirim batch), sehingga hanya dipakai untuk agregasi bulanan (bulan_file), tidak untuk analisis jam/hari.", False),
            ("3. 724 order dengan opsi pengiriman generik (tanpa kurir) selalu berstatus Batal karena efek pencatatan. Order ini tetap dihitung di KPI, tetapi dikecualikan dari analisis opsi pengiriman.", False),
            ("4. Jumlah retur sangat sedikit (118 order) dan kolom retur Jan 2024 kosong, sehingga retur hanya menjadi catatan sekunder.", False),
            ("5. Label alasan dengan arti sama digabung dari 18 menjadi 16 label. Pemetaan lengkap ada di 02_analysis.ipynb bagian 2.", False),
            ("6. Temuan menunjukkan keterkaitan antarvariabel dan tidak membuktikan hubungan sebab-akibat.", False),
        ]
        for i, (t, bold) in enumerate(readme, start=1):
            c = rd.Range(f"A{i}")
            c.Value = t
            c.Font.Bold = bold
            c.Font.Name = FONT
            c.Font.Size = 14 if i == 1 else 10
        rd.Columns("A").ColumnWidth = 160
        rd.Range(f"A1:A{len(readme)}").WrapText = True

        # ---------------------------------------------------------- 6. finalisasi
        wb.RefreshAll()
        xl.CalculateUntilAsyncQueriesDone()
        xl.Calculate()
        d.Activate()
        d.Range("A1").Select()
        xl.ActiveWindow.Zoom = 80
        wb.SaveAs(tmp, FileFormat=xlOpenXMLWorkbook)

        # ringkasan untuk verifikasi
        print("KPI:", [cd.Range(f"C{r}").Text for r in (4, 5, 6)])
        print("Tren Mei 2025:", cd.Range("B26").Text, "| rata-rata:", cd.Range("E10").Text)
        print("Sumber:", [(cd.Range(f"A{r}").Text, cd.Range(f"B{r}").Text) for r in (36, 37, 38)])
        print("Metode:", [(cd.Range(f"A{r}").Text, cd.Range(f"D{r}").Text) for r in range(43, 51)],
              "| rata2:", cd.Range("B52").Text)
        print("Kategori:", [(cd.Range(f"A{r}").Text, cd.Range(f"D{r}").Text) for r in range(56, 61)])
        print("Alasan:", [(cd.Range(f"A{r}").Text, cd.Range(f"C{r}").Text) for r in range(64, 69)],
              "| kumulatif:", cd.Range("C69").Text)
        print("Urutan slicer Bulan:", [it.Name for it in wb.SlicerCaches("Slicer_bulan").SlicerItems])
        wb.Close(False)
    finally:
        xl.Quit()
    os.replace(tmp, OUT)  # gagal kalau OUT sedang dibuka di Excel; file lama tetap utuh
    print("Tersimpan:", OUT)


if __name__ == "__main__":
    build()
