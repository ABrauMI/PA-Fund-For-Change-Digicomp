"""
Builds a multi-tab digital competitive report (one tab per race, plus a
Summary index) from an AdHawk CSV export. Generalized from the one-off
PA report so it keeps working as new exports arrive with different
races, different week ranges, and possibly non-partisan spenders.
"""

import re
from datetime import date, timedelta

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

PLATFORM_ORDER = {'CTV': 0, 'Google': 1, 'Facebook': 2, 'Twitter': 3}

NAVY = '323b51'
WHITE = 'FFFFFF'
R_LIGHT, R_TOTAL, R_PARTY = 'fcefed', 'f2c1bb', 'de5e4e'
D_LIGHT, D_TOTAL, D_PARTY = 'ecf0f4', 'b4c6d5', '3d6a91'
NP_LIGHT, NP_TOTAL, NP_PARTY = 'fef6ed', 'fcddb8', 'd37518'
FOOTER_GREY = '6B7280'
FONT_TITLE = 'Superior Title'
FONT_BODY = 'Figtree'

PARTY_COLORS = {
    'R': (R_LIGHT, R_TOTAL, R_PARTY, 'REPUBLICAN PARTY TOTAL'),
    'D': (D_LIGHT, D_TOTAL, D_PARTY, 'DEMOCRAT PARTY TOTAL'),
    'NP': (NP_LIGHT, NP_TOTAL, NP_PARTY, 'NON-PARTISAN PARTY TOTAL'),
}
PARTY_ORDER = ['R', 'D', 'NP']

MEDIUM_BOTTOM = Border(bottom=Side(style='medium'))
DOUBLE_BOTTOM = Border(bottom=Side(style='double'))

WEEK_COL_RE = re.compile(r'^(\d{2}/\d{2})-(\d{2}/\d{2}) (DEM|GOP|total)$')


def fill(hexcolor):
    return PatternFill('solid', fgColor=hexcolor)


def font(bold=False, size=9, color=NAVY, name=FONT_BODY):
    return Font(name=name, size=size, bold=bold, color=color)


def _week_labels_from_header(columns):
    labels = []
    for col in columns:
        m = WEEK_COL_RE.match(col)
        if m and m.group(3) == 'DEM' and col not in labels:
            label = f'{m.group(1)}-{m.group(2)}'
            if label not in labels:
                labels.append(label)
    return labels


def _infer_years(week_labels, reference=None):
    """AdHawk week labels have no year. Anchor the last (most recent) week
    to today, then walk backward, decrementing the year every time we cross
    a Dec -> Jan boundary."""
    reference = reference or date.today()
    parsed = [(int(lbl[:2]), int(lbl[3:5])) for lbl in week_labels]  # (month, day) of week start
    years = [None] * len(parsed)

    last_m, last_d = parsed[-1]
    year = reference.year
    try:
        candidate = date(year, last_m, last_d)
    except ValueError:
        candidate = date(year, last_m, min(last_d, 28))
    if candidate > reference + timedelta(days=45):
        year -= 1
    elif candidate < reference - timedelta(days=400):
        year += 1
    years[-1] = year

    for i in range(len(parsed) - 2, -1, -1):
        m_this, _ = parsed[i]
        m_next, _ = parsed[i + 1]
        years[i] = years[i + 1] - 1 if m_this > m_next else years[i + 1]

    return years


def _natural_sort_key(code):
    parts = re.split(r'(\d+)', code)
    return tuple(int(p) if p.isdigit() else p for p in parts)


def _race_group_and_number(race):
    m = re.match(r'^(.*?)-(\d+)$', race)
    if m:
        return m.group(1), m.group(2)
    return race, ''


def build_workbook(csv_path, out_path, reference_date=None):
    df = pd.read_csv(csv_path)
    df = df.dropna(subset=['Election Name'])
    if df.empty:
        raise ValueError('No races found in this export (no rows with an Election Name).')

    week_labels = _week_labels_from_header(df.columns)
    if not week_labels:
        raise ValueError('Could not find any weekly DEM/GOP/total columns in this CSV.')
    years = _infer_years(week_labels, reference_date)
    week_headers = [f'{lbl.split("-")[0]}/{yr}' for lbl, yr in zip(week_labels, years)]

    races = sorted(df['Election Name'].unique(), key=_natural_sort_key)

    records_by_race = {}
    for race in races:
        sub = df[df['Election Name'] == race]
        recs = []
        for _, row in sub.iterrows():
            dem, gop = float(row.get('DEM', 0) or 0), float(row.get('GOP', 0) or 0)
            party = 'D' if dem > 0 else ('R' if gop > 0 else 'NP')
            party_col = 'DEM' if party == 'D' else ('GOP' if party == 'R' else 'total')
            weekly = []
            for wk in week_labels:
                col = f'{wk} {party_col}'
                val = row.get(col)
                weekly.append(float(val) if pd.notna(val) else 0.0)
            recs.append({
                'advertiser': row['Advertiser'],
                'party': party,
                'platform': row['Spend Platform'],
                'weekly': weekly,
            })
        records_by_race[race] = recs

    group_tokens = {_race_group_and_number(r)[0] for r in races}
    prefix_tokens = {r.split('-')[0] for r in races}
    if len(prefix_tokens) == 1:
        title_prefix = next(iter(prefix_tokens))
        report_title = f'{title_prefix} DIGITAL COMPETITIVE REPORT'
    else:
        report_title = 'DIGITAL COMPETITIVE REPORT'

    last_col = 4 + len(week_labels)
    last_col_letter = get_column_letter(last_col)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    def style_header_bar(ws, title_text, ncols, title_cols=4):
        ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=title_cols)
        a1 = ws.cell(row=1, column=1, value=f'   {title_text}')
        a1.font = font(bold=True, size=14, color=WHITE, name=FONT_TITLE)
        a1.alignment = Alignment(horizontal='right', vertical='center')
        ws.merge_cells(start_row=1, start_column=title_cols + 1, end_row=1, end_column=ncols)
        for r in (1, 2):
            for col in range(1, ncols + 1):
                ws.cell(row=r, column=col).fill = fill(NAVY)
        ws.row_dimensions[1].height = 34
        ws.row_dimensions[2].height = 18

    def apply_print_setup(ws, ncols):
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_title_rows = '1:3'

    def set_col_widths(ws):
        ws.column_dimensions['A'].width = 34
        ws.column_dimensions['B'].width = 8
        ws.column_dimensions['C'].width = 16
        ws.column_dimensions['D'].width = 13
        for col in range(5, last_col + 1):
            ws.column_dimensions[get_column_letter(col)].width = 12

    def build_race_sheet(race, recs):
        ws = wb.create_sheet(title=race[:31])
        ncols = last_col
        style_header_bar(ws, race, ncols)

        headers = ['CANDIDATE / COMMITTEE', 'PARTY', 'PLATFORM', 'TOTAL SPEND'] + week_headers
        for col, text in enumerate(headers, start=1):
            c = ws.cell(row=3, column=col, value=text)
            c.font = font(bold=True, size=9, color=WHITE)
            c.fill = fill(NAVY)
            c.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[3].height = 30

        set_col_widths(ws)
        ws.freeze_panes = 'E4'
        ws.sheet_view.showGridLines = False
        apply_print_setup(ws, ncols)

        row = 4
        party_label_row = {}

        for party in PARTY_ORDER:
            party_recs = [r for r in recs if r['party'] == party]
            if not party_recs:
                continue
            light, totalfill, party_fill, party_name = PARTY_COLORS[party]

            by_adv = {}
            for r in party_recs:
                by_adv.setdefault(r['advertiser'], []).append(r)
            adv_order = sorted(by_adv.keys(), key=lambda a: -sum(sum(x['weekly']) for x in by_adv[a]))
            adv_total_rows = []

            for adv in adv_order:
                plats = sorted(by_adv[adv], key=lambda r: PLATFORM_ORDER.get(r['platform'], 9))
                start_row = row
                for r in plats:
                    ws.cell(row=row, column=1, value=adv).font = font(bold=True, size=9)
                    ws.cell(row=row, column=1).fill = fill(light)
                    ws.cell(row=row, column=2, value=party).font = font(bold=True, size=9)
                    ws.cell(row=row, column=2).fill = fill(light)
                    ws.cell(row=row, column=2).alignment = Alignment(horizontal='center')
                    c3 = ws.cell(row=row, column=3, value=r['platform'])
                    c3.font = font(bold=False, size=9)
                    c3.fill = fill(light)
                    for i, wk_val in enumerate(r['weekly']):
                        col = 5 + i
                        cell = ws.cell(row=row, column=col, value=wk_val)
                        cell.font = font(bold=False, size=9)
                        cell.fill = fill(light)
                        cell.number_format = '$#,##0;-$#,##0;""'
                    dcell = ws.cell(row=row, column=4, value=f'=SUM(E{row}:{last_col_letter}{row})')
                    dcell.font = font(bold=False, size=9)
                    dcell.fill = fill(light)
                    dcell.number_format = '$#,##0;-$#,##0;""'
                    row += 1
                end_row = row - 1
                ws.merge_cells(start_row=start_row, start_column=1, end_row=row, end_column=1)
                ws.merge_cells(start_row=start_row, start_column=2, end_row=row, end_column=2)

                tot_label = ws.cell(row=row, column=3, value=f'{adv} Total')
                tot_label.font = font(bold=True, size=9)
                tot_label.fill = fill(totalfill)
                tot_label.border = MEDIUM_BOTTOM
                ws.cell(row=row, column=1).fill = fill(totalfill)
                ws.cell(row=row, column=1).border = MEDIUM_BOTTOM
                ws.cell(row=row, column=2).fill = fill(totalfill)
                ws.cell(row=row, column=2).border = MEDIUM_BOTTOM
                for i in range(len(week_labels)):
                    col = 5 + i
                    col_letter = get_column_letter(col)
                    cell = ws.cell(row=row, column=col, value=f'=SUM({col_letter}{start_row}:{col_letter}{end_row})')
                    cell.font = font(bold=True, size=9)
                    cell.fill = fill(totalfill)
                    cell.number_format = '$#,##0'
                    cell.border = MEDIUM_BOTTOM
                dcell = ws.cell(row=row, column=4, value=f'=SUM(E{row}:{last_col_letter}{row})')
                dcell.font = font(bold=True, size=9)
                dcell.fill = fill(totalfill)
                dcell.number_format = '$#,##0'
                dcell.border = MEDIUM_BOTTOM

                adv_total_rows.append(row)
                row += 2

            label_row = row
            party_label_row[party] = label_row
            ws.merge_cells(start_row=label_row, start_column=1, end_row=label_row, end_column=3)
            lab = ws.cell(row=label_row, column=1, value=party_name)
            lab.font = font(bold=True, size=10, color=WHITE)
            for col in range(1, 4):
                ws.cell(row=label_row, column=col).fill = fill(party_fill)
                ws.cell(row=label_row, column=col).border = MEDIUM_BOTTOM
            for i in range(len(week_labels)):
                col = 5 + i
                col_letter = get_column_letter(col)
                formula = '+'.join(f'{col_letter}{r}' for r in adv_total_rows)
                cell = ws.cell(row=label_row, column=col, value=f'={formula}')
                cell.font = font(bold=True, size=10, color=WHITE)
                cell.fill = fill(party_fill)
                cell.number_format = '$#,##0'
                cell.border = MEDIUM_BOTTOM
            dcell = ws.cell(row=label_row, column=4, value=f'=SUM(E{label_row}:{last_col_letter}{label_row})')
            dcell.font = font(bold=True, size=10, color=WHITE)
            dcell.fill = fill(party_fill)
            dcell.number_format = '$#,##0'
            dcell.border = MEDIUM_BOTTOM
            row = label_row + 2

        grand_row = row
        ws.merge_cells(start_row=grand_row, start_column=1, end_row=grand_row, end_column=3)
        lab = ws.cell(row=grand_row, column=1, value='GRAND TOTAL')
        lab.font = font(bold=True, size=10, color=WHITE)
        for col in range(1, 4):
            ws.cell(row=grand_row, column=col).fill = fill(NAVY)
            ws.cell(row=grand_row, column=col).border = DOUBLE_BOTTOM
        party_rows = list(party_label_row.values())
        for i in range(len(week_labels)):
            col = 5 + i
            col_letter = get_column_letter(col)
            formula = '+'.join(f'{col_letter}{r}' for r in party_rows)
            cell = ws.cell(row=grand_row, column=col, value=f'={formula}')
            cell.font = font(bold=True, size=10, color=WHITE)
            cell.fill = fill(NAVY)
            cell.number_format = '$#,##0'
            cell.border = DOUBLE_BOTTOM
        dcell = ws.cell(row=grand_row, column=4, value=f'=SUM(E{grand_row}:{last_col_letter}{grand_row})')
        dcell.font = font(bold=True, size=10, color=WHITE)
        dcell.fill = fill(NAVY)
        dcell.number_format = '$#,##0'
        dcell.border = DOUBLE_BOTTOM

        footer_row = grand_row + 2
        ws.merge_cells(start_row=footer_row, start_column=1, end_row=footer_row, end_column=last_col)
        f = ws.cell(row=footer_row, column=1, value='Report prepared by GPS Impact  |  Confidential')
        f.font = font(bold=False, size=8, color=FOOTER_GREY)

        return {'grand_row': grand_row, 'party_row': party_label_row}

    def build_summary_sheet(race_meta):
        ws = wb.create_sheet(title='Summary', index=0)
        ncols = 7
        style_header_bar(ws, report_title, ncols, title_cols=2)

        headers = ['RACE', 'GROUP', 'NUMBER', 'TOTAL SPEND', 'GOP TOTAL', 'DEM TOTAL', 'TOP SPENDER']
        for col, text in enumerate(headers, start=1):
            c = ws.cell(row=3, column=col, value=text)
            c.font = font(bold=True, size=9, color=WHITE)
            c.fill = fill(NAVY)
            c.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[3].height = 24

        widths = {'A': 18, 'B': 16, 'C': 10, 'D': 14, 'E': 14, 'F': 14, 'G': 32}
        for col, w in widths.items():
            ws.column_dimensions[col].width = w
        ws.freeze_panes = 'A4'
        ws.sheet_view.showGridLines = False
        apply_print_setup(ws, ncols)

        row = 4
        first_row = row
        for race in races:
            meta = race_meta[race]
            group, number = _race_group_and_number(race)
            recs = records_by_race[race]
            by_adv_total = {}
            for r in recs:
                by_adv_total[r['advertiser']] = by_adv_total.get(r['advertiser'], 0.0) + sum(r['weekly'])
            top_spender = max(by_adv_total, key=by_adv_total.get)

            c = ws.cell(row=row, column=1, value=race)
            c.hyperlink = f"#'{race[:31]}'!A1"
            c.font = Font(name=FONT_BODY, size=9, bold=True, color='1155CC', underline='single')

            ws.cell(row=row, column=2, value=group).font = font(size=9)
            ws.cell(row=row, column=3, value=number).font = font(size=9)
            ws.cell(row=row, column=3).alignment = Alignment(horizontal='center')

            grand_row = meta['grand_row']
            r_row = meta['party_row'].get('R')
            d_row = meta['party_row'].get('D')

            dcell = ws.cell(row=row, column=4, value=f"='{race[:31]}'!D{grand_row}")
            dcell.number_format = '$#,##0'
            dcell.font = font(bold=True, size=9)

            ecell = ws.cell(row=row, column=5, value=(f"='{race[:31]}'!D{r_row}" if r_row else 0))
            ecell.number_format = '$#,##0'
            ecell.font = font(size=9)

            fcell = ws.cell(row=row, column=6, value=(f"='{race[:31]}'!D{d_row}" if d_row else 0))
            fcell.number_format = '$#,##0'
            fcell.font = font(size=9)

            ws.cell(row=row, column=7, value=top_spender).font = font(size=9)
            row += 1

        last_row = row - 1
        total_row = row
        ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=3)
        lab = ws.cell(row=total_row, column=1, value='TOTAL — ALL RACES')
        lab.font = font(bold=True, size=10, color=WHITE)
        for col in range(1, 8):
            ws.cell(row=total_row, column=col).fill = fill(NAVY)
            ws.cell(row=total_row, column=col).border = DOUBLE_BOTTOM
        for col in (4, 5, 6):
            col_letter = get_column_letter(col)
            cell = ws.cell(row=total_row, column=col, value=f'=SUM({col_letter}{first_row}:{col_letter}{last_row})')
            cell.font = font(bold=True, size=10, color=WHITE)
            cell.number_format = '$#,##0'
        ws.cell(row=total_row, column=7).font = font(bold=True, size=10, color=WHITE)

        footer_row = total_row + 2
        ws.merge_cells(start_row=footer_row, start_column=1, end_row=footer_row, end_column=7)
        f = ws.cell(row=footer_row, column=1, value='Report prepared by GPS Impact  |  Confidential')
        f.font = font(bold=False, size=8, color=FOOTER_GREY)

    race_meta = {race: build_race_sheet(race, records_by_race[race]) for race in races}
    build_summary_sheet(race_meta)

    wb.save(out_path)
    return {'races': races, 'out_path': out_path}
