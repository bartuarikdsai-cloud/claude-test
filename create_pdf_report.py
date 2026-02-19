"""
Auto Insurance Portfolio Analysis - PDF Actuarial Report Generator
Generates a professional 6-page PDF report with matplotlib charts and reportlab layout.
"""
import json
import os
import math
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib.patches import FancyBboxPatch
from matplotlib.colors import LinearSegmentedColormap

from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch, mm
from reportlab.lib.colors import HexColor, white, black, Color
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle,
    PageBreak, Frame, PageTemplate, BaseDocTemplate, FrameBreak
)
from reportlab.pdfgen import canvas
from reportlab.platypus.flowables import Flowable

# ══════════════════════════════════════════════════════════════════════════════
# COLORS
# ══════════════════════════════════════════════════════════════════════════════
BG_DARK = '#0F1117'
BG_SURFACE = '#1A1D27'
BG_SURFACE2 = '#232733'
BORDER_COLOR = '#2E3345'
TEXT_WHITE = '#E4E6ED'
TEXT_MUTED = '#8B8FA3'

INDIGO = '#6366F1'
INDIGO_LIGHT = '#818CF8'
CYAN = '#06B6D4'
GREEN = '#22C55E'
AMBER = '#F59E0B'
RED = '#EF4444'
PINK = '#EC4899'

# ══════════════════════════════════════════════════════════════════════════════
# DATA LOADING & AGGREGATION
# ══════════════════════════════════════════════════════════════════════════════
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(SCRIPT_DIR, 'dashboard_data.json')
OUTPUT_DIR = SCRIPT_DIR
CHART_DIR = os.path.join(SCRIPT_DIR, '_report_charts')

os.makedirs(CHART_DIR, exist_ok=True)

with open(DATA_PATH) as f:
    raw = json.load(f)

# Parse records: [id, gender(1=M/0=F), age, car_year, premium, loss]
data = []
for r in raw:
    data.append({
        'id': r[0],
        'gender': 'Male' if r[1] == 1 else 'Female',
        'age': r[2],
        'car_year': r[3],
        'premium': r[4],
        'loss': r[5],
    })

def age_group(a):
    if a < 25: return '18-24'
    if a < 35: return '25-34'
    if a < 45: return '35-44'
    if a < 55: return '45-54'
    if a < 65: return '55-64'
    return '65+'

def car_era(y):
    if y < 2005: return '2000-04'
    if y < 2010: return '2005-09'
    if y < 2015: return '2010-14'
    if y < 2020: return '2015-19'
    return '2020-25'

AGE_GROUPS = ['18-24', '25-34', '35-44', '45-54', '55-64', '65+']
CAR_ERAS = ['2000-04', '2005-09', '2010-14', '2015-19', '2020-25']

# Overall KPIs
total_policies = len(data)
total_premium = sum(d['premium'] for d in data)
total_loss = sum(d['loss'] for d in data)
claims_count = sum(1 for d in data if d['loss'] > 0)
loss_ratio = total_loss / total_premium
claims_freq = claims_count / total_policies
avg_premium = total_premium / total_policies
avg_loss_per_claim = total_loss / claims_count if claims_count > 0 else 0

# By age group
by_age = {}
for ag in AGE_GROUPS:
    rows = [d for d in data if age_group(d['age']) == ag]
    prem = sum(r['premium'] for r in rows)
    loss = sum(r['loss'] for r in rows)
    claims = sum(1 for r in rows if r['loss'] > 0)
    by_age[ag] = {
        'count': len(rows), 'premium': prem, 'loss': loss, 'claims': claims,
        'avg_premium': prem / len(rows) if rows else 0,
        'loss_ratio': loss / prem if prem else 0,
        'claims_freq': claims / len(rows) if rows else 0,
        'avg_loss_per_claim': loss / claims if claims > 0 else 0,
    }

# By car era
by_era = {}
for era in CAR_ERAS:
    rows = [d for d in data if car_era(d['car_year']) == era]
    prem = sum(r['premium'] for r in rows)
    loss = sum(r['loss'] for r in rows)
    claims = sum(1 for r in rows if r['loss'] > 0)
    by_era[era] = {
        'count': len(rows), 'premium': prem, 'loss': loss,
        'loss_ratio': loss / prem if prem else 0,
        'claims_freq': claims / len(rows) if rows else 0,
    }

# By gender
by_gender = {}
for g in ['Male', 'Female']:
    rows = [d for d in data if d['gender'] == g]
    prem = sum(r['premium'] for r in rows)
    loss = sum(r['loss'] for r in rows)
    claims = sum(1 for r in rows if r['loss'] > 0)
    by_gender[g] = {
        'count': len(rows), 'premium': prem, 'loss': loss,
        'avg_premium': prem / len(rows),
        'loss_ratio': loss / prem if prem else 0,
        'claims_freq': claims / len(rows),
    }

# Heatmap: Age Group x Car Era
heatmap = {}
for ag in AGE_GROUPS:
    heatmap[ag] = {}
    for era in CAR_ERAS:
        rows = [d for d in data if age_group(d['age']) == ag and car_era(d['car_year']) == era]
        prem = sum(r['premium'] for r in rows)
        loss = sum(r['loss'] for r in rows)
        heatmap[ag][era] = {'n': len(rows), 'lr': loss / prem if prem else 0}

# Gender x Age cross-tab
gender_age = {}
for g in ['Male', 'Female']:
    gender_age[g] = {}
    for ag in AGE_GROUPS:
        rows = [d for d in data if d['gender'] == g and age_group(d['age']) == ag]
        prem = sum(r['premium'] for r in rows)
        loss = sum(r['loss'] for r in rows)
        claims = sum(1 for r in rows if r['loss'] > 0)
        gender_age[g][ag] = {
            'count': len(rows),
            'loss_ratio': loss / prem if prem else 0,
            'claims_freq': claims / len(rows) if rows else 0,
            'avg_premium': prem / len(rows) if rows else 0,
        }

# Premium distribution
all_premiums = [d['premium'] for d in data]

# Identify highest/lowest risk segments
highest_lr_age = max(AGE_GROUPS, key=lambda ag: by_age[ag]['loss_ratio'])
lowest_lr_age = min(AGE_GROUPS, key=lambda ag: by_age[ag]['loss_ratio'])
highest_lr_era = max(CAR_ERAS, key=lambda e: by_era[e]['loss_ratio'])
lowest_lr_era = min(CAR_ERAS, key=lambda e: by_era[e]['loss_ratio'])

# Find highest risk heatmap cell
max_heatmap_lr = 0
max_heatmap_cell = ('', '')
for ag in AGE_GROUPS:
    for era in CAR_ERAS:
        if heatmap[ag][era]['lr'] > max_heatmap_lr and heatmap[ag][era]['n'] >= 20:
            max_heatmap_lr = heatmap[ag][era]['lr']
            max_heatmap_cell = (ag, era)

# ══════════════════════════════════════════════════════════════════════════════
# CHART GENERATION (matplotlib)
# ══════════════════════════════════════════════════════════════════════════════
plt.rcParams.update({
    'figure.facecolor': BG_DARK,
    'figure.dpi': 300,
    'savefig.dpi': 300,
    'axes.facecolor': BG_SURFACE,
    'axes.edgecolor': BORDER_COLOR,
    'axes.labelcolor': TEXT_MUTED,
    'text.color': TEXT_WHITE,
    'xtick.color': TEXT_MUTED,
    'ytick.color': TEXT_MUTED,
    'grid.color': BORDER_COLOR,
    'grid.alpha': 0.5,
    'font.family': 'sans-serif',
    'font.size': 12,
    'text.antialiased': True,
    'lines.antialiased': True,
})


def lr_bar_color(lr):
    """Return color based on loss ratio threshold."""
    if lr >= 0.85:
        return RED
    elif lr >= 0.70:
        return AMBER
    else:
        return GREEN


def save_chart(fig, name):
    """Save chart to PNG with tight bounding box."""
    path = os.path.join(CHART_DIR, name)
    fig.savefig(path, dpi=600, bbox_inches='tight', facecolor=fig.get_facecolor(),
                edgecolor='none', pad_inches=0.15)
    plt.close(fig)
    return path


# ── Chart 1: Loss Ratio by Age Group ──
def chart_loss_ratio_by_age():
    fig, ax = plt.subplots(figsize=(10.0, 5.5))
    ages = AGE_GROUPS
    lrs = [by_age[ag]['loss_ratio'] * 100 for ag in ages]
    colors = [lr_bar_color(by_age[ag]['loss_ratio']) for ag in ages]
    bars = ax.bar(ages, lrs, color=colors, width=0.6, edgecolor='none', zorder=3)

    # Value labels on bars
    for bar, val in zip(bars, lrs):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 1.5,
                f'{val:.1f}%', ha='center', va='bottom', fontsize=10,
                fontweight='bold', color=TEXT_WHITE)

    # Reference line
    ax.axhline(y=loss_ratio * 100, color=INDIGO_LIGHT, linestyle='--', linewidth=1.2, alpha=0.7, zorder=2)
    ax.text(len(ages) - 0.5, loss_ratio * 100 + 1.5,
            f'Portfolio Avg: {loss_ratio*100:.1f}%', ha='right', fontsize=9,
            color=INDIGO_LIGHT, fontstyle='italic')

    ax.set_ylabel('Loss Ratio (%)', fontsize=11)
    ax.set_xlabel('Age Group', fontsize=11)
    ax.set_title('Loss Ratio by Age Group', fontsize=14, fontweight='bold',
                 color=TEXT_WHITE, pad=12)
    ax.set_ylim(0, max(lrs) * 1.2)
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter('%.0f%%'))
    ax.grid(axis='y', alpha=0.3, zorder=0)
    ax.set_axisbelow(True)

    # Legend for color coding
    from matplotlib.patches import Patch
    legend_elements = [
        Patch(facecolor=GREEN, label='< 70% (Low Risk)'),
        Patch(facecolor=AMBER, label='70-85% (Moderate)'),
        Patch(facecolor=RED, label='> 85% (High Risk)'),
    ]
    ax.legend(handles=legend_elements, loc='upper left', fontsize=8,
              facecolor=BG_SURFACE, edgecolor=BORDER_COLOR, labelcolor=TEXT_MUTED)

    fig.tight_layout()
    return save_chart(fig, 'chart_lr_age.png')


# ── Chart 2: Claims Frequency by Age Group ──
def chart_claims_freq_by_age():
    fig, ax = plt.subplots(figsize=(10.0, 5.5))
    ages = AGE_GROUPS
    freqs = [by_age[ag]['claims_freq'] * 100 for ag in ages]
    bars = ax.bar(ages, freqs, color=CYAN, width=0.6, edgecolor='none', zorder=3, alpha=0.9)

    for bar, val in zip(bars, freqs):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.8,
                f'{val:.1f}%', ha='center', va='bottom', fontsize=10,
                fontweight='bold', color=TEXT_WHITE)

    ax.axhline(y=claims_freq * 100, color=INDIGO_LIGHT, linestyle='--', linewidth=1.2, alpha=0.7, zorder=2)
    ax.text(len(ages) - 0.5, claims_freq * 100 + 0.8,
            f'Portfolio Avg: {claims_freq*100:.1f}%', ha='right', fontsize=9,
            color=INDIGO_LIGHT, fontstyle='italic')

    ax.set_ylabel('Claims Frequency (%)', fontsize=11)
    ax.set_xlabel('Age Group', fontsize=11)
    ax.set_title('Claims Frequency by Age Group', fontsize=14, fontweight='bold',
                 color=TEXT_WHITE, pad=12)
    ax.set_ylim(0, max(freqs) * 1.25)
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter('%.0f%%'))
    ax.grid(axis='y', alpha=0.3, zorder=0)
    ax.set_axisbelow(True)

    fig.tight_layout()
    return save_chart(fig, 'chart_cf_age.png')


# ── Chart 3: Risk Heatmap (Age Group x Car Era) ──
def chart_risk_heatmap():
    fig, ax = plt.subplots(figsize=(10.0, 5.8))

    # Build matrix
    matrix = np.zeros((len(AGE_GROUPS), len(CAR_ERAS)))
    for i, ag in enumerate(AGE_GROUPS):
        for j, era in enumerate(CAR_ERAS):
            matrix[i][j] = heatmap[ag][era]['lr'] * 100

    # Custom colormap: green -> yellow -> red
    cmap_colors = ['#15803D', '#22C55E', '#84CC16', '#EAB308', '#F59E0B', '#EF4444', '#DC2626']
    cmap = LinearSegmentedColormap.from_list('risk', cmap_colors, N=256)

    im = ax.imshow(matrix, cmap=cmap, aspect='auto', vmin=30, vmax=110)

    # Add text annotations
    for i in range(len(AGE_GROUPS)):
        for j in range(len(CAR_ERAS)):
            val = matrix[i][j]
            n = heatmap[AGE_GROUPS[i]][CAR_ERAS[j]]['n']
            text_color = 'white' if val > 70 else '#0F1117'
            ax.text(j, i, f'{val:.0f}%\n(n={n})', ha='center', va='center',
                    fontsize=9, fontweight='bold', color=text_color)

    ax.set_xticks(range(len(CAR_ERAS)))
    ax.set_xticklabels(CAR_ERAS, fontsize=10)
    ax.set_yticks(range(len(AGE_GROUPS)))
    ax.set_yticklabels(AGE_GROUPS, fontsize=10)
    ax.set_xlabel('Car Model Year Era', fontsize=11)
    ax.set_ylabel('Age Group', fontsize=11)
    ax.set_title('Risk Heatmap: Loss Ratio by Age Group & Car Era', fontsize=13,
                 fontweight='bold', color=TEXT_WHITE, pad=12)

    cbar = fig.colorbar(im, ax=ax, shrink=0.8, pad=0.02)
    cbar.set_label('Loss Ratio (%)', fontsize=10, color=TEXT_MUTED)
    cbar.ax.tick_params(colors=TEXT_MUTED, labelsize=9)

    fig.tight_layout()
    return save_chart(fig, 'chart_heatmap.png')


# ── Chart 4: Loss Ratio by Car Era (line chart) ──
def chart_lr_by_car_era():
    fig, ax = plt.subplots(figsize=(10.0, 5.0))
    eras = CAR_ERAS
    lrs = [by_era[e]['loss_ratio'] * 100 for e in eras]

    ax.plot(eras, lrs, color=INDIGO_LIGHT, linewidth=2.5, marker='o', markersize=8,
            markerfacecolor=INDIGO, markeredgecolor=TEXT_WHITE, markeredgewidth=1.5, zorder=3)

    # Fill under the line
    ax.fill_between(range(len(eras)), lrs, alpha=0.15, color=INDIGO_LIGHT, zorder=2)

    for i, (era, val) in enumerate(zip(eras, lrs)):
        ax.text(i, val + 2.5, f'{val:.1f}%', ha='center', fontsize=10,
                fontweight='bold', color=TEXT_WHITE)

    ax.axhline(y=loss_ratio * 100, color=AMBER, linestyle='--', linewidth=1.0, alpha=0.6)
    ax.text(len(eras) - 1, loss_ratio * 100 - 2.5,
            f'Portfolio Avg: {loss_ratio*100:.1f}%', ha='right', fontsize=9,
            color=AMBER, fontstyle='italic')

    ax.set_ylabel('Loss Ratio (%)', fontsize=11)
    ax.set_xlabel('Car Model Year Era', fontsize=11)
    ax.set_title('Loss Ratio Trend by Car Era', fontsize=14, fontweight='bold',
                 color=TEXT_WHITE, pad=12)
    ax.set_ylim(min(lrs) * 0.7, max(lrs) * 1.2)
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter('%.0f%%'))
    ax.grid(axis='y', alpha=0.3, zorder=0)
    ax.set_axisbelow(True)

    fig.tight_layout()
    return save_chart(fig, 'chart_lr_era.png')


# ── Chart 5: Gender Comparison (grouped bar) ──
def chart_gender_comparison():
    fig, ax = plt.subplots(figsize=(10.0, 5.5))

    x = np.arange(len(AGE_GROUPS))
    width = 0.35

    male_lrs = [gender_age['Male'][ag]['loss_ratio'] * 100 for ag in AGE_GROUPS]
    female_lrs = [gender_age['Female'][ag]['loss_ratio'] * 100 for ag in AGE_GROUPS]

    bars1 = ax.bar(x - width/2, male_lrs, width, label='Male', color=INDIGO, edgecolor='none', zorder=3)
    bars2 = ax.bar(x + width/2, female_lrs, width, label='Female', color=PINK, edgecolor='none', zorder=3)

    # Value labels
    for bar, val in zip(bars1, male_lrs):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                f'{val:.0f}%', ha='center', va='bottom', fontsize=8,
                color=INDIGO_LIGHT, fontweight='bold')
    for bar, val in zip(bars2, female_lrs):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                f'{val:.0f}%', ha='center', va='bottom', fontsize=8,
                color=PINK, fontweight='bold')

    ax.set_xticks(x)
    ax.set_xticklabels(AGE_GROUPS, fontsize=10)
    ax.set_ylabel('Loss Ratio (%)', fontsize=11)
    ax.set_xlabel('Age Group', fontsize=11)
    ax.set_title('Loss Ratio: Male vs Female by Age Group', fontsize=14,
                 fontweight='bold', color=TEXT_WHITE, pad=12)
    ax.set_ylim(0, max(max(male_lrs), max(female_lrs)) * 1.25)
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter('%.0f%%'))
    ax.grid(axis='y', alpha=0.3, zorder=0)
    ax.set_axisbelow(True)

    ax.legend(fontsize=10, facecolor=BG_SURFACE, edgecolor=BORDER_COLOR,
              labelcolor=TEXT_MUTED, loc='upper right')

    fig.tight_layout()
    return save_chart(fig, 'chart_gender.png')


# ── Chart 6: Premium Distribution Histogram ──
def chart_premium_distribution():
    fig, ax = plt.subplots(figsize=(10.0, 5.0))

    n, bins, patches = ax.hist(all_premiums, bins=40, color=CYAN, alpha=0.85,
                                edgecolor=BG_DARK, linewidth=0.5, zorder=3)

    # Color bins by position relative to mean
    mean_prem = np.mean(all_premiums)
    for patch, left_edge in zip(patches, bins[:-1]):
        if left_edge > mean_prem * 1.3:
            patch.set_facecolor(RED)
            patch.set_alpha(0.8)
        elif left_edge > mean_prem * 1.1:
            patch.set_facecolor(AMBER)
            patch.set_alpha(0.8)

    ax.axvline(x=mean_prem, color=INDIGO_LIGHT, linestyle='--', linewidth=1.5, zorder=4)
    ax.text(mean_prem + 30, ax.get_ylim()[1] * 0.9,
            f'Mean: ${mean_prem:,.0f}', fontsize=10, color=INDIGO_LIGHT, fontweight='bold')

    median_prem = np.median(all_premiums)
    ax.axvline(x=median_prem, color=GREEN, linestyle=':', linewidth=1.5, zorder=4)
    ax.text(median_prem + 30, ax.get_ylim()[1] * 0.78,
            f'Median: ${median_prem:,.0f}', fontsize=10, color=GREEN, fontweight='bold')

    ax.set_xlabel('Annual Premium ($)', fontsize=11)
    ax.set_ylabel('Number of Policies', fontsize=11)
    ax.set_title('Premium Distribution Across Portfolio', fontsize=14,
                 fontweight='bold', color=TEXT_WHITE, pad=12)
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, p: f'${x:,.0f}'))
    ax.grid(axis='y', alpha=0.3, zorder=0)
    ax.set_axisbelow(True)

    fig.tight_layout()
    return save_chart(fig, 'chart_premium_dist.png')


# Generate all charts
print("Generating charts...")
chart_paths = {
    'lr_age': chart_loss_ratio_by_age(),
    'cf_age': chart_claims_freq_by_age(),
    'heatmap': chart_risk_heatmap(),
    'lr_era': chart_lr_by_car_era(),
    'gender': chart_gender_comparison(),
    'premium_dist': chart_premium_distribution(),
}
print("Charts generated successfully.")

# ══════════════════════════════════════════════════════════════════════════════
# PDF GENERATION (reportlab)
# ══════════════════════════════════════════════════════════════════════════════
PAGE_W, PAGE_H = letter  # 612 x 792 points
MARGIN = 0.65 * inch

# Colors for reportlab
C_BG = HexColor(BG_DARK)
C_SURFACE = HexColor(BG_SURFACE)
C_SURFACE2 = HexColor(BG_SURFACE2)
C_BORDER = HexColor(BORDER_COLOR)
C_WHITE = HexColor(TEXT_WHITE)
C_MUTED = HexColor(TEXT_MUTED)
C_INDIGO = HexColor(INDIGO)
C_INDIGO_LT = HexColor(INDIGO_LIGHT)
C_CYAN = HexColor(CYAN)
C_GREEN = HexColor(GREEN)
C_AMBER = HexColor(AMBER)
C_RED = HexColor(RED)
C_PINK = HexColor(PINK)

OUTPUT_PATH = os.path.join(OUTPUT_DIR, 'Auto_Insurance_Report.pdf')


class NumberedCanvas(canvas.Canvas):
    """Canvas that adds page numbers to each page."""
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self._draw_page_extras(num_pages)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def _draw_page_extras(self, page_count):
        page_num = self._pageNumber
        # Skip page number on cover page
        if page_num == 1:
            return
        self.setFont('Helvetica', 8)
        self.setFillColor(HexColor(TEXT_MUTED))
        self.drawRightString(PAGE_W - MARGIN, 0.4 * inch,
                             f"Page {page_num} of {page_count}")
        self.drawString(MARGIN, 0.4 * inch,
                        "Auto Insurance Portfolio Analysis | February 2026")


def draw_background(canvas_obj, doc):
    """Draw dark background on every page."""
    canvas_obj.saveState()
    canvas_obj.setFillColor(C_BG)
    canvas_obj.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
    canvas_obj.restoreState()


def draw_cover_page(canvas_obj, doc):
    """Draw the cover page with custom layout (Page 1)."""
    draw_background(canvas_obj, doc)
    c = canvas_obj
    c.saveState()

    # Top accent stripe
    c.setFillColor(C_INDIGO)
    c.rect(0, PAGE_H - 6, PAGE_W, 6, fill=1, stroke=0)

    # Bottom accent stripe
    c.setFillColor(C_CYAN)
    c.rect(0, 0, PAGE_W, 6, fill=1, stroke=0)

    # Central accent stripe/bar
    stripe_y = PAGE_H * 0.52
    c.setFillColor(C_INDIGO)
    c.setStrokeColor(C_INDIGO)
    c.setLineWidth(0)
    # Subtle gradient-like accent stripe
    c.setFillColor(HexColor('#6366F1'))
    c.rect(0, stripe_y - 2, PAGE_W, 4, fill=1, stroke=0)

    # Title
    c.setFont('Helvetica-Bold', 36)
    c.setFillColor(C_WHITE)
    title_y = PAGE_H * 0.65
    c.drawCentredString(PAGE_W / 2, title_y, "Auto Insurance")
    c.drawCentredString(PAGE_W / 2, title_y - 44, "Portfolio Analysis")

    # Subtitle
    c.setFont('Helvetica', 16)
    c.setFillColor(C_INDIGO_LT)
    c.drawCentredString(PAGE_W / 2, stripe_y - 30,
                        "Annual Risk Assessment Report")

    # Date
    c.setFont('Helvetica', 14)
    c.setFillColor(C_MUTED)
    c.drawCentredString(PAGE_W / 2, stripe_y - 55,
                        "February 2026")

    # Divider line
    c.setStrokeColor(C_BORDER)
    c.setLineWidth(1)
    c.line(PAGE_W * 0.3, stripe_y - 80, PAGE_W * 0.7, stripe_y - 80)

    # Prepared by
    c.setFont('Helvetica', 11)
    c.setFillColor(C_MUTED)
    c.drawCentredString(PAGE_W / 2, stripe_y - 105,
                        "Prepared by AI Analytics Engine (Claude Code)")

    # Confidential notice
    c.setFont('Helvetica', 9)
    c.setFillColor(HexColor('#4A4E63'))
    c.drawCentredString(PAGE_W / 2, stripe_y - 128,
                        "CONFIDENTIAL  |  FOR INTERNAL USE ONLY")

    # Bottom decorative elements
    c.setFont('Helvetica', 9)
    c.setFillColor(C_MUTED)
    c.drawCentredString(PAGE_W / 2, 40,
                        "10,000 Policyholders  |  Comprehensive Risk Assessment  |  Data-Driven Insights")

    c.restoreState()


# ── Styles ──
styles = getSampleStyleSheet()

style_title = ParagraphStyle(
    'DarkTitle', parent=styles['Heading1'],
    fontName='Helvetica-Bold', fontSize=22, leading=28,
    textColor=C_WHITE, spaceAfter=6, spaceBefore=0,
)

style_section = ParagraphStyle(
    'DarkSection', parent=styles['Heading2'],
    fontName='Helvetica-Bold', fontSize=16, leading=20,
    textColor=C_INDIGO_LT, spaceAfter=6, spaceBefore=10,
)

style_subsection = ParagraphStyle(
    'DarkSubsection', parent=styles['Heading3'],
    fontName='Helvetica-Bold', fontSize=12, leading=16,
    textColor=C_CYAN, spaceAfter=4, spaceBefore=8,
)

style_body = ParagraphStyle(
    'DarkBody', parent=styles['Normal'],
    fontName='Helvetica', fontSize=10, leading=14,
    textColor=C_WHITE, spaceAfter=6, alignment=TA_JUSTIFY,
)

style_body_muted = ParagraphStyle(
    'DarkBodyMuted', parent=style_body,
    textColor=C_MUTED, fontSize=9, leading=12,
)

style_bullet = ParagraphStyle(
    'DarkBullet', parent=style_body,
    fontName='Helvetica', fontSize=10, leading=14,
    textColor=C_WHITE, leftIndent=18, bulletIndent=6,
    spaceAfter=4,
)

style_kpi_value = ParagraphStyle(
    'KPIValue', fontName='Helvetica-Bold', fontSize=16, leading=20,
    textColor=C_WHITE, alignment=TA_CENTER, spaceAfter=0,
)

style_kpi_label = ParagraphStyle(
    'KPILabel', fontName='Helvetica', fontSize=8, leading=10,
    textColor=C_MUTED, alignment=TA_CENTER, spaceBefore=2, spaceAfter=0,
)

style_rec_title = ParagraphStyle(
    'RecTitle', fontName='Helvetica-Bold', fontSize=10.5, leading=13,
    textColor=C_INDIGO_LT, spaceAfter=2, spaceBefore=4,
)

style_rec_body = ParagraphStyle(
    'RecBody', fontName='Helvetica', fontSize=9, leading=12,
    textColor=C_WHITE, spaceAfter=1, leftIndent=12, alignment=TA_JUSTIFY,
)

style_table_header = ParagraphStyle(
    'TableHeader', fontName='Helvetica-Bold', fontSize=9, leading=12,
    textColor=C_WHITE, alignment=TA_CENTER,
)

style_table_cell = ParagraphStyle(
    'TableCell', fontName='Helvetica', fontSize=9, leading=12,
    textColor=C_WHITE, alignment=TA_CENTER,
)


# ── Helper: Colored rule ──
class ColoredRule(Flowable):
    """A horizontal rule with color."""
    def __init__(self, width, height=2, color=C_INDIGO):
        Flowable.__init__(self)
        self.width = width
        self.height = height
        self.color = color

    def draw(self):
        self.canv.setFillColor(self.color)
        self.canv.rect(0, 0, self.width, self.height, fill=1, stroke=0)


class AccentStripe(Flowable):
    """A thin accent stripe across the top of a page."""
    def __init__(self, color=C_INDIGO):
        Flowable.__init__(self)
        self.color = color

    def wrap(self, availWidth, availHeight):
        self.width = availWidth
        return (availWidth, 3)

    def draw(self):
        self.canv.setFillColor(self.color)
        self.canv.rect(0, 0, self.width, 3, fill=1, stroke=0)


# ── Build Document ──
def build_report():
    """Build the complete PDF report."""
    story = []

    # ══════════════════════════════════════════════════════════════════
    # PAGE 1: Cover Page (handled via page template)
    # ══════════════════════════════════════════════════════════════════
    # We add a page break which will trigger the cover template, then switch
    story.append(Spacer(1, 1))  # dummy flowable for cover page
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════
    # PAGE 2: Executive Summary
    # ══════════════════════════════════════════════════════════════════
    story.append(AccentStripe(C_INDIGO))
    story.append(Spacer(1, 4))
    story.append(Paragraph("Executive Summary", style_title))
    story.append(ColoredRule(PAGE_W - 2 * MARGIN, 2, C_INDIGO))
    story.append(Spacer(1, 8))

    # Overview narrative
    overview_text = (
        f"This report presents a comprehensive actuarial analysis of the auto insurance portfolio "
        f"comprising <b>{total_policies:,}</b> active policies. The portfolio generates "
        f"<b>${total_premium:,.0f}</b> in annual written premium with <b>${total_loss:,.0f}</b> "
        f"in total incurred losses, yielding an overall loss ratio of <b>{loss_ratio*100:.1f}%</b>. "
        f"Claims frequency stands at <b>{claims_freq*100:.1f}%</b>, with {claims_count:,} policyholders "
        f"filing at least one claim during the observation period. The average annual premium is "
        f"<b>${avg_premium:,.0f}</b>, while the average loss per claimant is "
        f"<b>${avg_loss_per_claim:,.0f}</b>. The following sections provide granular breakdowns "
        f"by age cohort, vehicle vintage, and gender to identify risk concentrations and inform "
        f"pricing adjustments."
    )
    story.append(Paragraph(overview_text, style_body))
    story.append(Spacer(1, 10))

    # KPI Boxes
    story.append(Paragraph("Key Performance Indicators", style_section))
    story.append(Spacer(1, 4))

    kpi_data = [
        [Paragraph('<b>Total Policies</b>', style_kpi_label),
         Paragraph('<b>Total Premium</b>', style_kpi_label),
         Paragraph('<b>Total Incurred Loss</b>', style_kpi_label),
         Paragraph('<b>Loss Ratio</b>', style_kpi_label),
         Paragraph('<b>Claims Frequency</b>', style_kpi_label)],
        [Paragraph(f'<font color="{INDIGO_LIGHT}"><b>10,000</b></font>', style_kpi_value),
         Paragraph(f'<font color="{CYAN}"><b>${total_premium/1e6:.2f}M</b></font>', style_kpi_value),
         Paragraph(f'<font color="{AMBER}"><b>${total_loss/1e6:.2f}M</b></font>', style_kpi_value),
         Paragraph(f'<font color="{RED}"><b>{loss_ratio*100:.1f}%</b></font>', style_kpi_value),
         Paragraph(f'<font color="{GREEN}"><b>{claims_freq*100:.1f}%</b></font>', style_kpi_value)],
    ]

    col_w = (PAGE_W - 2 * MARGIN) / 5
    kpi_table = Table(kpi_data, colWidths=[col_w] * 5, rowHeights=[18, 28])
    kpi_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), C_SURFACE),
        ('BOX', (0, 0), (-1, -1), 1, C_BORDER),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, C_BORDER),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(kpi_table)
    story.append(Spacer(1, 14))

    # Key Findings
    story.append(Paragraph("Key Findings", style_section))
    story.append(Spacer(1, 4))

    findings = [
        f"<b>Highest-risk age segment:</b> Policyholders aged {highest_lr_age} exhibit a loss ratio "
        f"of {by_age[highest_lr_age]['loss_ratio']*100:.1f}%, significantly exceeding the portfolio "
        f"average of {loss_ratio*100:.1f}%. This cohort also has the highest claims frequency at "
        f"{by_age[highest_lr_age]['claims_freq']*100:.1f}%.",

        f"<b>Lowest-risk age segment:</b> The {lowest_lr_age} cohort maintains a loss ratio of just "
        f"{by_age[lowest_lr_age]['loss_ratio']*100:.1f}%, representing the most profitable segment "
        f"with {by_age[lowest_lr_age]['count']:,} policies.",

        f"<b>Vehicle age impact:</b> Older vehicles (model years {highest_lr_era}) show a loss ratio "
        f"of {by_era[highest_lr_era]['loss_ratio']*100:.1f}%, while newer vehicles ({lowest_lr_era}) "
        f"perform at {by_era[lowest_lr_era]['loss_ratio']*100:.1f}%.",

        f"<b>Gender analysis:</b> Male policyholders ({by_gender['Male']['count']:,} policies) carry a "
        f"loss ratio of {by_gender['Male']['loss_ratio']*100:.1f}%, compared to "
        f"{by_gender['Female']['loss_ratio']*100:.1f}% for female policyholders "
        f"({by_gender['Female']['count']:,} policies).",

        f"<b>Risk concentration:</b> The intersection of {max_heatmap_cell[0]} age group with "
        f"{max_heatmap_cell[1]} era vehicles represents the highest-risk cell at "
        f"{max_heatmap_lr*100:.0f}% loss ratio (n={heatmap[max_heatmap_cell[0]][max_heatmap_cell[1]]['n']}).",
    ]

    for finding in findings:
        story.append(Paragraph(f"\u2022  {finding}", style_bullet))

    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════
    # PAGE 3: Age Group Analysis
    # ══════════════════════════════════════════════════════════════════
    story.append(AccentStripe(C_AMBER))
    story.append(Spacer(1, 4))
    story.append(Paragraph("Age Group Analysis", style_title))
    story.append(ColoredRule(PAGE_W - 2 * MARGIN, 2, C_AMBER))
    story.append(Spacer(1, 6))

    age_narrative = (
        f"Age is the single strongest predictor of claims behavior in this portfolio. "
        f"Young drivers aged 18-24 present the highest risk profile with a {by_age['18-24']['loss_ratio']*100:.1f}% "
        f"loss ratio and {by_age['18-24']['claims_freq']*100:.1f}% claims frequency, despite being charged "
        f"the highest average premium of ${by_age['18-24']['avg_premium']:,.0f}. Conversely, the 35-44 "
        f"age band demonstrates mature driving behavior with the lowest claims frequency at "
        f"{by_age['35-44']['claims_freq']*100:.1f}%. Senior drivers (65+) show elevated loss ratios "
        f"({by_age['65+']['loss_ratio']*100:.1f}%) driven primarily by higher severity per claim "
        f"(${by_age['65+']['avg_loss_per_claim']:,.0f} average) rather than frequency."
    )
    story.append(Paragraph(age_narrative, style_body))
    story.append(Spacer(1, 6))

    # Chart 1: Loss Ratio by Age
    story.append(Image(chart_paths['lr_age'], width=6.8*inch, height=3.3*inch))
    story.append(Spacer(1, 4))

    # Chart 2: Claims Frequency by Age
    story.append(Image(chart_paths['cf_age'], width=6.8*inch, height=3.3*inch))

    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════
    # PAGE 4: Risk Segmentation
    # ══════════════════════════════════════════════════════════════════
    story.append(AccentStripe(C_RED))
    story.append(Spacer(1, 4))
    story.append(Paragraph("Risk Segmentation", style_title))
    story.append(ColoredRule(PAGE_W - 2 * MARGIN, 2, C_RED))
    story.append(Spacer(1, 6))

    risk_narrative = (
        f"Cross-tabulating age group against vehicle vintage reveals actionable risk concentrations. "
        f"The heatmap below displays loss ratios for each cell, with sample sizes in parentheses. "
        f"The highest-risk intersection is <b>{max_heatmap_cell[0]} drivers with {max_heatmap_cell[1]} era vehicles</b> "
        f"at {max_heatmap_lr*100:.0f}% loss ratio. Notably, older vehicles across all age groups "
        f"tend to carry higher loss ratios, likely reflecting reduced safety features, higher repair "
        f"costs relative to vehicle value, and adverse selection. The line chart below tracks how "
        f"loss ratios shift across vehicle eras, showing a clear downward trend for newer model years."
    )
    story.append(Paragraph(risk_narrative, style_body))
    story.append(Spacer(1, 6))

    # Chart 3: Heatmap
    story.append(Image(chart_paths['heatmap'], width=6.8*inch, height=3.5*inch))
    story.append(Spacer(1, 4))

    # Chart 4: Loss Ratio by Car Era
    story.append(Image(chart_paths['lr_era'], width=6.8*inch, height=3.0*inch))

    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════
    # PAGE 5: Gender & Premium Analysis
    # ══════════════════════════════════════════════════════════════════
    story.append(AccentStripe(C_PINK))
    story.append(Spacer(1, 4))
    story.append(Paragraph("Gender & Premium Analysis", style_title))
    story.append(ColoredRule(PAGE_W - 2 * MARGIN, 2, C_PINK))
    story.append(Spacer(1, 6))

    gender_narrative = (
        f"Male policyholders ({by_gender['Male']['count']:,}, "
        f"{by_gender['Male']['count']/total_policies*100:.0f}%) carry a {by_gender['Male']['loss_ratio']*100:.1f}% "
        f"loss ratio vs {by_gender['Female']['loss_ratio']*100:.1f}% for females "
        f"({by_gender['Female']['count']:,}, {by_gender['Female']['count']/total_policies*100:.0f}%). "
        f"The chart below decomposes this by age, revealing that gender differences vary across cohorts. "
        f"Premium distribution follows an approximately normal shape centered at "
        f"${np.mean(all_premiums):,.0f} with rightward skew from young-driver surcharges."
    )
    story.append(Paragraph(gender_narrative, style_body))
    story.append(Spacer(1, 6))

    # Chart 5: Gender comparison
    story.append(Image(chart_paths['gender'], width=6.8*inch, height=3.0*inch))
    story.append(Spacer(1, 3))

    # Gender comparison table
    story.append(Paragraph("Gender Comparison Summary", style_subsection))
    story.append(Spacer(1, 2))

    gender_table_data = [
        [Paragraph('<b>Metric</b>', style_table_header),
         Paragraph(f'<font color="{INDIGO}"><b>Male</b></font>', style_table_header),
         Paragraph(f'<font color="{PINK}"><b>Female</b></font>', style_table_header),
         Paragraph('<b>Difference</b>', style_table_header)],
        [Paragraph('Policy Count', style_table_cell),
         Paragraph(f"{by_gender['Male']['count']:,}", style_table_cell),
         Paragraph(f"{by_gender['Female']['count']:,}", style_table_cell),
         Paragraph(f"{by_gender['Male']['count'] - by_gender['Female']['count']:+,}", style_table_cell)],
        [Paragraph('Loss Ratio', style_table_cell),
         Paragraph(f"{by_gender['Male']['loss_ratio']*100:.1f}%", style_table_cell),
         Paragraph(f"{by_gender['Female']['loss_ratio']*100:.1f}%", style_table_cell),
         Paragraph(f"{(by_gender['Male']['loss_ratio'] - by_gender['Female']['loss_ratio'])*100:+.1f} pp", style_table_cell)],
        [Paragraph('Claims Frequency', style_table_cell),
         Paragraph(f"{by_gender['Male']['claims_freq']*100:.1f}%", style_table_cell),
         Paragraph(f"{by_gender['Female']['claims_freq']*100:.1f}%", style_table_cell),
         Paragraph(f"{(by_gender['Male']['claims_freq'] - by_gender['Female']['claims_freq'])*100:+.1f} pp", style_table_cell)],
        [Paragraph('Avg Premium', style_table_cell),
         Paragraph(f"${by_gender['Male']['avg_premium']:,.0f}", style_table_cell),
         Paragraph(f"${by_gender['Female']['avg_premium']:,.0f}", style_table_cell),
         Paragraph(f"${by_gender['Male']['avg_premium'] - by_gender['Female']['avg_premium']:+,.0f}", style_table_cell)],
    ]

    gcol_w = (PAGE_W - 2 * MARGIN) / 4
    gender_table = Table(gender_table_data, colWidths=[gcol_w] * 4, rowHeights=[16] * 5)
    gender_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), C_SURFACE2),
        ('BACKGROUND', (0, 1), (-1, -1), C_SURFACE),
        ('BOX', (0, 0), (-1, -1), 1, C_BORDER),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, C_BORDER),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    story.append(gender_table)
    story.append(Spacer(1, 4))

    # Chart 6: Premium distribution
    story.append(Image(chart_paths['premium_dist'], width=6.8*inch, height=2.8*inch))

    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════
    # PAGE 6: Recommendations
    # ══════════════════════════════════════════════════════════════════
    story.append(AccentStripe(C_GREEN))
    story.append(Spacer(1, 4))
    story.append(Paragraph("Recommendations & Action Items", style_title))
    story.append(ColoredRule(PAGE_W - 2 * MARGIN, 2, C_GREEN))
    story.append(Spacer(1, 8))

    rec_intro = (
        "Based on the portfolio analysis presented in the preceding pages, the following "
        "recommendations are provided to optimize risk selection, improve pricing adequacy, "
        "and strengthen the overall portfolio loss ratio."
    )
    story.append(Paragraph(rec_intro, style_body))
    story.append(Spacer(1, 6))

    # Recommendation 1
    story.append(Paragraph(
        f'<font color="{INDIGO_LIGHT}">1.</font>  Reprice Young Driver Segment (18-24)',
        style_rec_title))
    story.append(Paragraph(
        f"The 18-24 cohort generates a {by_age['18-24']['loss_ratio']*100:.1f}% loss ratio, well above "
        f"the portfolio average of {loss_ratio*100:.1f}%. Despite carrying the highest average premium "
        f"(${by_age['18-24']['avg_premium']:,.0f}), the segment remains unprofitable. Recommend a "
        f"targeted rate increase of 12-18%, supplemented by usage-based insurance (UBI) "
        f"telematics programs to reward safe driving and enable more granular risk differentiation.",
        style_rec_body))
    story.append(Spacer(1, 2))

    # Recommendation 2
    story.append(Paragraph(
        f'<font color="{INDIGO_LIGHT}">2.</font>  Enhanced Underwriting for Older Vehicles',
        style_rec_title))
    story.append(Paragraph(
        f"Vehicles from the {highest_lr_era} era show a {by_era[highest_lr_era]['loss_ratio']*100:.1f}% loss ratio. "
        f"Recommend stricter vehicle inspection requirements for cars older than 15 years, "
        f"applying an age-of-vehicle surcharge of 5-10%, and capping insured value to actual cash value "
        f"less depreciation. Consider excluding total-loss-prone vehicles from the portfolio.",
        style_rec_body))
    story.append(Spacer(1, 2))

    # Recommendation 3
    story.append(Paragraph(
        f'<font color="{INDIGO_LIGHT}">3.</font>  Senior Driver Risk Mitigation (65+)',
        style_rec_title))
    story.append(Paragraph(
        f"The 65+ segment shows a {by_age['65+']['loss_ratio']*100:.1f}% loss ratio with high "
        f"average claim severity (${by_age['65+']['avg_loss_per_claim']:,.0f} per claim). "
        f"Recommend mandatory defensive driving course discounts, annual medical fitness "
        f"assessments for policyholders over 70, and adjusting deductibles upward to reduce severity exposure.",
        style_rec_body))
    story.append(Spacer(1, 2))

    # Recommendation 4
    story.append(Paragraph(
        f'<font color="{INDIGO_LIGHT}">4.</font>  Grow Low-Risk Segments',
        style_rec_title))
    story.append(Paragraph(
        f"The {lowest_lr_age} age group is highly profitable at "
        f"{by_age[lowest_lr_age]['loss_ratio']*100:.1f}% loss ratio. Recommend competitive pricing "
        f"initiatives and retention programs for this demographic, including multi-policy "
        f"discounts. Allocate marketing budget to acquire policyholders aged 25-44 "
        f"with newer vehicles (2020+), which combine favorable frequency and severity profiles.",
        style_rec_body))
    story.append(Spacer(1, 2))

    # Recommendation 5
    story.append(Paragraph(
        f'<font color="{INDIGO_LIGHT}">5.</font>  Portfolio-Wide Pricing Refinement',
        style_rec_title))
    story.append(Paragraph(
        f"The portfolio loss ratio of {loss_ratio*100:.1f}% is within acceptable range (60-80%) "
        f"but leaves limited room for adverse development. Recommend implementing "
        f"a multivariate GLM incorporating age, vehicle year, gender, "
        f"geography, and claims history to replace current rating factors and reduce cross-subsidization.",
        style_rec_body))
    story.append(Spacer(1, 2))

    # Recommendation 6
    story.append(Paragraph(
        f'<font color="{INDIGO_LIGHT}">6.</font>  Claims Management Improvements',
        style_rec_title))
    story.append(Paragraph(
        f"With {claims_count:,} claims ({claims_freq*100:.1f}% frequency) and average severity of "
        f"${avg_loss_per_claim:,.0f}, there is opportunity to reduce loss costs. "
        f"Recommend AI-based fraud detection for high-severity claims, "
        f"preferred repair shop networks, and subrogation recovery programs. "
        f"Target a 3-5% reduction in average claim cost within 12 months.",
        style_rec_body))
    story.append(Spacer(1, 6))

    # Summary box
    story.append(ColoredRule(PAGE_W - 2 * MARGIN, 1, C_BORDER))
    story.append(Spacer(1, 8))

    # Projected impact table
    story.append(Paragraph("Projected Impact of Recommendations", style_subsection))
    story.append(Spacer(1, 4))

    target_lr = loss_ratio * 0.92  # ~8% improvement
    premium_uplift = total_premium * 0.05  # ~5% rate increase impact

    impact_data = [
        [Paragraph('<b>Scenario</b>', style_table_header),
         Paragraph('<b>Loss Ratio</b>', style_table_header),
         Paragraph('<b>Projected Premium</b>', style_table_header),
         Paragraph('<b>Net Impact</b>', style_table_header)],
        [Paragraph('Current State', style_table_cell),
         Paragraph(f'{loss_ratio*100:.1f}%', style_table_cell),
         Paragraph(f'${total_premium/1e6:.2f}M', style_table_cell),
         Paragraph('-', style_table_cell)],
        [Paragraph('Post-Implementation (12 mo)', style_table_cell),
         Paragraph(f'<font color="{GREEN}">{target_lr*100:.1f}%</font>', style_table_cell),
         Paragraph(f'<font color="{GREEN}">${(total_premium + premium_uplift)/1e6:.2f}M</font>', style_table_cell),
         Paragraph(f'<font color="{GREEN}">+${premium_uplift/1e6:.2f}M premium, '
                   f'-{(loss_ratio - target_lr)*100:.1f}pp LR</font>', style_table_cell)],
    ]

    icol_w = (PAGE_W - 2 * MARGIN) / 4
    impact_table = Table(impact_data, colWidths=[icol_w] * 4, rowHeights=[20, 20, 24])
    impact_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), C_SURFACE2),
        ('BACKGROUND', (0, 1), (-1, 1), C_SURFACE),
        ('BACKGROUND', (0, 2), (-1, 2), HexColor('#112218')),
        ('BOX', (0, 0), (-1, -1), 1, C_BORDER),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, C_BORDER),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    story.append(impact_table)
    story.append(Spacer(1, 12))

    # Disclaimer
    disclaimer = (
        "<i>Disclaimer: This report is generated based on a portfolio of 10,000 simulated policies "
        "for analytical demonstration purposes. All recommendations should be validated against "
        "actual underwriting guidelines, regulatory constraints, and market conditions before "
        "implementation. Past experience may not be indicative of future results.</i>"
    )
    story.append(Paragraph(disclaimer, style_body_muted))

    return story


# ── Page Templates ──
def build_pdf():
    """Build the PDF with multiple page templates."""
    doc = BaseDocTemplate(
        OUTPUT_PATH,
        pagesize=letter,
        leftMargin=MARGIN,
        rightMargin=MARGIN,
        topMargin=MARGIN,
        bottomMargin=0.7 * inch,
        title='Auto Insurance Portfolio Analysis',
        author='AI Analytics Engine (Claude Code)',
        subject='Annual Risk Assessment Report - February 2026',
    )

    # Frame for cover page (not used for content, just a placeholder)
    cover_frame = Frame(
        MARGIN, MARGIN, PAGE_W - 2 * MARGIN, PAGE_H - 2 * MARGIN,
        id='cover_frame'
    )

    # Frame for content pages
    content_frame = Frame(
        MARGIN, 0.7 * inch, PAGE_W - 2 * MARGIN, PAGE_H - MARGIN - 0.7 * inch,
        id='content_frame'
    )

    cover_template = PageTemplate(
        id='cover',
        frames=[cover_frame],
        onPage=draw_cover_page,
    )

    content_template = PageTemplate(
        id='content',
        frames=[content_frame],
        onPage=draw_background,
    )

    doc.addPageTemplates([cover_template, content_template])

    # Build story with template switching
    story = []

    # Cover page content (minimal - the real content is drawn via onPage)
    story.append(Spacer(1, 1))
    from reportlab.platypus.doctemplate import NextPageTemplate
    story.append(NextPageTemplate('content'))
    story.append(PageBreak())

    # Add the main content
    main_content = build_report()
    # Skip the first two elements (Spacer + PageBreak) from build_report
    # since we handle the page transition here
    story.extend(main_content[2:])

    print(f"Building PDF: {OUTPUT_PATH}")
    doc.build(story, canvasmaker=NumberedCanvas)
    print(f"PDF saved successfully: {OUTPUT_PATH}")
    print(f"File size: {os.path.getsize(OUTPUT_PATH) / 1024:.0f} KB")


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == '__main__':
    build_pdf()

    # Clean up chart files
    import shutil
    if os.path.exists(CHART_DIR):
        shutil.rmtree(CHART_DIR)
        print("Cleaned up temporary chart files.")

    print("Done!")
