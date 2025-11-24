import os

# ==========================================
# File Paths (Sample)
# ==========================================
# 実際にはここにあなたの経歴書ファイル名を指定します
INPUT_FILE = 'Resume_Template.xlsx'
OUTPUT_DIR = 'output'
OUTPUT_FILE = os.path.join(OUTPUT_DIR, 'resume_master.json')

# ==========================================
# User Settings (Sample)
# ==========================================
# 年齢計算の基準日など
BIRTH_DATE = "1990-01-01"
CAREER_START_DATE = "2013-04-01"

# ==========================================
# System Settings
# ==========================================
# シート名など
TARGET_SHEET_NAME = 'SkillSheet'
TEMPLATE_SHEET_NAME = '_Template'