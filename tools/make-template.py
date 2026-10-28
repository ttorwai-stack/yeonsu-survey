#!/usr/bin/env python3
"""배포용 세트(dist/yeonsu-survey-template.zip)를 만든다.
운영 중인 파일에서 내 서버 주소를 비우고, 매뉴얼·README를 함께 묶는다.
사용: python3 tools/make-template.py
"""
import re, sys, zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DIST = ROOT / 'dist'
CREDIT = 'Prompted by kidsreturn_79 · © 2026.10'

def read(p): return (ROOT / p).read_text(encoding='utf-8')

def sub_once(text, pattern, repl, what):
    new, n = re.subn(pattern, repl, text)
    if n != 1:
        sys.exit('치환 실패 (%s): %d곳 발견' % (what, n))
    return new

files = {}
for name in ('index.html', 'admin.html'):
    t = read(name)
    t = sub_once(t, r"var DEFAULT_API_ID = '[^']*';", "var DEFAULT_API_ID = '';", name + ' DEFAULT_API_ID')
    files[name] = t
files['Code.gs'] = sub_once(read('server/Code.gs'), r"var PAGES_BASE = '[^']*';",
                            "var PAGES_BASE = 'https://YOUR-ID.github.io/YOUR-REPO/';", 'PAGES_BASE')
files['MANUAL.html'] = read('template/MANUAL.html')
files['README.md'] = read('template/README.md')

# 점검: 내 주소가 남아 있지 않아야 하고, 제작 표시는 모든 파일에 있어야 한다
for n, t in files.items():
    for bad in ('ttorwai', 'AKfycbyUUAH8'):
        if bad in t: sys.exit('%s 에 내 정보(%s)가 남아 있어요' % (n, bad))
    if CREDIT not in t: sys.exit('%s 에 제작 표시가 없어요' % n)

DIST.mkdir(exist_ok=True)
out = DIST / 'yeonsu-survey-template.zip'
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    for n, t in files.items(): z.writestr('yeonsu-survey-template/' + n, t)
print('만들었어요:', out)
for n, t in files.items(): print(' -', n, len(t.encode('utf-8')), 'bytes')
