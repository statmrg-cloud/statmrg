# -*- mode: python ; coding: utf-8 -*-
"""
전자책 자동 생성기 - PyInstaller 빌드 스펙
Windows 10/11 64비트 타겟
빌드 명령: pyinstaller ebook_creator.spec
"""

import os

block_cipher = None

# 포함할 데이터 파일 (소스 → 번들 내 경로)
added_files = [
    ('templates',   'templates'),
    ('static',      'static'),
    ('modules',     'modules'),
    ('config.py',   '.'),
]

a = Analysis(
    ['launcher.py'],
    pathex=['.'],
    binaries=[],
    datas=added_files,
    hiddenimports=[
        # Flask 관련
        'flask',
        'flask.templating',
        'flask_cors',
        'jinja2',
        'jinja2.ext',
        'werkzeug',
        'werkzeug.serving',
        'werkzeug.utils',
        'werkzeug.routing',
        'werkzeug.exceptions',
        # OpenAI
        'openai',
        'openai.types',
        'httpx',
        'httpcore',
        'anyio',
        'sniffio',
        # PDF/DOCX/PPTX
        'reportlab',
        'reportlab.pdfgen',
        'reportlab.lib',
        'reportlab.lib.pagesizes',
        'reportlab.lib.styles',
        'reportlab.lib.units',
        'reportlab.platypus',
        'reportlab.platypus.flowables',
        'reportlab.pdfbase',
        'reportlab.pdfbase.ttfonts',
        'docx',
        'docx.oxml',
        'docx.oxml.ns',
        'pptx',
        'pptx.util',
        'pptx.enum',
        'pptx.dml.color',
        # 이미지
        'PIL',
        'PIL.Image',
        'PIL.ImageDraw',
        'PIL.ImageFont',
        # 기타
        'requests',
        'certifi',
        'charset_normalizer',
        'idna',
        'urllib3',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'pytest',
        'test',
        'unittest',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='전자책생성기',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,           # 로그 확인을 위해 콘솔창 표시
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='전자책생성기',
)
