# -*- mode: python -*-

block_cipher = None


a = Analysis(['__init__.py', 'extractors.py', 'extract_benchmark.py', 'extract_parcc.py', 'usage_report.py'],
             pathex=['C:\\Users\\Malcolm\\Documents\\Atom Projects\\Extractor Hub'],
             binaries=[],
             datas=[],
             hiddenimports=['pandas._libs.tslibs.timedeltas'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='Extractor Hub',
          debug=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False)
