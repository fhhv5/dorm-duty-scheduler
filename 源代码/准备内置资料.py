"""Prepare the resources bundled by the PyInstaller spec before building."""
from pathlib import Path
import shutil

root = Path(__file__).resolve().parent.parent
dest = root / '构建临时' / '内置资料'
dest.mkdir(parents=True, exist_ok=True)
for folder in ('模板', '第三方许可'):
    shutil.copytree(root / folder, dest / folder, dirs_exist_ok=True,
                    ignore=shutil.ignore_patterns('*.inspect.ndjson', '~$*'))
for file in ('使用说明.txt', '免责声明.txt', 'LICENSE'):
    shutil.copy2(root / file, dest / file)
print(f'已准备内置资料：{dest}。打包前请确认模板只包含可公开的示例数据。')
