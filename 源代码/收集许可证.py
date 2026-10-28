from pathlib import Path
import importlib.metadata as metadata
import shutil
root = Path(__file__).resolve().parent.parent / '第三方许可'
for dist in metadata.distributions():
    name = dist.metadata.get('Name', 'unknown')
    for item in dist.files or []:
        text = str(item).replace('\\', '/')
        if '.dist-info/' not in text:
            continue
        base = Path(text).name.lower()
        if not ('license' in text.lower() or base in ('copying', 'copying.txt', 'notice', 'notice.txt', 'authors')):
            continue
        source = dist.locate_file(item)
        if source.is_file():
            dest = root / name / Path(*Path(text).parts[1:])
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, dest)
