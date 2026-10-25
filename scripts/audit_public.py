"""发布前检查源码白名单、个人信息与 Office 包内元数据。"""  # 不读取论文原件或用户默认模板。
import json  # 仅输出检查结论。
import re  # 检查本机绝对路径和凭据样式。
import subprocess  # 发布审核依据实际 Git 跟踪文件。
import sys  # 通过退出码阻止错误发布。
import zipfile  # 检查模板元数据和外部关系。
from pathlib import Path  # 以脚本位置确定审核边界。
from xml.etree import ElementTree as ET  # 不引入额外 Python 包。

ROOT = Path(__file__).resolve().parents[1]  # 唯一允许审核的项目。
FORBIDDEN_SUFFIXES = {'.ttf', '.otf', '.dotm', '.doc', '.pfx', '.pem', '.key'}  # 这些材料不属于公共项目。
ALLOWED_TOP = {'src', 'tests', 'assets', 'docs', 'scripts', '.github'}  # 子目录按白名单发布。
ALLOWED_ROOT = {'README.md', 'LICENSE', 'CHANGELOG.md', 'THIRD-PARTY-NOTICES.md', 'Directory.Build.props', 'global.json', '.gitignore', '.gitattributes'}  # 根目录只发布必要文件。

def audit():  # 校验实际版本库，而不是任意整个工作区。
    listed = subprocess.run(['git', 'ls-files', '-z'], cwd=ROOT, capture_output=True, check=True).stdout.decode('utf-8').split('\0')  # 文件名保留中文。
    failures = []  # 输出规则编号和相对文件名，不输出被匹配的敏感内容。
    count = 0  # 记录实际审核数量。
    for relative in filter(None, listed):  # 跳过结束空项。
        path = ROOT / relative  # 禁止从仓库之外读取内容。
        parts = Path(relative).parts  # 检查顶层路径。
        if parts[0] not in ALLOWED_TOP and relative not in ALLOWED_ROOT:
            failures.append((relative, 'not-in-public-allowlist'))
        if path.suffix.lower() in FORBIDDEN_SUFFIXES or 'Normal.dotm' in path.name or any(part in ('bin', 'obj', 'qa', '_qa', '__pycache__') for part in parts):
            failures.append((relative, 'private-or-build-artifact'))
        if path.suffix.lower() in ('.docx', '.dotx'):
            with zipfile.ZipFile(path) as archive:  # 读取公开包，不修改。
                for name in archive.namelist():
                    if name.startswith('word/fonts/'):
                        failures.append((relative, 'embedded-font'))
                    if name.endswith('.rels') and b'TargetMode="External"' in archive.read(name):
                        failures.append((relative, 'external-office-relationship'))
                core = ET.fromstring(archive.read('docProps/core.xml'))
                for element in core:
                    if element.tag.rsplit('}', 1)[-1] in ('creator', 'lastModifiedBy') and (element.text or '').strip():
                        failures.append((relative, 'personal-author-metadata'))
        elif path.suffix.lower() not in ('.png', '.jpg', '.jpeg', '.ico'):
            text = path.read_text(encoding='utf-8-sig')  # 源码和文档应当使用 UTF-8。
            if re.search(r'[A-Za-z]:[\\/](?:Users|ZZDX-NSCC|郑州大学)', text, re.I):
                failures.append((relative, 'personal-absolute-path'))
            if re.search(r'gh[pousr]_[A-Za-z0-9]{25,}|github_pat_[A-Za-z0-9_]{30,}', text):
                failures.append((relative, 'credential-pattern'))
        count += 1
    return {'passed': not failures, 'files_checked': count, 'failures': failures}  # 安全的公共审核摘要。

if __name__ == '__main__':
    result = audit()  # 只检查，不自动修改仓库。
    print(json.dumps(result, ensure_ascii=False, indent=2))
    sys.exit(0 if result['passed'] else 1)
