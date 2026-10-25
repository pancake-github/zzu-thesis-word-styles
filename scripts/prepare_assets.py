"""清理公开文档元数据并从既有模板提取样式；不修改原件或 Word 设置。"""  # 说明用途与边界。
from __future__ import annotations  # 允许清楚的类型注解。
import argparse  # 读取明确的源文件与目标目录参数。
import hashlib  # 为公开资产生成 SHA-256。
import io  # 用内存字节流验证命名空间声明。
import json  # 输出可审查且不依赖机器路径的规格。
from pathlib import Path  # 正确处理中文文件名。
from zipfile import ZIP_DEFLATED, ZipFile  # 在不改变文档内容的情况下重打包。
from xml.etree import ElementTree as ET  # 仅使用 Python 标准库解析 OOXML。

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"  # 定义 Word 名称空间。
NS = {"w": W}  # 供 XPath 查询使用。
CORE_NAMESPACES = {"cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties", "dc": "http://purl.org/dc/elements/1.1/", "dcterms": "http://purl.org/dc/terms/", "dcmitype": "http://purl.org/dc/dcmitype/", "xsi": "http://www.w3.org/2001/XMLSchema-instance"}  # 保持 core.xml 标准前缀，尤其保护 xsi:type 中的 QName。
for prefix, uri in CORE_NAMESPACES.items():  # 注册原文档采用的公开标准名称空间。
    ET.register_namespace(prefix, uri)  # 避免自动改写成 ns0/ns1 后留下悬空的属性值前缀。
APPROVED = ["Normal", "Heading1", "Heading2", "Heading3", "Heading4", "中文摘要标题", "英文摘要标题", "TOCHeading", "图题", "表题", "附属正文", "致谢正文", "TOC1", "TOC2", "TOC3", "TableofFigures", "Header", "Footer", "英文页眉", "局部加粗", "封面信息四号", "封面信息三号", "封面题目二号", "封面校名初号", "封面学位小初", "英文封面三号"]  # 固定批准的 26 项，不导出库内其他默认样式。
USES = ["正文、附录正文、中英文摘要正文", "一级标题", "二级标题", "三级标题", "四级标题", "中文摘要标题", "英文摘要标题", "目录标题", "图题", "表题", "参考文献、符号说明、个人简历及成果", "致谢正文", "一级目录条目", "二级目录条目", "三级目录条目", "图表清单条目", "中文页眉", "页码", "英文页眉", "关键词标签等局部加粗", "封面四号文字", "封面三号文字", "封面二号论文题目", "封面校名", "封面学位名称", "英文封面三号文字"]  # 保留使用位置与样式之间的对应关系。


def q(name: str) -> str:  # 转换 Word 限定名称。
    return "{" + W + "}" + name  # 返回标准 Clark 形式。


def attrs(node: ET.Element | None) -> dict:  # 将 OOXML 属性转换为可读的键值。
    return {} if node is None else {key.rsplit("}", 1)[-1]: value for key, value in node.attrib.items()}  # 不丢失原始数值单位。


def value(node: ET.Element, path: str, default=None):  # 读取单一 w:val 属性。
    item = node.find(path, NS)  # 查询实际节点。
    return default if item is None else item.get(q("val"), "1")  # 空布尔节点按开启处理。


def enabled(node: ET.Element, path: str) -> bool:  # 读取 OOXML 开关值。
    return value(node, path, "0") not in ("0", "false", "off")  # 同时兼容布尔写法。


def property_tree(node: ET.Element | None):  # 保留完整字符和段落属性树供审核。
    if node is None:  # 允许字符样式没有段落属性。
        return None  # 用空值表示继承而非擅自补默认值。
    return {"tag": node.tag.rsplit("}", 1)[-1], "attributes": attrs(node), "children": [property_tree(child) for child in node]}  # 保留顺序与所有已写入参数。


def validate_core_qnames(data: bytes) -> None:  # 验证 XML 可解析之外的 xsi:type QName 语义。
    scopes: list[dict[str, str]] = []  # 逐元素跟踪当前有效的前缀映射。
    pending: dict[str, str] = {}  # 暂存即将作用于下一个元素的声明。
    for event, item in ET.iterparse(io.BytesIO(data), events=("start-ns", "start", "end")):  # 获取声明事件，而不把字符串 QName 误当普通文本。
        if event == "start-ns":  # 名称空间声明先于对应元素事件发生。
            pending[item[0]] = item[1]  # 保存当前声明的真实 URI。
        elif event == "start":  # 元素开始时创建独立作用域。
            active = dict(scopes[-1]) if scopes else {}  # 继承祖先声明，允许合法的子级覆盖。
            active.update(pending)  # 合并本元素新声明。
            pending.clear()  # 避免声明被错误应用到后续兄弟元素。
            scopes.append(active)  # 入栈后检查当前元素的 QName 属性。
            type_value = item.get("{" + CORE_NAMESPACES["xsi"] + "}type")  # xsi:type 的值本身也是一个限定名称。
            if type_value and ":" in type_value:  # 对带前缀的类型值检查声明。
                prefix = type_value.split(":", 1)[0]  # 提取 QName 中实际引用的前缀。
                if prefix not in active:  # Word 会拒绝没有声明的 dcterms:W3CDTF 等类型。
                    raise ValueError(f"core.xml 的 xsi:type 引用了未声明前缀：{prefix}")  # 在发布前明确阻断。
                if prefix in CORE_NAMESPACES and active[prefix] != CORE_NAMESPACES[prefix]:  # 防止标准前缀绑定到错误名称空间。
                    raise ValueError(f"core.xml 的标准前缀绑定错误：{prefix}")  # 不允许仅名称相同却语义不同。
        else:  # 元素结束时恢复祖先作用域。
            scopes.pop()  # 保持验证与 XML 名称空间生命周期一致。


def sanitize(source: Path, target: Path) -> dict:  # 清除身份元数据并验证格式与内容不变。
    if source.resolve() == target.resolve():  # 原件不得被覆盖。
        raise ValueError("公开副本必须与来源文件分开保存。")  # 返回可操作的错误。
    with ZipFile(source) as archive:  # 只读打开原件。
        parts = {info.filename: archive.read(info.filename) for info in archive.infolist()}  # 保存每个 ZIP 条目的解压内容。
        core = ET.fromstring(parts["docProps/core.xml"])  # 只编辑身份元数据文件。
        for item in list(core):  # 检查已存在的元数据项。
            if item.tag.rsplit("}", 1)[-1] in {"creator", "lastModifiedBy", "lastPrinted"}:  # 清除作者、修改者及私人打印时间。
                core.remove(item)  # 不触碰标题、日期、正文与样式。
        core.set("xmlns:dcmitype", CORE_NAMESPACES["dcmitype"])  # 保留原核心属性中的类型名称空间，虽然本文件未直接使用其元素。
        parts["docProps/core.xml"] = ET.tostring(core, encoding="utf-8", xml_declaration=True)  # 使用已注册的标准前缀序列化清理后的属性。
        validate_core_qnames(parts["docProps/core.xml"])  # 防止 XML 可解析但 Word 不接受的悬空 QName 回归。
        with ZipFile(target, "w", compression=ZIP_DEFLATED, compresslevel=9) as output:  # 只创建公开副本。
            for info in archive.infolist():  # 保留原 ZIP 条目顺序及属性。
                output.writestr(info, parts[info.filename])  # 原样写入其他所有文件内容。
    with ZipFile(source) as original, ZipFile(target) as cleaned:  # 独立回读确认修改范围。
        changed = [name for name in original.namelist() if original.read(name) != cleaned.read(name)]  # 比较每个解压条目的字节。
        if any(name != "docProps/core.xml" for name in changed):  # 只允许身份元数据变更。
            raise ValueError("元数据清理意外改变了文档格式或内容。")  # 停止发布而非继续使用异常文件。
        validate_core_qnames(cleaned.read("docProps/core.xml"))  # 回读最终发布字节，再次核对名称空间语义。
    return {"file": target.name, "sha256": hashlib.sha256(target.read_bytes()).hexdigest(), "bytes": target.stat().st_size, "changedParts": changed}  # 回执只含公共文件名。


def extract(template: Path) -> dict:  # 以实际模板内容生成机器无关配置。
    with ZipFile(template) as archive:  # 只读公开模板。
        root = ET.fromstring(archive.read("word/styles.xml"))  # 获取模板样式表。
    styles = []  # 按常用样式顺序收集批准项。
    for style_id, use in zip(APPROVED, USES, strict=True):  # 避免使用位置清单与样式错位。
        node = root.find(f"w:style[@w:styleId='{style_id}']", NS)  # 依据既有模板的稳定 ID 读取。
        if node is None:  # 源资产缺少批准项时不能生成不完整规格。
            raise ValueError(f"模板缺少样式：{style_id}")  # 给出准确缺失项。
        rpr, ppr = node.find("w:rPr", NS), node.find("w:pPr", NS)  # 区分字符和段落属性。
        fonts = attrs(node.find("w:rPr/w:rFonts", NS))  # 保留中文与三个西文/复杂文字字体槽。
        spacing = attrs(node.find("w:pPr/w:spacing", NS))  # 保留行距和段前后原始单位。
        indent = attrs(node.find("w:pPr/w:ind", NS))  # 同时保留磅制和字符制缩进。
        notes = ["未在规范中单独明确的实现参数见 docs/样式对照表.md；不把完整属性树全部宣称为学校规定。"]  # 明确规格和规范并不完全等同。
        if style_id == "Normal":  # 标注已确认的摘要补充默认。
            notes.append("中英文摘要正文采用正文的两端对齐、首行缩进 2 字符，为已确认补充默认值。")  # 防止误标学校明文要求。
        if style_id == "附属正文":  # 标注小字正文补充默认。
            notes.append("两端对齐、首行缩进 0 字符，为已确认补充默认值。")  # 区别字体行距等规范参数。
        if style_id == "致谢正文":  # 标注致谢补充默认。
            notes.append("两端对齐、首行缩进 2 字符，为已确认补充默认值。")  # 保持来源可追溯。
        styles.append({"id": style_id, "name": value(node, "w:name"), "type": node.get(q("type")), "usage": use, "quick": enabled(node, "w:qFormat"), "priority": value(node, "w:uiPriority"), "fonts": fonts, "sizePt": float(value(node, "w:rPr/w:sz")) / 2 if value(node, "w:rPr/w:sz") is not None else None, "bold": enabled(node, "w:rPr/w:b"), "basedOn": value(node, "w:basedOn"), "next": value(node, "w:next"), "alignment": value(node, "w:pPr/w:jc"), "spacingRaw": spacing, "indentRaw": indent, "outlineLevel": value(node, "w:pPr/w:outlineLvl"), "pageBreakBefore": enabled(node, "w:pPr/w:pageBreakBefore"), "keepNext": enabled(node, "w:pPr/w:keepNext"), "characterProperties": property_tree(rpr), "paragraphProperties": property_tree(ppr), "provenanceNotes": notes})  # 完整记录可比较的实际参数。
    if len(styles) != 26 or sum(item["quick"] for item in styles) != 12:  # 验证发布承诺与文件一致。
        raise ValueError("公开资产必须包含 26 项批准样式、12 项常用入口。")  # 拒绝静默发布差异。
    return {"schemaVersion": 1, "version": "0.1.0-beta.1", "source": {"title": "郑州大学研究生学位论文写作规范", "schoolPublished": "2024-04-25", "documentDate": "2023-09", "url": "https://gs3.zzu.edu.cn/info/1045/13029.htm", "sha256": "cdf8f1d47fe0331ed0469d852acd36f257c7973dc3cec29474cc73dfae7949d4"}, "scope": "文字与段落样式；不安装页面布局、分节、页眉内容或自动标题编号。", "quickStyleIds": [item["id"] for item in styles if item["quick"]], "styles": styles}  # 不写入私人来源目录。


def main() -> None:  # 提供可复现的维护者入口。
    parser = argparse.ArgumentParser(description=__doc__)  # 创建命令行说明。
    parser.add_argument("--template", type=Path, required=True)  # 要求维护者选择已确认模板。
    parser.add_argument("--example", type=Path, required=True)  # 要求维护者选择已确认示例。
    parser.add_argument("--output", type=Path, required=True)  # 要求明确公开输出位置。
    args = parser.parse_args()  # 解析真实输入，不提供机器绑定路径。
    args.output.mkdir(parents=True, exist_ok=True)  # 创建公开资产目录。
    manifest = [sanitize(args.template, args.output / "郑大论文样式.dotx"), sanitize(args.example, args.output / "郑大论文样式示例.docx")]  # 各生成一个独立副本。
    spec_path = args.output / "style-spec.json"  # 固定公开样式规格名称。
    spec_path.write_text(json.dumps(extract(args.output / "郑大论文样式.dotx"), ensure_ascii=False, indent=2) + "\n", encoding="utf-8", newline="\n")  # 固定 LF，确保 Windows 与 Git/CI 中的 JSON 字节和哈希一致。
    manifest.append({"file": spec_path.name, "sha256": hashlib.sha256(spec_path.read_bytes()).hexdigest(), "bytes": spec_path.stat().st_size})  # 将配置本身纳入校验。
    (args.output / "asset-manifest.json").write_text(json.dumps({"version": "0.1.0-beta.1", "files": manifest}, ensure_ascii=False, indent=2) + "\n", encoding="utf-8", newline="\n")  # 固定 LF，输出跨平台一致且不含本机路径的发布清单。
    print(json.dumps({"styles": 26, "quickStyles": 12, "files": manifest}, ensure_ascii=False))  # 汇报实际产物与校验值。


if __name__ == "__main__":  # 只在直接运行脚本时执行。
    main()  # 保证作为模块导入时不会写文件。
