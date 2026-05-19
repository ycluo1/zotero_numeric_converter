文件准备
建议将以下文件放在同一个文件夹中：
zotero_numeric_converter.py
requirements.txt
paper.docx
references.bib

提供 Zotero 本地数据库：
C:\Users\<YOUR_USERNAME>\Zotero\zotero.sqlite（可通过zotero-编辑-设置-高级设置找到）

安装依赖：
pip install -r requirements.txt

第一步：dry-run 检查
python zotero_numeric_converter.py --docx paper.docx --bib references.bib --zotero-sqlite ".\Zotero\zotero.sqlite" --out paper_zotero.docx --dry-run --require-zotero-uris

理想输出应类似：
Matched references: 180/180
Matched references with Zotero URI: 180/180
Matched references with Zotero numeric itemID: 180/180
Dry run only: no DOCX was written.

dry-run 检查无误后，运行正式转换：

python zotero_numeric_converter.py --docx paper.docx --bib references.bib --zotero-sqlite ".\Zotero\zotero.sqlite" --out paper_zotero.docx --require-zotero-uris

然后打开word文件，随便选中一个插入链接，编辑后点击确定，然后不停的点击确定即可


如果出现：
ERROR: Could not find the Reference section heading.
说明程序没有识别到文末参考文献标题。
可以手动指定标题：
--reference-heading "References"
