#!/bin/bash

# 📌 设置 Python 虚拟环境目录
VENV_DIR="$HOME/my_python_env"

# 🚀 **Step 0: 确保 Python 虚拟环境已创建**
if [ ! -d "$VENV_DIR" ]; then
    echo "🔄 创建 Python 虚拟环境: $VENV_DIR"
    python3 -m venv "$VENV_DIR"
    source "$VENV_DIR/bin/activate"
    pip install --upgrade pip
    pip install python-docx
    deactivate
    echo "✅ `venv` 创建完成，并安装了 python-docx！"
fi

# 🚀 **Step 0.1: 确保 `python-docx` 安装**
source "$VENV_DIR/bin/activate"
if ! python3 -c "import docx" &>/dev/null; then
    echo "📦 安装 python-docx..."
    pip install python-docx
fi
deactivate

# 📌 交互式输入
read -e -p "请输入Markdown文件路径，注意Md语法！（可跳过Md-Odt转化）: " INPUT_MD
read -e -p "请输入 ODT 模板路径，注意由Libre转化来！（可跳过Md-Odt转化）: " TEMPLATE_ODT

read -e -p "请输入 DOCX 模板路径（必填！）: " TEMPLATE_DOCX
while [[ ! -f "$TEMPLATE_DOCX" ]]; do
    echo "❌ 错误: DOCX 模板文件不存在，请重新输入。"
    read -e -p "请输入 DOCX 模板路径: " TEMPLATE_DOCX
done




# 🚀 **Step 1: 处理 Markdown → ODT**
if [[ -n "$INPUT_MD" && -n "$TEMPLATE_ODT" ]]; then
    if [[ -f "$INPUT_MD" && -f "$TEMPLATE_ODT" ]]; then
        # 生成 ODT 输出文件名
        BASENAME=$(basename "$INPUT_MD" .md)
        # 生成 ODT 文件名，若已存在则编号递增
ODT_BASE="/Users/apple/Documents/Word文档/${BASENAME}"
OUTPUT_ODT="${ODT_BASE}.odt"

if [[ -f "$OUTPUT_ODT" ]]; then
    COUNT=1
    while [[ -f "${ODT_BASE}_$COUNT.odt" ]]; do
        ((COUNT++))
    done
    OUTPUT_ODT="${ODT_BASE}_$COUNT.odt"
fi

        echo "🚀 Step 1: Converting MD to ODT using template..."
pandoc --standalone \
    --reference-doc="$TEMPLATE_ODT" \
    --wrap=none \
    -f markdown+smart+ignore_line_breaks \
    -t odt \
    -o "$OUTPUT_ODT" "$INPUT_MD"

        if [[ $? -eq 0 ]]; then
            echo "✅ ODT 文件已生成: $OUTPUT_ODT"
            echo "⚠️ 请在 Zotero 中使用 ODF-Scan 将 ODT 转化，并调整、转化为适宜的DOCX"
        else
            echo "❌ Pandoc 转换失败，请检查 Markdown 格式或 Pandoc 版本。"
        fi
    else
        echo "❌ Markdown 或 ODT 模板路径无效，跳过转换。"
    fi
fi

# **无限循环，必须输入 DOCX 文件**
while true; do
    read -e -p "🔄 请输入要转换的 DOCX 文件路径（必填！）: " FINAL_DOCX
    while [[ ! -f "$FINAL_DOCX" ]]; do
    echo "❌ 错误: 文件不存在，请重新输入。"
    read -e -p "🔄 请输入要转换的 DOCX 文件路径: " FINAL_DOCX

    # ⬇️ **新增：如果用户按回车，不输入文件，则退出**
    if [[ -z "$FINAL_DOCX" ]]; then
        echo "🚪 退出转换流程..."
        break 2  # **退出整个 while true 循环**
    fi
done

    # 🚀 **生成输出文件名**
    DOCX_BASE="/Users/apple/Documents/Word文档/$(basename "$FINAL_DOCX" .docx)_final"
OUTPUT_FINAL_DOCX="${DOCX_BASE}.docx"

if [[ -f "$OUTPUT_FINAL_DOCX" ]]; then
    COUNT=1
    while [[ -f "${DOCX_BASE}_$COUNT.docx" ]]; do
        ((COUNT++))
    done
    OUTPUT_FINAL_DOCX="${DOCX_BASE}_$COUNT.docx"
fi

    # 🚀 **Step 2: 应用 Word 模板**
    echo "🚀 Converting $FINAL_DOCX to final formatted DOCX..."
    pandoc --standalone \
        --reference-doc="$TEMPLATE_DOCX" \
        -f docx \
        -t docx \
        -o "$OUTPUT_FINAL_DOCX" "$FINAL_DOCX"

    if [[ $? -eq 0 ]]; then
        echo "✅ Word 模板刷格式完成: $OUTPUT_FINAL_DOCX"
    else
        echo "❌ Pandoc 转换失败，请检查 DOCX 文件。"
        OUTPUT_FINAL_DOCX="${FINAL_DOCX%.docx}_final.docx"  # 保持文件格式，避免丢失
        cp "$FINAL_DOCX" "$OUTPUT_FINAL_DOCX"
    fi

    # 🚀 **Step 3: 运行 Python 进行文本处理**
    echo "🔄 Running Python script to adjust final DOCX..."

    source "$VENV_DIR/bin/activate"
    python3 - <<EOF
from docx import Document
from docx.oxml import OxmlElement

# 📌 目标文件路径
input_docx = "$OUTPUT_FINAL_DOCX"
output_docx = "$OUTPUT_FINAL_DOCX"

# 📌 打开 Word 文档
doc = Document(input_docx)

# 📌 目标文本
target_text = "然后按 Zotero 插件中的 Refresh 继续使用引注。"

# 📌 遍历 Word 文档中的所有段落，找到目标文本并在其后面添加两个换段符（↵）
found = False
for para in doc.paragraphs:
    if target_text in para.text:
        # **在当前段落后插入两个新的空段落**
        new_p1 = OxmlElement("w:p")  # 创建新段落
        new_p2 = OxmlElement("w:p")  # 再创建一个新段落
        para._element.addnext(new_p1)
        para._element.addnext(new_p2)
        
        found = True
        break  # 只修改第一个匹配项，避免多次插入

# 📌 **删除所有书签**
for bookmark in doc.element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bookmarkStart'):
    bookmark.getparent().remove(bookmark)

for bookmark in doc.element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bookmarkEnd'):
    bookmark.getparent().remove(bookmark)

# 只有在找到目标文本时才保存
if found:
    doc.save(output_docx)
    print(f"✅ 处理完成，已在目标文本后插入 2 个换段符，防止zotero转化吞字。并删除了所有书签: {output_docx}")
else:
    print(f"⚠️ 未找到zotero乱码字段的目标文本。但已删除所有书签: {output_docx}")

EOF
    deactivate

    echo "✅ 最终 Word 文档完成: $OUTPUT_FINAL_DOCX"
    open "$OUTPUT_FINAL_DOCX"

    # 🔄 继续处理下一个 Word 文件
    read -e -p "🟡 如有，请输入下一个 DOCX 文件路径。如无，请直接回车退出: " FINAL_DOCX

# ⬇️ **新增：如果用户按回车，不输入文件，则退出**
if [[ -z "$FINAL_DOCX" ]]; then
    echo "🎉 恭喜！所有转换任务已完成！"
    break
fi
done

echo "✍️ 作者：as小黄"