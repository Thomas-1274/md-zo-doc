#!/bin/bash

# 📌 设置 Python 虚拟环境目录
VENV_DIR="$HOME/my_python_env"

# 🚀 **Step 0: 确保 Python 虚拟环境和 python-docx 安装**
if [ ! -d "$VENV_DIR" ]; then
    echo "🔄 创建 Python 虚拟环境: $VENV_DIR"
    python3 -m venv "$VENV_DIR"
    source "$VENV_DIR/bin/activate"
    pip install --upgrade pip
    pip install python-docx
    deactivate
    echo "✅ venv 创建完成，并安装了 python-docx！"
fi

source "$VENV_DIR/bin/activate"
if ! python3 -c "import docx" &>/dev/null; then
    echo "📦 安装 python-docx..."
    pip install python-docx
fi
deactivate

# 📌 交互式输入文件路径
read -e -p "请输入要清纯化的 Word 文件路径: " INPUT_DOCX
read -e -p "请输入 DOCX 模板路径: " TEMPLATE_DOCX

# 生成输出文件名
BASENAME=$(basename "$INPUT_DOCX" .docx)
OUTPUT_CLEAN_DOCX="/Users/apple/Documents/Word文档/${BASENAME}_purified.docx"

# 🚀 **Step 1: 使用模板标准化 Word**
echo "🚀 Step 1: Converting DOCX to standardized format using template..."
pandoc --standalone \
    --reference-doc="$TEMPLATE_DOCX" \
    -f docx \
    -t docx \
    -o "$OUTPUT_CLEAN_DOCX" "$INPUT_DOCX"

echo "✅ Word 规范化完成: $OUTPUT_CLEAN_DOCX"

# 🚀 **Step 2: 运行 Python 代码删除书签**
echo "🔖 Removing all bookmarks from the Word document..."

source "$VENV_DIR/bin/activate"
python3 - <<EOF
from docx import Document
from docx.oxml import OxmlElement

# 📌 目标文件路径
input_docx = "$OUTPUT_CLEAN_DOCX"
output_docx = "$OUTPUT_CLEAN_DOCX"

# 📌 打开 Word 文档
doc = Document(input_docx)

# 📌 删除所有书签
for bookmark in doc.element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bookmarkStart'):
    bookmark.getparent().remove(bookmark)

for bookmark in doc.element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bookmarkEnd'):
    bookmark.getparent().remove(bookmark)

# 📌 保存修改后的 Word 文档
doc.save(output_docx)

print(f"✅ 书签已全部删除，最终文件: {output_docx}")
EOF
deactivate  # 退出 Python 虚拟环境

# 🚀 **打开最终文件**
echo "✅ Word 清纯化完成: $OUTPUT_CLEAN_DOCX"
open "$OUTPUT_CLEAN_DOCX"

echo "🎉 清纯化任务完成！"