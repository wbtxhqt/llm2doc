[English](./README.md) | [简体中文](./README.zh-CN.md)

# LLM 文档编辑器 (llm2doc)

本项目是一个基于 Python 的工具，允许您使用自然语言命令来编辑 Microsoft Word (`.docx`) 文件。它利用大型语言模型 (LLM) 来解释您的指令，并对文档的内容和样式进行复杂的更改。

## 功能

-   **全面的内容与样式编辑**: 修改多种文档属性，包括：
    -   **文本级别格式**:
        -   文本内容、粗体、斜体、下划线（单线、双线等）。
        -   字体名称、大小和颜色 (十六进制)。
        -   高亮颜色。
        -   删除线、双删除线、上标和下标。
        -   全部大写和小写大写字母。
    -   **段落格式**:
        -   对齐方式（左对齐、居中、右对齐、两端对齐）。
        -   缩进（左、右和首行缩进）。
        -   间距（段前和段后）。
        -   行距（多倍、磅值和规则）。
        -   项目符号和编号列表。
    -   **结构元素**:
        -   创建和修改表格。
        -   添加和解析超链接。
-   **文档创建**: 从自然语言提示生成完整的 `.docx` 文件。
-   **多提供商支持**: 支持多个 LLM 提供商，包括 OpenAI、Google Gemini、Groq 等。
-   **精确定位修改**: 将 `.docx` 文件转换为带有唯一 ID 的详细 JSON 结构，从而实现高度精确的编辑。
-   **命令行界面**: 简洁的命令行工具 (`run.py`) 通过 `edit` 和 `create` 命令简化了整个流程。
-   **可扩展性**: 通过实现新的客户端，可以轻松添加新的 LLM 提供商。

## 工作原理

该工具遵循以下四个步骤：

1.  **转换为 JSON**: 将输入的 `.docx` 文件转换为详细的 JSON 格式，该格式捕获了其内容、结构和样式。
2.  **LLM 编辑**: 您的自然语言查询和 JSON 数据被发送到选定的 LLM。模型会分析您的请求并返回一个包含已修改对象的 JSON 列表。
3.  **应用修改**: 来自 LLM 的更改被合并回原始的 JSON 结构中。
4.  **重建 DOCX**: 使用更新后的 JSON 构建一个新的 `.docx` 文件，其中包含了您请求的修改。

## 核心组件

-   `run.py`: 用于运行端到端工作流的主命令行界面。
-   `llm2doc/converter.py`: 处理 `.docx` 与项目特定 JSON 格式之间的转换。
-   `llm2doc/editor.py`: 管理与 LLM 的交互，发送用户查询并处理响应。
-   `llm2doc/processor.py`: 将 LLM 返回的修改应用于 JSON 数据。
-   `llm2doc/llm_clients.py`: 包含用于连接各种 LLM API 的客户端。
-   `llm2doc/prompt/`: 此目录包含系统提示，用于指导 LLM 如何处理创建和修改任务。

## 安装与使用

### 先决条件

-   Python 3.7+
-   您打算使用的 LLM 服务的 API 密钥。

### 安装

1.  **克隆仓库:**
    ```bash
    git clone <repository-url>
    cd <repository-directory>
    ```

2.  **安装所需包:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **设置您的 API 密钥**
    为您计划使用的 LLM 提供商设置相应的环境变量。这是推荐且最安全的方法。

    *对于 Linux/macOS:*
    ```bash
    export OPENAI_API_KEY="your-openai-api-key"
    export GEMINI_API_KEY="your-gemini-api-key"
    export GROQ_API_KEY="your-groq-api-key"
    export DASHSCOPE_API_KEY="your-dashscope-api-key"
    export DEEPSEEK_API_KEY="your-deepseek-api-key"
    export DOUBAO_API_KEY="your-doubao-api-key"
    ```

    *对于 Windows (命令提示符):*
    ```bash
    set OPENAI_API_KEY="your-openai-api-key"
    set GEMINI_API_KEY="your-gemini-api-key"
    set GROQ_API_KEY="your-groq-api-key"
    set DASHSCOPE_API_KEY="your-dashscope-api-key"
    set DEEPSEEK_API_KEY="your-deepseek-api-key"
    set DOUBAO_API_KEY="your-doubao-api-key"
    ```

### 运行工具

`run.py` 脚本提供两个主要命令: `edit` 和 `create`。

#### 编辑现有文档

要修改现有的 `.docx` 文件，请使用 `edit` 命令:

```bash
python run.py edit \
  --input-docx "path/to/your/document.docx" \
  --query "将主标题更改为‘年度报告’并加粗。" \
  --output-docx "path/to/your/modified_document.docx" \
  --provider "chatgpt"
```

#### 创建新文档

要从查询生成新的 `.docx` 文件，请使用 `create` 命令:

```bash
python run.py create \
  --query "创建一个标题为‘我的第一个文档’的文档，标题加粗并居中。在下面添加一个段落，内容为‘你好，世界！’，文本为蓝色。" \
  --output-docx "new_document.docx" \
  --provider "chatgpt"
```

## 支持的提供商

您可以使用 `--provider` 参数指定 LLM 提供商。支持的提供商及其对应的所需环境变量如下：

| 提供商 | `--provider` 值 | 环境变量  |
|----------|--------------------|-----------------------|
| OpenAI   | `chatgpt`          | `OPENAI_API_KEY`      |
| Google   | `gemini`           | `GEMINI_API_KEY`      |
| Groq     | `groq`             | `GROQ_API_KEY`        |
| Qwen     | `qwen`             | `DASHSCOPE_API_KEY`   |
| Deepseek | `deepseek`         | `DEEPSEEK_API_KEY`    |
| Doubao   | `doubao`           | `DOUBAO_API_KEY`      |

如果未指定提供商，则默认为 `chatgpt`。

## 如何在 Jupyter Notebook 中使用

您可以在一个 Jupyter Notebook 单元格中运行整个过程。这对于交互式测试和实验非常有用。将以下代码复制并粘贴到一个单元格中：

```python
# 步骤 1: 如果尚未安装，请安装依赖项
!pip install -r requirements.txt

import os
from docx import Document
from run import create_docx, modify_docx

# --- 配置 ---
# 1. 在此处设置您要使用的 LLM 提供商
# 支持: "chatgpt", "gemini", "groq", "qwen", "deepseek", "doubao"
PROVIDER = "doubao"

# 2. 设置对应的 API 密钥
# 确保环境变量与您选择的提供商匹配。
os.environ["DOUBAO_API_KEY"] = "your-doubao-api-key-here"
# os.environ["OPENAI_API_KEY"] = "your-openai-api-key-here"
# os.environ["GEMINI_API_KEY"] = "your-gemini-api-key-here"


# --- 选项 1: 创建一个新文档 ---
print(f"--- 正在使用 {PROVIDER.upper()} 运行创建示例 ---")
creation_query = "创建一个标题为‘新的 Notebook 文档’的文档，标题加粗。在下面添加一个段落，内容为‘这是从 Notebook 创建的。’"
creation_output = "notebook_created_zh.docx"

try:
    create_docx(
        query=creation_query,
        output_docx=creation_output,
        provider=PROVIDER
    )
    print(f"成功创建 '{creation_output}'")
except Exception as e:
    print(f"创建过程中发生错误: {e}")


# --- 选项 2: 编辑现有文档 ---
print(f"\n--- 正在使用 {PROVIDER.upper()} 运行修改示例 ---")
# 创建一个用于测试的虚拟文档
edit_input_file = "notebook_for_editing_zh.docx"
if not os.path.exists(edit_input_file):
    doc = Document()
    doc.add_paragraph("这是原始文本。它是黑色的，没有加粗。")
    doc.save(edit_input_file)
    print(f"已创建虚拟文件: '{edit_input_file}'")

edit_query = "将文本‘黑色’改为‘紫色’，并使字体颜色变为紫色。"
edit_output_file = "notebook_modified_zh.docx"

try:
    modify_docx(
        input_docx=edit_input_file,
        query=edit_query,
        output_docx=edit_output_file,
        provider=PROVIDER
    )
    print(f"成功修改 '{edit_input_file}'。输出已保存到 '{edit_output_file}'")
except Exception as e:
    print(f"修改过程中发生错误: {e}")

## 未来计划

-   **PPTX 支持**: 我们计划扩展此工具的功能，以类似的方式支持编辑 Microsoft PowerPoint (`.pptx`) 文件。