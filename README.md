# VLM_anysis 🩻✨

一个给医学影像研究、VLM 实验和人工标注准备的小工具箱。

它不是那种“一个命令包打天下”的大而全项目，更像是一个认真干活的工作台：

- 🤖 让视觉语言模型看医学图像
- 📝 用 Excel 组织病例并做人工评估
- 🔎 浏览、筛选、勾选、导出病例结果

如果你第一次打开这个仓库，不用从所有脚本里找方向。
最推荐直接从 `doctor_gui/` 开始。

## 🌟 What’s Inside

### `doctor_gui/`：当前主力标注界面

这是现在最完整、最适合日常使用的 GUI。

它可以：

- 📂 从 Excel 加载病例
- 🖼️ 根据图像路径列显示一张或多张图片
- 📄 显示指定文本列内容
- 👩‍⚕️👨‍⚕️ 支持多位操作者分别保存进度
- 💾 自动保存标注
- 🔁 重新打开后回到上次最后标注的位置
- 🔍 按一列或多列条件筛选病例
- 🔎 按住鼠标在原图上局部放大查看
- 📦 导出 `json / csv / xlsx / zip`

适合做：

- 医生或标注员逐例评估模型报告
- 多人并行标注同一批病例
- 只查看满足特定列值条件的病例子集

### `medical_image_analyzer.py`：轻量调用封装

这是一个很直接的 Python 封装，用来把本地图像发给支持视觉输入的聊天模型。

支持：

- 单图 / 多图分析
- 从文件读取 prompt
- 直接传入 prompt 文本
- 流式 / 非流式调用
- 自定义 `base_url`
- OpenAI SDK 兼容接口

### 其他脚本：研究流程小零件

根目录下还有一批更偏研究流程的脚本，比如：

- `run_vlm_loop.py`
- `run_vlm_loop_gemini.py`
- `sample_vlm_cases.py`
- `extract_vlm_samples_to_gold.py`
- `stylize_qa_pairs.py`

这部分更像“实验室桌面”，通常需要你按自己的数据路径、模型名和输出路径做本地调整。

### 其他 GUI

- `app.py`
  更偏骨病病例浏览 / 标注 / 统计
- `bi_case_gui/`
  BI 影像记录编辑与勾选界面

## 🚀 Quick Start

推荐 Python 3.10+。

先安装基础依赖：

```bash
pip install flask pandas openpyxl pillow openai
```

如果你的 Excel 导出环境缺少依赖，再补一个：

```bash
pip install xlsxwriter
```

然后直接启动主界面：

```bash
python3 doctor_gui/doctor.py
```

默认地址：

```text
http://127.0.0.1:5050
```

也可以带参数启动：

```bash
python3 doctor_gui/doctor.py \
  --excel /path/to/cases.xlsx \
  --sheet Sheet1 \
  --image-root /path/to/image_root \
  --port 5050
```

## 🧭 `doctor_gui` 使用流程

1. 打开首页，填写操作者和 Excel 路径。
2. 可选填写图片根目录，用于路径回退匹配。
3. 在配置页选择：
   - 图像路径列
   - 显示内容列
   - ID 列
   - 高亮相关列
   - 多列筛选条件
4. 进入标注页逐条评分。
5. 系统会自动保存，并在下次打开时恢复到这个操作者上次最后标注的位置。

## 🖼️ 图片路径怎么找

- 支持绝对路径
- 支持相对 Excel 文件目录的相对路径
- 单元格里可以放多张图，分隔符支持换行、`;`、`|`
- 如果原路径在当前机器不存在，但设置了 `image_root`，程序会尝试按路径后缀去匹配真实文件

## 🎛️ 目前支持的交互

- 按住鼠标在原图上看局部放大
- 松开鼠标时放大框消失
- 支持多列联合筛选病例
- 首页默认恢复到上次最后标注病例
- 顶部“首页”按钮可以强制回到首页

## 📦 导出结果

每位操作者的结果会单独保存在：

```text
doctor_gui/outputs/<operator_slug>/
```

导出 zip 里通常包含：

- `*.json`
- `*.csv`
- `*.xlsx`

## 🧪 一个最小调用示例

```python
from medical_image_analyzer import MedicalImageAnalyzer

analyzer = MedicalImageAnalyzer(
    model_name="Qwen/Qwen2-VL-72B-Instruct",
    api_key="YOUR_API_KEY",
    prompt_file="Prompt.md",
    base_url="https://api.siliconflow.cn/v1",
    stream=False,
)

result = analyzer.analyze_single_image("example.jpg")
print(result)
```

## 🗂️ 仓库结构

```text
VLM_anysis/
├── doctor_gui/              # Excel 图像/文本标注 GUI
├── bi_case_gui/             # BI 编辑/勾选 GUI
├── medical_image_analyzer.py
├── run_vlm_loop.py
├── run_vlm_loop_gemini.py
├── app.py
├── Prompt.md
└── README.md
```

## 🧹 放到 GitHub 前建议再看一眼

- 仓库里还有一些本地绝对路径，公开前最好先整理掉
- 如果包含医学影像、报告或标注结果，记得先做脱敏
- 检查历史提交里有没有敏感文件
- 最好再补一个 `requirements.txt` 或 `pyproject.toml`
- 如果之后准备给别人用，建议再加一份示例数据格式说明

## 💡 小提醒

这些工具更适合研究、标注和辅助分析，不应该直接当作临床诊断依据。

## 📄 License

如果准备公开发布，建议补一个明确的开源许可证，比如 `MIT` 或 `Apache-2.0`。
