# 开发手册

## 1. 设计约束

1. **单文件即一个批次**：读取、映射、写出均以"记录集合"为单位，单文件与批量走同一条代码路径。
2. **单份失败不中断整批**：解析失败的文件被跳过并记录原因，失败清单回传调用方。
3. **命令行默认安全，界面默认直连**：命令行下 `classify`、`upload` 需显式 `--apply` 才执行；
   图形界面不设预演步骤，但上传前提供 Ping 连通性测试，且输入框不预填占位内容。
4. **界面与命令行为同一实现**：界面只负责收集参数，业务逻辑集中在可独立测试的服务函数中。

## 2. 模块职责

| 模块            | 职责                                        | 主要接口                                                                                                                             |
| --------------- | ------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------ |
| `formats.py`    | 格式注册表：可读、可写、格式分类            | `INPUT_SUFFIXES` `OUTPUT_SUFFIXES` `format_kind` `support_matrix`                                                                    |
| `core.py`       | 字段映射（纯函数，无 IO）                   | `map_records(records, target_headers, field_mapping)`                                                                                |
| `readers.py`    | 多格式读取、目录批量读取、图片/扫描件 OCR   | `read_any` `read_folder` `iter_input_files` `read_image` `ocr_pdf_text` `ocr_available` `is_plain_content` `content_of`              |
| `writers.py`    | 记录写出、正文直转、Excel/CSV 追加          | `write_output` `write_text_document`                                                                                                 |
| `converters.py` | 格式转换（单文件 / 整目录）                 | `convert_file` `convert_folder` `normalize_target_format`                                                                            |
| `merger.py`     | 多文档合并为一个文档                        | `merge_any` `merge_tables` `merge_word_documents` `merge_text_documents` `collect_headers` `pick_mode`                               |
| `classify.py`   | 按规则归档                                  | `ClassifyConfig` `plan_moves` `apply_moves` `organize`                                                                               |
| `rules.py`      | 规则默认值、界面行与配置字典互转、临时 JSON | `DEFAULT_MAPPING` `DEFAULT_CLASSIFY` `mapping_rows` `build_mapping` `classify_rows` `build_classify` `write_temp_rules` `read_rules` |
| `uploader.py`   | 上传与连通性测试                            | `UploadConfig` `build_request` `upload_file` `upload_folder` `ping_server`                                                           |
| `cli.py`        | 命令行入口（含旧式调用兼容）                | `main(argv)` `build_parser`                                                                                                          |
| `gui.py`        | 图形界面与服务层                            | `OfficeAssistantApp`、`run_report` / `run_convert` / `run_merge` / `run_scan` / `run_classify` / `run_upload` / `ping_upload`        |

## 3. 数据流与记录约定

```
输入文件 ──read_any──> records: list[dict]
                          │
                          ├─ map_records（可选：字段映射）
                          │
                          └─ 合并/转换策略 ──> write_output / write_text_document ──> 输出文件
```

`readers` 的返回约定：

| 情况                           | 返回值                                    |
| ------------------------------ | ----------------------------------------- |
| 有结构（表格行、`字段：内容`） | `[{"字段": "值", ...}, ...]`              |
| 无结构（正文、无键值文本）     | `[{"内容": "整篇正文"}]`                  |
| 解析不到内容                   | `[]`                                      |
| 格式不支持                     | `raise ValueError("不支持的输入格式：…")` |

- 表格类输入天然是记录；正文类输入用 `is_plain_content()` 判定，转换时走"原文直转"。
- 批量读取默认 `skip_errors=True`，失败项写入传入的 `failed` 列表。

## 4. 界面结构（gui.py）

### 4.1 布局

主窗口使用 grid 划分三个区域：

| 区域   | 位置           | 说明                                                                                              |
| ------ | -------------- | ------------------------------------------------------------------------------------------------- |
| 工作区 | 第 0 行第 0 列 | 标签页；随窗口缩放                                                                                |
| 信息栏 | 第 0 行第 1 列 | 宽度固定：`columnconfigure(1, minsize=SIDE_PANEL_WIDTH)`，内部 `Text` 使用 `width=1` 以免撑开列宽 |
| 状态栏 | 第 1 行        | 高度固定、单行                                                                                    |

标签页分两类：

- 表单页（报表汇总、格式转换、合并文档、批量提取、文件归档、自动上传）：`_new_tab()` 返回 `ScrollableTab` 的内容容器，
  窗口偏小时可滚动；构建完成后需调用 `bind_wheel()` 递归绑定滚轮事件。
- 展示页（使用说明、支持的格式）：`_new_plain_tab()`。此类页面内部自带滚动 `Text`，
  若放入 `ScrollableTab` 的 Canvas，`Text` 高度会退化为 1 像素。

### 4.2 状态栏与消息

- 状态栏仅显示一行：`就绪` / `需要…` / `成功 · …` / `失败 · …` / `正在运行…` / `启动耗时 N.NN 秒`。
  除 `fatal=True` 的错误外，均由 `status_reset_ms`（默认 4000 毫秒）后自动复位为 `就绪`。
- 提示与处理明细统一写入信息栏，入口为 `_append_message(text, kind)`，
  `kind` 取 `info` / `success` / `error` / `detail`，对应不同颜色；明细缩进显示。
  行数超过 `MESSAGE_MAX_LINES` 时一次丢弃最早的 `MESSAGE_TRIM_LINES` 行。
- 界面不使用 `messagebox`：缺少输入、任务失败等均通过状态栏与信息栏表达。

### 4.3 线程模型

- `run_task(task, background=True)` 在后台线程执行任务，任务通过 `log` 回调投递过程消息到 `queue.Queue`。
- `_poll_queue()` 为常驻轮询（`POLL_INTERVAL_MS`，默认 80 毫秒），在主线程消费队列并更新界面。
- 执行期间执行按钮禁用，`_busy` 置位；收到 `finish` 消息后恢复。
- `background=False` 时同步执行并立即处理队列，供测试在无并发依赖的情况下验证完整链路。

### 4.4 规则编辑与临时文件

字段映射与归档规则均可在界面中编辑，避免要求使用者手工编写 JSON：

| 组件                               | 说明                                                                                                                                      |
| ---------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| `MappingDialog` / `ClassifyDialog` | 继承 `_TableRuleDialog`，以 `ttk.Treeview` 展示规则行，支持添加/编辑/删除/上移/下移；确定后由 `collect()` 返回配置字典                    |
| `_FieldsDialog`                    | 单行编辑弹窗，字段标签由子类的 `edit_labels` 给出；尺寸按内容自适应后居中                                                                 |
| `rules.py`                         | 纯逻辑：默认值、`mapping_rows` / `build_mapping`、`classify_rows` / `build_classify`（含扩展名规范化）、`write_temp_rules` / `read_rules` |

数据流：`设置…` → 对话框返回配置字典 → `rules.write_temp_rules(kind, payload)` 写入
`<系统临时目录>/info-map-rules/<kind>.json` → 路径回填到输入框 → 执行时按普通配置文件读取。
输入框已填写路径时，编辑器会先加载该文件（读取失败则回退默认规则并在状态栏提示）。
对话框的模态设置封装在 `_activate_modal()` 中，对无显示环境容错，便于自动化测试直接构造。

### 4.5 窗口定位

主窗口与弹窗均由工具函数定位，不依赖窗口管理器的默认层叠位置：

| 函数                             | 用途                                                                   |
| -------------------------------- | ---------------------------------------------------------------------- |
| `center_on_screen(window, w, h)` | 主窗口启动时按 `MAIN_WIDTH` × `MAIN_HEIGHT` 居中到屏幕，无需先完成布局 |
| `center_window(window, parent)`  | 弹窗居中到父窗口；父窗口不可见时（如测试环境）退回屏幕居中             |
| `_frame_origin(window)`          | 取窗口外框左上角坐标，供居中计算使用                                   |

两个约束：

1. `center_window()` 必须在控件构建完成后调用。布局结束前 `wm_geometry()` 一律返回
   `1x1+0+0`，`winfo_width()` 也会返回 1，只有调用 `update_idletasks()` 之后才能取到真实尺寸。
2. 居中必须使用 `wm_geometry()` 的 `+X+Y`（外框坐标），不能用 `winfo_rootx()` / `winfo_rooty()`。
   后者返回客户区坐标，比外框右下偏移约 8 × 31 像素（Windows 实测），据此居中会让弹窗整体偏右下。

### 4.6 界面测试要点

- 使用 `geometry("1040x840+4000+4000")` 将窗口移出可视区域后再测量尺寸；
  `withdraw()` 状态下所有 `winfo_*` 返回 1，无法用于布局断言。
- 尺寸断言：状态栏高度 ≤ 40 像素且不随文本长度变化；右栏宽度等于 `SIDE_PANEL_WIDTH` 减去外边距。
- 不依赖后台线程与定时器的用例应使用 `run_task(..., background=False)`。
- 窗口定位不单独断言：`center_window()` 依赖真实映射后的外框坐标，`withdraw()` 或未映射状态下取不到有效值。

## 5. 扩展方式

### 5.1 新增输入格式（以 `.rtf` 为例）

1. `formats.py`：将 `.rtf` 加入 `TEXT_INPUT_SUFFIXES`，并在 `FORMAT_NOTES` 补充说明。
2. `readers.py`：实现 `read_rtf(path) -> list[dict]`，按第 3 节的返回约定；在 `read_any()` 的分发分支中接入。
3. `tests/`：在 `conftest.py` 增加样例 fixture，在 `test_readers.py` 增加用例。

新增输出格式同理：在 `formats.py` 注册，在 `writers.py` 增加 `_write_xxx`，并在 `write_output()` 中分发。

### 5.2 新增界面功能

1. 在 `gui.py` 的服务层实现 `run_xxx(...) -> str`（返回可显示的结果摘要，不依赖 tkinter）。
2. 在 `_build_xxx_tab()` 中构建控件；表单页用 `_new_tab()`，展示页用 `_new_plain_tab()`。
3. 添加事件处理方法：校验输入后调用 `run_task`，输入缺失时用 `show_message` 提示。
4. 在 `TAB_ORDER`、`TAB_HINTS` 中登记该页，并同步 `tests/test_gui.py` 的 `EXPECTED_TABS`。

## 6. 上传接口约定

### multipart 模式（默认）

```
POST {UPLOAD_URL}
Content-Type: multipart/form-data; boundary=...
Authorization: Bearer {API_TOKEN}

--boundary
Content-Disposition: form-data; name="<extra_fields 的键>"
<值>
--boundary
Content-Disposition: form-data; name="file"; filename="报表汇总.xlsx"
Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet

<文件二进制>
--boundary--
```

### json 模式

```json
{
  "filename": "报表汇总.xlsx",
  "content_base64": "UEsDBBQ...",
  "content_type": "application/vnd...",
  "<extra_fields>": "..."
}
```

行为约定：

- 2xx 视为成功；4xx 不重试；5xx 与网络错误按 `retries`、`retry_delay` 重试。
- 超时由 `timeout` 控制；内网自签名证书可将 `verify_ssl` 设为 `false`。
- 配置仍为占位符时 `validate()` 抛出异常，命令行 `--apply` 会被拒绝。
- `ping_server()` 仅发送 GET 请求：收到任何 HTTP 响应（含 404、405）即判定可达，连接失败或超时判定不可达。

## 7. 构建、运行与测试

项目只使用一个环境：根目录下的 `.venv`，由 `scripts/setup_env.py` 维护（`.gitignore` 忽略 `.venv*/`）。
若该目录不可用或结构不完整，配置脚本会先删除再重建，不依赖系统 Python 环境。
在 VS Code 中选择解释器时指向 `.venv\Scripts\python.exe`。

```powershell
# 安装（含测试依赖）
pip install -e ".[dev]"

# 运行测试
pytest -q

# 启动图形界面
python scripts\gui_app.py

# 生成演示数据
python scripts\make_demo_data.py
```

测试约定：

- 全部用例离线运行，不产生真实网络请求（用 `monkeypatch` 替换 `urllib.request.urlopen`）。
- 样例数据集中在 `tests/conftest.py`；`inbox` fixture 内含一份损坏的 `.docx`，用于验证容错路径。
- 纯逻辑（`map_records`、`collect_headers`、`pick_mode`、`ClassifyRule.matches`）优先做成无副作用函数并单独测试。

### 批处理的编码约束

`一键配置环境.bat`、`启动界面.bat` 为 UTF-8 无 BOM 文件，配合 `chcp 65001` 使用。
经实测，cmd 在该代码页下解析含中文的较长行时可能发生字节错位，导致 `echo` 内容被当作命令执行。

因此遵循以下约定：

- 批处理只保留定位 Python 解释器的逻辑与极短中文提示，其余逻辑与中文输出全部放在 `scripts/setup_env.py`。
- 设置 `PYTHONUTF8=1`、`PYTHONIOENCODING=utf-8`，避免 Python 输出与 cmd 代码页不一致导致乱码。
- 探测解释器必须实际执行 `X -c "import sys"` 并判断返回值；`where` 能定位到 `py.exe` 不代表 `py -3` 可用。

## 8. 回归检查清单

1. `pytest -q` 全部通过。
2. `info-map formats` 输出正常。
3. `python scripts\make_demo_data.py` 生成演示数据。
4. 依次执行 `report`、`merge`、`convert`、`scan`、`classify`（预演）、`upload`（预演）。
5. 确认 `upload --apply` 在占位符未替换时仍被拒绝。
6. 图形界面：8 个标签页存在；切换标签页时信息栏说明随之更新；执行一次任务后状态栏显示 `成功 · …` 并自动复位。
7. 在干净环境双击 `一键配置环境.bat`，应完成依赖安装并通过自检；
   输出中应能看到 ".xls 读取组件检查通过" 与 "OCR 组件检查通过"。
8. OCR 链路：用一张含文字的图片验证 `read_image()` 能返回记录；扫描版 PDF 应自动回退到 OCR。
9. 需要交付免安装版时，`python scripts\make_release.py` 应输出「自检结果：通过」，
   且报告中 `打包运行：是`、`OCR 模型：3 个`、`OCR 识别：成功`、`界面构建：成功（8 个标签页）` 均符合预期。
10. 需要交付库版时，`python scripts\make_package.py` 应生成 whl、tar.gz 与源码包；
    安装该 whl 后 `info-map formats` 输出正常。
11. `git ls-files` 不应出现 `materials/`、`image/`、`release/`、`packages/` 下的任何文件。

## 9. 已知限制

- PDF 优先使用文本层；无文本层时由 `ocr_pdf_text()` 逐页渲染后识别；图片则直接由 `read_image()` 识别。
  OCR 引擎（rapidocr-onnxruntime）在进程内惰性创建并复用，模型加载耗时较长。
- 复杂表头按首行取值。
- 文本编码依赖自动探测（UTF-8 / GBK / GB18030 等），非常规编码需先转换。
- 图形界面的自动化测试依赖桌面环境，无显示环境下相关用例会被跳过。

## 10. 打包与发布

项目按交付形态分两个版本，产物目录分开存放，均不进入仓库：

| 版本        | 产物目录    | 生成方式                     | 适用对象                          |
| ----------- | ----------- | ---------------------------- | --------------------------------- |
| 免安装版    | `release/`  | `scripts/make_release.py`    | 没有 Python 环境的基层使用者      |
| Python 库版 | `packages/` | `scripts/make_package.py`    | 有 Python 环境，需命令行或二次开发 |

### 免安装版（release/）

由 `scripts/make_release.py` 调用 PyInstaller 生成，面向没有 Python 环境的使用者。

```powershell
python -m pip install pyinstaller              # 维护者一次性准备（已列入 pyproject 的 dev 额外依赖）
python scripts\make_release.py                 # 目录版（默认，推荐）
python scripts\make_release.py --onefile       # 单文件 exe
python scripts\make_release.py --zip           # 额外生成 zip，供 Release 附件上传
python scripts\make_release.py --keep-build    # 保留 build/ 与 dist/ 以便排查
```

### 两种形态的选择

| 形态                    | 启动                                       | 说明                                                               |
| ----------------------- | ------------------------------------------ | ------------------------------------------------------------------ |
| 目录版（默认）          | 直接启动                                   | 实测构建 41.8 秒，裁剪后 248.2 MB；分发整个目录（含 `_internal/`） |
| 单文件版（`--onefile`） | 每次启动先解包到临时目录，大体积下明显变慢 | 实测构建 54.3 秒、119.2 MB；只需发一个 exe                         |

### 必须显式收集的资源

下列资源不进包时会「构建成功但功能失效」，是打包的主要坑位：

| 资源                              | 体积     | 参数                 | 缺失后果                     |
| --------------------------------- | -------- | -------------------- | ---------------------------- |
| `rapidocr_onnxruntime` 模型与字典 | 约 16 MB | `--collect-all`      | 图片与扫描件无法识别         |
| `onnxruntime` 运行时              | 约 37 MB | `--collect-binaries` | 创建 OCR 引擎时崩溃          |
| `pypdfium2_raw\pdfium.dll`        | 约 7 MB  | `--collect-all`      | 扫描版 PDF 无法渲染          |
| `docs/USER_GUIDE.md`              | 12 KB    | `--add-data`         | 「使用说明」页退化为内置简版 |

体积主要来自 `cv2`（111.8 MB）、`onnxruntime`（35.8 MB）与 `numpy`（26.3 MB）。
其中 `cv2\opencv_videoio_ffmpeg500_64.dll`（29.4 MB）仅在调用 `cv2.VideoCapture`
等视频接口时才加载，由 `assemble()` 末尾的 `trim()` 在构建后删除，目录版因此
从 277.6 MB 降到 248.2 MB。该文件仅存在于目录版；单文件版的资源在包内，无法裁剪。
自检中的真实 OCR 识别在裁剪后运行，可用于确认裁剪未破坏图像链路。

入口脚本固定为 `scripts/gui_app.py`，构建名用 ASCII（`OfficeAssistant`），组装到发布目录时
再改名为 `基层办公自动化助手.exe`，避开工具链对非 ASCII 路径的兼容问题。

### 打包后的路径解析

tkinter/PyInstaller 环境下的路径与源码运行不同，已在 `gui.manual_roots()` 中统一处理：

1. `sys._MEIPASS`：**包内资源目录**，`--add-data` 放进去的 `docs/` 在这里；
2. `sys.executable` 所在目录：随包分发的 `docs/` 在这里（使用者可单独打开手册）；
3. 源码布局的项目根目录与当前工作目录。

`gui_app.py` 在打包后不再插入 `src/` 到 `sys.path`，模块由 PyInstaller 提供。

### 产物结构与分发

```
release/基层办公自动化助手/
  基层办公自动化助手.exe     主程序
  _internal/                 运行时（目录版；勿删、勿改名）
  docs/USER_GUIDE.md         手册独立副本
  config/*.example.json      示例配置（uploader 的报错提示会引用该路径）
  README.md
  使用说明.txt               由 make_release.py 生成，联系方式取自 SUPPORT_CONTACTS
```

`build/`、`dist/`、`release/` 与 `self-check.txt` 均已加入 `.gitignore`；
体积远超仓库限制的 zip 应作为 GitHub Release 附件上传，不要提交。

### 产物自检

构建完成后脚本会用 `--self-check` 启动产物（`scripts/gui_app.py`）：

- 记录版本、Python 版本、`sys._MEIPASS` 与手册实际路径；
- 逐项校验 pandas / openpyxl / python-docx / pypdf / xlrd / Pillow / OCR 组件；
- 核对 OCR 模型数量与 `pdfium.dll` 是否存在（即上表的收集项）；
- 真实加载一次 OCR 引擎，并用 Pillow 生成图片跑通一次完整识别（`read_image()`），
  确认 onnxruntime 与 cv2 在裁剪后仍可用；
- 真实构建一次 `OfficeAssistantApp`（移到屏幕外），确认 tkinter 与 8 个标签页正常。

窗口版（`--windowed`）不显示异常堆栈，启动期错误几乎是静默失败，因此以上检查不能省。
结果写入当前目录的 `self-check.txt`，脚本读取后判定成败并删除该文件。

### 与源码运行的区别

- `report_failure()` / `gui_app.fatal()`：无控制台时改用系统消息框提示启动失败。
- `scripts/setup_env.py` 不安装 PyInstaller，基层使用者的环境保持精简。
- exe 为窗口版（`--windowed`），没有控制台；排查问题靠 `--self-check` 报告与界面右侧「消息」栏。

### Python 库版（packages/）

```powershell
python -m pip install build                    # 维护者一次性准备（dev 额外依赖）
python scripts\make_package.py                 # wheel + sdist + 源码包
python scripts\make_package.py --no-source     # 只要 wheel 与 sdist
```

| 产物                                      | 体积   | 用途                                                                   |
| ----------------------------------------- | ------ | ---------------------------------------------------------------------- |
| `information_mapper-<版本>-py3-none-any.whl` | 58 KB  | `pip install` 后提供 `info-map` 命令，界面用 `python -m information_mapper.gui` |
| `information_mapper-<版本>.tar.gz`           | 69 KB  | sdist，供 pip 构建或源码分发                                            |
| `information-mapper-<版本>-源码.zip`         | 120 KB | 含 `scripts/`、`docs/`、`config/` 与两个 `.bat`，解压后双击「一键配置环境.bat」 |

默认用 `--no-isolation` 本地构建，避免联网拉取构建依赖；失败时自动改用隔离环境重试。
wheel 只包含 `src/information_mapper`，不含 `docs/`，因此 pip 安装后的「使用说明」页
会退化为内置简版；需要完整手册请用源码包或免安装版。
`make_package.py` 与 `make_release.py` 都以 `build/` 为中间目录，二者均在 `.gitignore` 中。

### 项目目录组织

仓库内只保留源码、脚本、文档、示例配置与测试，其余内容分三类处理：

| 目录                  | 内容                                                     | 是否入库       |
| --------------------- | -------------------------------------------------------- | -------------- |
| `materials/`          | 演示视频、答辩材料、软件截图、实践证明等材料与证明       | 已忽略（约 103 MB） |
| `release/`、`packages/` | 两个版本的构建产物                                     | 已忽略         |
| `build/`、`dist/`、`output/`、`self-check.txt` | 中间产物与运行输出            | 已忽略         |

`materials/` 按用途分「演示视频 / 答辩材料 / 证明材料 / 素材 / 截图」五个子目录；
历史上被跟踪的材料文件已从 git 历史中清除，仓库转为公开时不存在历史副本。
