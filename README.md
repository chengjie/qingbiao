<p align="center"><img src="src/excel_compressor/assets/logo.svg" width="96" alt="轻表 Logo"></p>

# 轻表 · Excel 图片压缩

Windows 本地批量压缩工具。拖入 `.xlsx` 文件或文件夹，选择质量档位，生成压缩副本。无需安装 WPS 或 Excel，运行时不联网、不上传文件。

## 下载

[直接下载 Windows 版 QingBiao.exe](https://github.com/chengjie/qingbiao/releases/latest/download/QingBiao.exe) · [版本说明与 SHA-256 校验文件](https://github.com/chengjie/qingbiao/releases/latest)

适用于 Windows 10/11 x64，单文件约 43 MB，下载后双击运行。

## 使用

双击单文件 `QingBiao.exe`（可以重命名为“轻表.exe”）即可运行，不需要安装 Python、WPS 或 Excel。启动时会自动解包运行组件到临时目录，首次启动可能需要几秒。

1. 添加文件或文件夹；需要时勾选“包含子文件夹”。
2. 默认“均衡压缩”，适合商品与包装图片。
3. 输出默认位于各原文件旁的“压缩结果”目录，也可以指定统一目录。
4. 点击“开始压缩”，完成后双击文件行打开结果。

原文件始终保留。同名结果自动编号。重新压缩从队列里的原文件开始，不反复压缩上次结果。取消仅停止当前与后续任务，已完成的结果保留。单个文件失败后继续处理队列；双击失败行查看原因。

| 档位 | 图片质量 | 最大长边 |
|---|---:|---:|
| 清晰优先 | 80 | 原尺寸 |
| 均衡压缩 | 70 | 2000 px |
| 体积优先 | 60 | 1600 px |

高级设置支持质量 40–95、长边 800–8000 px 或保留原始分辨率。压缩属于有损操作，包装小字、条码和检测报告请放大核对。表格中的图片显示位置和尺寸保持不变。压缩收益低于 1% 时不额外生成文件。

## 兼容范围

- `.xlsx`，包括 WPS `DISPIMG` 单元格图片及普通浮动图片。
- PNG、JPEG 和 MPO 手机照片的主图。透明 PNG 保留透明度。
- 其他媒体、动态图、超过 40 百万像素或 128 MiB 的单张图片保持原样并提示。
- 只修改媒体及必要的图片关系与内容类型，不通过表格库重建工作簿，不计算公式，不更新外部链接。
- 所有输出条目经过 CRC 和 SHA-256 校验，修改图片完整解码检查，图片引用逐项核对。
- 带打开密码、数字签名的文件不处理；`.xls`、`.xlsm` 暂不支持。
- 压缩的是磁盘上已保存的版本，请先保存 WPS/Excel 中的最新修改。压缩期间若检测到原文件变化则停止交付该文件。

## 开发与测试

Python 3.12，建议 Windows 10/11 64 位。开发依赖锁定在 `requirements-build.txt`。

```powershell
py -3.12 -m venv .venv
.venv\Scripts\python -m pip install -r requirements-build.txt
.venv\Scripts\python -m pip install --no-build-isolation --no-deps -e .
.venv\Scripts\python -m pytest -q
.venv\Scripts\python launch.py
```

macOS/Linux 可使用相同源码验证；虚拟环境 Python 路径改成 `.venv/bin/python`。UI 测试使用 Qt offscreen，不需要操作用户桌面。

命令行批量压缩（输出逐文件 JSON）：

```powershell
.venv\Scripts\python -m excel_compressor.cli "D:\供应商文件" --recursive --preset balanced --output "D:\压缩结果"
```

## Windows 打包

在 Windows 安装 Python 3.12 x64 后，双击 `build-windows.cmd`。也可以从 PowerShell 执行 `./build-windows.ps1`。

脚本安装依赖、运行测试并生成：

- `dist/QingBiao.exe` 单文件程序。

首次构建需要网络下载 PyPI 依赖，成品运行无需网络。不要以管理员身份运行本程序。完整源码和构建脚本可用于替换 Qt 库后重新打包。

`.github/workflows/windows.yml` 提供 Windows 自动构建配置，手动触发或推送 `v*` 标签后生成构建附件；本项目不会自动上传用户的 Excel 文件。

官方技术参考：[Qt for Python](https://doc.qt.io/qtforpython-6/)、[PyInstaller 打包说明](https://www.pyinstaller.org/en/stable/usage.html)。Windows exe 必须在 Windows 环境构建。

## 代码结构

- `src/excel_compressor/engine.py`：流式文件包处理、图片压缩、校验、取消和输出保护。
- `src/excel_compressor/app.py`：PySide6 界面、后台队列与结果操作。
- `src/excel_compressor/cli.py`：命令行入口。
- `tests/`：WPS 关系保留、透明度、取消、失败恢复、重名保护、MPO 及界面测试。

引擎逐张处理图片，内存消耗主要由当前图片像素决定，不随整个工作簿大小线性增长。文件夹发现目前在添加阶段进行，特别大的目录可能需要等待；批量压缩在后台线程运行。

第三方组件说明见 `THIRD_PARTY_NOTICES.md`。


## 版本与验证

版本变更见 [更新记录](CHANGELOG.md)，已完成的验证见 [验证记录](docs/validation.md)。Logo 源文件与 Windows 图标位于 `src/excel_compressor/assets/`。
