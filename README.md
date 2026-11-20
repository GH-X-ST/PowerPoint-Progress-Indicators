# PowerPoint Progress Indicators 中文版

这是基于 Olaf Nöhring 发布的 Progress Indicator 1.960 独立维护的中文 VBA 加载项。它读取 PowerPoint 文稿中的章节和页序，在每页生成章节导航。放映时，点击章节名可跳到该章首页，点击页标可跳到对应的幻灯片。[上游项目](https://datenbank-projekt.de/projekte/progress-indicator-powerpoint-2010-2013-2016-32-and-64-bit-windows-mac)

下载并解压 CN.11.1 发布包，选择 `ProgressIndicators_v1960_3264b_SB_CN11_1.ppam` 安装。同名 `.pptm` 是可在 PowerPoint 中打开的维护文稿。[下载 CN.11.1](https://github.com/GH-X-ST/PowerPoint-Progress-Indicators/releases/download/CN.11.1/ProgressIndicators-CN.11.1.zip)

## 导航样式

- 40 种预设提供不同的章节和页标结构，包括胶囊、圆节点、刻度、直角分段和连续轨道。
- 可停靠在顶部、底部、左侧或右侧。“章节均长”给每个非空章节相同的长度。“进度均长”按计入进度的页数分配各章长度。
- CN.11.1 移除了“层次 > 深红层叠”在章节交界处的灰色分隔线。章节面无缝相接，靠整块投影呈现层次。当前章与深红连接条共用轮廓，当前页以带光晕的短条标识。

## 在 Windows 桌面 PowerPoint 中安装

1. 打开“文件 > 选项 > 加载项”，在窗口底部选择“管理 > PowerPoint 加载项”，点击“转到”。
2. 若已有旧版 Progress Indicators，先取消勾选。点击“添加”，选中下载的 `.ppam`，按本机现有的宏授权流程启用。
3. 打开已有章节的文稿，进入“演讲进度条”选项卡，在“设计样式”中选择预设。设定停靠边和均长模式，点击“更新当前页”查看效果，再点“生成”应用到需要的页面。

更新旧文稿里的深红样式时，重新选取“深红层叠”再生成，文稿中已保存的导航才会采用新配方。

## 从源码构建

在仓库根目录打开 PowerShell，使用 Python 3.12 和 uv 生成维护文稿。

```powershell
uv venv --python 3.12 .venv
uv pip install --python .venv\Scripts\python.exe -r requirements-build.txt
.\.venv\Scripts\python.exe build_package.py --output-dir rebuild --version-token CN11_1 --report rebuild\package-validation.json
```

正式加载项由桌面 PowerPoint 从维护文稿另存为 `.ppam`。发布命令见[构建说明](BUILD_zh-CN.txt)。

VBA 和 Ribbon XML 位于 `source/`。预设目录位于 `styles.json`，对应的 VBA 配方在 `source/mdlProgressIndicator.bas` 的 `PI_StyleRecipe` 中。构建器使用 `original/` 下的上游模板保留窗体和资源。

## 许可

软件采用 MIT 许可证，上游功能区图标另有署名与使用条款。[软件许可证](LICENSE) · [第三方声明](THIRD_PARTY_NOTICES.md)
