# 第三方组件与许可证说明

文枢文档转换器使用了以下开源组件。各组件版权归其作者所有。

## 直接依赖

| 组件 | 版本 | 许可证 | 用途 |
| --- | --- | --- | --- |
| PySide6 / Qt for Python | 6.10.3 | LGPL-3.0 / GPL-3.0 / 商业许可 | 桌面界面 |
| comtypes | 1.4.17 | MIT | 调用本机 WPS/Word 的 COM 接口 |
| pdf2docx | 0.5.13 | MIT | PDF → DOCX（保留版面模式） |
| PyMuPDF | 2.1 | AGPL-3.0 或商业许可 | PDF 解析与渲染 |
| python-docx | 1.2.0 | MIT | DOCX 读写 |
| OpenCV (headless) | 5.0.0 | Apache-2.0 | 图像版面分析 |
| NumPy | 2.4 | BSD-3-Clause | 数值计算与图像透明通道合成 |
| fontTools | 4.68 | MIT | 字体解析 |
| lxml | 6.1.3 | BSD-3-Clause | XML 处理 |

## 可选的本机软件

- **WPS Office / Microsoft Word**：程序不会分发这两款软件，只在本机已安装时通过系统 COM 接口调用它们导出 PDF。它们各自受其商业许可约束。
- **LibreOffice**：发布包内置的 LibreOffice Kit 采用 MPL-2.0；系统安装版本同样采用 MPL-2.0。

## 内置转换引擎

| 组件 | 版本 | 许可证 |
| --- | --- | --- |
| LibreOffice Kit （含 LibreOffice） | 0.1.5 / 引擎内置 | MPL-2.0 |
| Node.js | 24.21 | MIT |
| @deepseek-ai/libreoffice-kit | 0.1.5 | MPL-2.0 |

内置引擎的许可证和第三方声明位于：

- engine\node_modules\@deepseek-ai\libreoffice-kit\LICENSE
- engine\node_modules\@deepseek-ai\libreoffice-kit-win32-x64\licenses\

## 构建期组件

- PyInstaller 6.22.3（GPL-2.0-or-later，带启动器例外）
- Pillow（HPND）

## 重要提示

1. **PyMuPDF 采用 AGPL-3.0（或商业许可）**。本程序以源码 + 可执行文件的形式提供；如果你修改并对外分发本程序，需要遵守 AGPL-3.0 的源码开放要求，或向 Artifex 购买商业许可。仅在自己电脑上使用不受影响。
2. **PySide6 / Qt** 采用 LGPL-3.0。发布时请保留本说明、Qt 许可证文本与动态链接方式，或遵循 LGPL 的其他合规方式。
3. **LibreOffice** 采用 MPL-2.0。内置引擎中保留了其许可证、NOTICE 与源码对应说明。
4. 其他传递依赖（如 koffi、fflate、fontkit、saxes、brotli、unicode-properties 等）的许可证文件保存在 engine\node_modules 对应包目录中。

## 参考链接

- pdf2docx: https://github.com/ArtifexSoftware/pdf2docx
- PyMuPDF: https://github.com/pymupdf/PyMuPDF
- PySide6: https://doc.qt.io/qtforpython-6/licenses.html
- LibreOffice: https://www.libreoffice.org/about-us/licenses/
- Node.js: https://github.com/nodejs/node/blob/main/LICENSE
- OpenCV: https://github.com/opencv/opencv/blob/4.x/LICENSE
- NumPy: https://numpy.org/doc/stable/license.html
- python-docx: https://github.com/python-openxml/python-docx
- lxml: https://github.com/lxml/lxml
- fontTools: https://github.com/fonttools/fonttools

最后更新：2026-11-03
