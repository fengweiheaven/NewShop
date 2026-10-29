# 第三方组件声明

## JSZip 3.10.1

- 项目：[JSZip 官方仓库](https://github.com/Stuk/jszip)
- 分发文件：`dist/vendor/jszip.min.js`；单文件版 `standalone/OrderFlow.html` 也包含该组件。
- 用途：在浏览器中读取和生成 ZIP 容器，支持 `.xlsx` 文件的本地处理。
- 上游采用双许可：`MIT OR GPL-3.0-or-later`。本项目选择在 **MIT 许可证**下使用和分发 JSZip。
- 版本已依据本地 `jszip/package.json` 和所分发文件的版本头确认。原文件头部的版权与许可声明应继续保留。

以下为上游 `LICENSE.markdown` 中的完整 MIT 许可文本：

```text
The MIT License
===============

Copyright (c) 2009-2016 Stuart Knightley, David Duponchel, Franz Buchinger, António Afonso

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
```

JSZip 分发文件还声明使用 MIT 许可的 pako 组件，保留在原文件头部；其上游项目与许可见 [pako](https://github.com/nodeca/pako) 和 [pako LICENSE](https://github.com/nodeca/pako/blob/main/LICENSE)。
