# Excel Format Cleaner

一个本地桌面小工具，用来对现有 `.xlsx` 文件做“裁剪式清理”。

## 功能

- 按 `<<DNP` 和 `DNP>>` 边界筛选 worksheet，仅保留两者中间的页
- 仅保留打印区域内的内容
- 删除隐藏行和隐藏列
- 删除批注/备注
- 默认不保留图片等绘图对象
- 名称包含 `cover` 的 worksheet 会保留截图
- 将公式替换为静态值
- 无法取到公式结果时，写为空值
- 没有打印区域的 sheet 会跳过并提示
- 缺少 DNP 边界或出现多重边界时直接报错
- 输出新文件，不覆盖原文件

## 运行

推荐直接运行：

```bash
zsh run_app.sh
```

或者直接使用 Codex 工作区自带的 Python：

```bash
'/Users/shawppa/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3' app.py
```

如果你希望在普通 Python 环境运行，需要先安装：

```bash
pip install openpyxl
python3 app.py
```

## Windows 打包

Windows 用户不能直接使用 macOS 的 `.app`，需要单独生成 Windows 安装包。

项目已经包含 Windows 打包脚本和安装包配置：

- 构建脚本：`packaging/windows/build_windows.ps1`
- 安装包脚本：`packaging/windows/ExcelCleanerSetup.iss`
- 详细说明：[WINDOWS_BUILD.md](/Users/shawppa/Codex%20project/EXCEL%20PV/WINDOWS_BUILD.md)

在 Windows 机器上可以生成：

- 便携运行目录
- 安装包 `Excel Cleaner Setup v1.0.0.exe`

## 打包为 macOS `.app`

如果你想生成一个可双击启动的 macOS 应用包，可以运行：

```bash
zsh build_app.sh
```

生成后的应用位置：

```bash
dist/Excel Cleaner.app
```

## 输出位置

清理后的文件会生成在源文件同级目录下的 `cleaned_output` 文件夹中，文件名默认追加 `_cleaned`。
