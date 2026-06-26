# Word AI Equation Converter（Word AI 公式转换工具）

一个用于 Microsoft Word 的 VBA 宏工具，用于将 AI（如 ChatGPT）生成文本中的 LaTeX 数学公式自动转换为 Word 原生公式（OMath 专业模式）。

------

## 项目用途

当你从 AI（ChatGPT / Claude / DeepSeek 等）复制内容到 Word 时，经常会包含如下格式：

```
这是一个公式 $E = mc^2$，以及一个展示公式 $$\int_0^1 x^2 dx$$。
```

Word 默认无法正确识别 LaTeX，需要手动转换。

本工具可以自动完成：

- `$...$` → Word 行内公式
- `$$...$$` → Word 独立公式（块级公式）
- 自动转换为 **专业公式模式（Professional OMath）**

------

## 功能特性

- 支持 AI 输出的 LaTeX 公式转换
- 自动识别 `$...$` 行内公式
- 自动识别 `$$...$$` 块级公式
- 批量处理整个 Word 文档
- 自动应用 Word 专业公式排版（BuildUp）

------

## 安装方法

### 方法一：导入 VBA 模块（推荐）

1. 打开 Word

2. 按 `Alt + F11` 打开 VBA 编辑器

3. 选择：

   ```
   插入 → 模块（Module）
   ```

4. 导入或粘贴文件：

   ```
   Dollar_To_Equation_Both.bas
   ```

------

### 方法二：保存到全局模板（高级）

将代码保存到：

```
Normal.dotm
```

即可在所有 Word 文档中使用。

------

## 使用方法

### 方法一：快捷运行

```
Alt + F8 → Dollar_To_Equation_Both → 运行
```

------

### 方法二：按钮运行（推荐）

启用开发工具后：

```
开发工具 → 插入 → 按钮 → 绑定宏
```

------

## 工作原理

本工具执行以下流程：

### 1️ 扫描文档

查找：

- `$...$`
- `$$...$$`

### 2️ 转换为 Word 公式对象

使用：

- `OMaths.Add`

### 3️ 渲染为专业模式

执行：

- `BuildUp()`

最终输出 Word 原生公式（Professional Equation）