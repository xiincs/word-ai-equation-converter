# Word AI Equation Converter

A VBA tool for Microsoft Word that converts AI-generated LaTeX-style math expressions into native Word equations (OMath).

It is designed for workflows where users copy AI outputs (ChatGPT, Claude, etc.) containing `$...$` or `$$...$$` math expressions into Word, and need automatic conversion into professionally rendered Word equations.

------

## ✨ Features

- Converts inline math: `$...$`
- Converts display math: `$$...$$`
- Automatically builds Word Professional Equation format (OMath)
- Batch processes entire Word document
- Works directly on pasted AI-generated content

------

## 📌 Use Case

This tool is designed for:

- Copying AI-generated text into Microsoft Word
- Converting LaTeX-style math into Word-native equations
- Academic writing workflows
- Technical documentation editing

Example input from AI:

```
The equation is $E = mc^2$ and also $$\int_0^1 x^2 dx$$.
```

After running the macro:

- `$E = mc^2$` → Word Equation (Professional format)
- `$$\int_0^1 x^2 dx$$` → Display Equation

------

## 🚀 Installation

### 1. Open VBA Editor

In Word:

```
Alt + F11
```

------

### 2. Import Module

- Right-click on VBAProject

- Choose:

  ```
  Insert → Module
  ```

- Paste the code from:

  ```
  Dollar_To_Equation_Both.bas
  ```

------

### 3. Save Document

Save as:

```
Word Macro-Enabled Document (*.docm)
```

or install into:

```
Normal.dotm (global usage)
```

------

## ▶ Usage

### Method 1 (Recommended)

```
Alt + F8 → Dollar_To_Equation_Both → Run
```

------

### Method 2 (Developer Tab)

Assign macro to a button:

```
Developer → Insert → Button → Assign Macro
```

------

## ⚙ How It Works

The macro performs three steps:

1. Scan document for `$$...$$` blocks
2. Scan document for `$...$` inline math
3. Convert all detected ranges into Word OMath objects
4. Apply `BuildUp()` to render Professional equation format