<div align="center">

English | [简体中文](README.md)

</div>

![](https://fastly.jsdelivr.net/gh/Achuan-2/PicBed/assets/20250117115019-2025-01-17.png)

## 📄 Special Notes

* Any scripts involved in the `SlideSCI` project released by this repository are for testing and learning purposes only. Commercial use is strictly prohibited. We do not guarantee the legality, accuracy, completeness, or effectiveness of these scripts. Use them at your own discretion.

* All resource files in this project are prohibited from being reposted or published in any form by any public accounts or self-media platforms.

* `The author` shall not be held responsible for any issues arising from the scripts, including but not limited to losses or damages caused by script errors.

* Unauthorized use of any content from the `SlideSCI` project for commercial or illegal purposes is strictly prohibited. Violators will bear all consequences.

* Anyone who views this project or directly/indirectly uses any scripts from the `SlideSCI` project must read this disclaimer carefully. `The author` reserves the right to modify or supplement this disclaimer at any time. By using or copying any related scripts or the `SlideSCI` project, you are deemed to have accepted this disclaimer.

* This project follows the `AGPL-3.0 License`. If there is any conflict between this Special Note and the `AGPL-3.0 License`, this Special Note shall prevail.

> If you use or copy any code or projects created by this repository and `the author`, you are deemed to have `accepted` this disclaimer. Please read carefully.

> If you used or copied any code or projects created by this repository and `the author` before this disclaimer was issued and are still using them, you are deemed to have `accepted` this disclaimer. Please read carefully.

## 📝 Development Background

Does anyone else share my long-standing grievances with PowerPoint? 😡:

💔 **No Image Titles**: Unlike Word, you can't directly add titles to images. You have to manually insert text boxes and spend ages aligning them, only to end up with crooked results!

💔 **No Copy-Paste Element Positioning**: To keep similar elements in the same position across different slides, you have to copy-paste and modify each time. No way to copy-paste positions directly!

💔 **No Auto-Align for Images**: Insert multiple images and want them neatly arranged in rows and columns? Either drag each one manually for eternity or align them column by column horizontally and then vertically.

💔 **No Code Block Insertion**: You have to copy-paste from external editors (like VSCode) or specialized websites, or screenshot code blocks as images. So tedious!

💔 **No LaTeX Math Formula Support**: Nowadays, I rely on AI to recognize and generate math formulas in LaTeX format, which can't be directly pasted into PPT.

...

Most PPT plugins on the market are packed with flashy but impractical features. As a graduate student, I need to create clear, content-focused progress reports quickly every week—aesthetics are secondary.

With AI's help, I developed solutions for these pain points swiftly! The sense of accomplishment is real! (Over 99% of this plugin's code was AI-generated. Thank you, AI sensei!)

In the spirit of open source, this plugin is available on GitHub. Stars are appreciated! 🌟

GitHub: [https://github.com/Achuan-2/SlideSCI](https://github.com/Achuan-2/SlideSCI)

## ✨ Key Features

* **Batch Add Image Titles**: <u>Batch</u> select images and add centered captions below them. Supports auto-grouping images and titles.

  ![](https://fastly.jsdelivr.net/gh/Achuan-2/PicBed/assets/20250116004806-2025-01-16.png)
* **Auto-Arrange Images**: Automatically align multiple images. Set columns per row, column spacing, row spacing (defaults to column spacing), and image dimensions.

  * If width/height isn't set, the first image's height is used for alignment.
  * Enable "Arrange by Position" to auto-detect order based on manual placement. Otherwise, uses the selection order.

  ![](https://fastly.jsdelivr.net/gh/Achuan-2/PicBed/assets/20250116004816-2025-01-16.png)
* **Copy & Paste Element Positions**: Copy positions of multiple elements and paste them to others (supports multi-select!). Useful for aligning elements across slides or within a slide.

  ![](https://fastly.jsdelivr.net/gh/Achuan-2/PicBed/assets/复制粘贴位置-2025-01-17.gif)

  ![](https://fastly.jsdelivr.net/gh/Achuan-2/PicBed/assets/复制粘贴位置-2025-01-16.gif)
* **Copy & Paste Element Dimensions**: Quickly standardize image sizes via multi-select paste.
* **Insert Syntax-Highlighted Code Blocks**:

  ![](https://fastly.jsdelivr.net/gh/Achuan-2/PicBed/assets/20250116004856-2025-01-16.png)

  * **Supported Languages**: MATLAB, Python, JavaScript, HTML, CSS, C#.
  * **Toggle Black/White Background**: Default is black. Click "Code Black Background" to deactivate for white.
* **Insert LaTeX Math Formulas**:

  ![](https://fastly.jsdelivr.net/gh/Achuan-2/PicBed/assets/20250116004910-2025-01-16.png)
* **Insert Markdown Text**: Paste entire Markdown notes into PPT at once! Preserves original order!

  ![](https://fastly.jsdelivr.net/gh/Achuan-2/PicBed/assets/20250116004919-2025-01-16.png)

  * **Inline Formats**: Bold, underline, superscript, subscript, italic, links, inline code, inline math.
  * **Block Formats**: Headings, lists, code blocks, tables, math formulas, blockquotes.

    * **List Enhancements**:
      * Preserves hanging indents (lost in default HTML-to-PPT pasting).
      * Converts task lists to ☑ (completed) and ☐ (unchecked).
    * **Code Blocks**:
      * Independent text boxes with editable syntax highlighting (black/white themes).
    * **Tables**:
      * Limited to 500px width by default, with 1pt black borders.
    * **Math Formulas**:
      * Independent, editable text boxes.
    * **Blockquotes**:
      * Independent text boxes with black borders.

## 🪟 Supported Environments

Developed on Windows 11 using [Visual Studio Tools for Office](https://www.visualstudio.com/de/vs/office-tools/) and C#. Designed for Microsoft PowerPoint. Compatible with WPS (note: WPS does not support LaTeX formulas or Markdown insertion).

**Windows only**. No Mac support.

## 🖥️ Installation

1. Download `msi` file from the GitHub [Releases](https://github.com/Achuan-2/my_ppt_plugin/releases).
2. Extract and double-click to install.

**Note**: Close PowerPoint before installation. Otherwise, the plugin won't load immediately.

**Required Dependencies**:
- [Microsoft .NET Framework 4.0+](https://www.microsoft.com/zh-cn/download/details.aspx?id=17718)
- [Microsoft Visual Studio 2010 Tools for Office Runtime](https://www.microsoft.com/zh-cn/download/details.aspx?id=105522)

If the plugin fails to load (e.g., "Runtime error loading COM add-in"), install the dependencies above.

## ❓ FAQs

* **How to add plugin features to the Quick Access Toolbar?**  
  Right-click a button and select "Add to Quick Access Toolbar."  
  ![](https://fastly.jsdelivr.net/gh/Achuan-2/PicBed/assets/PixPin_2025-01-16_16-56-07-2025-01-16.png)  
  Move the Quick Access Toolbar below the ribbon for easier access.
* **Caption text boxes aren't centered or lack proper width?**  
  Set a default text box: Widen the caption, center it, then set it as the default.
* **LaTeX formulas display incorrectly?**  
  Best for single-line formulas. For complex multi-line formulas, use [IguanaTex](https://github.com/Jonathan-LeRoux/IguanaTex).  
  See examples of PPT-specific LaTeX syntax [here](https://github.com/Achuan-2/my_ppt_plugin/issues/7).

## ❤️ Support

If you enjoy this plugin, consider giving a ⭐ on GitHub or donating to fuel further development.  

![](https://fastly.jsdelivr.net/gh/Achuan-2/PicBed/assets/20241118182532-2024-11-18.png)

Donor list: https://www.yuque.com/achuan-2

## 👨‍💻 Feedback

Report issues via:
1. GitHub [Issues](https://github.com/Achuan-2/my_ppt_plugin/issues)
2. Email: [achuan-2@outlook.com](mailto:achuan-2@outlook.com)

## 🔍 Credits & References

* [jph00/latex-ppt](https://github.com/jph00/latex-ppt): LaTeX in PowerPoint support.
* [Markdig](https://github.com/xoofx/markdig): Markdown parsing.
* Thanks to Visual Studio Tools for Office for development tools.
* Gratitude to all users for suggestions and feedback.