```markdown
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

切换语言：[English](README.md)

# ⚡ LocalKB | 极速本地知识库
纯Python构建**离线本地知识库**，所有数据保存在本机，不上传云端，跨 Windows / Mac / Linux 全平台兼容。

## 📋 项目简介
- 多格式文档解析入库、本地FTS5全文检索
- 两种导入策略区分备份：**网页上传自动备份（并提取文本入库） / 目录导入不备份原文件（仅提取文本入库）**
- 双浏览视图：文件全文预览 / 分段分页浏览
- 检索结果一键导出 MD / TXT / Word

### ✅ 支持格式
`PDF / DOCX / PPTX / TXT / MD / HTML / HTM`
⚠️ 暂未兼容：.doc(97-2003)、.ppt(97-2003)、加密/扫描/特殊字体PDF

## 📌 备份规则（重要）
1. **网页上传文件**：自动原样备份至 `upload_files` 文件夹
2. **本地文件夹批量导入**：仅提取文本入库，**不复制原文件、不占用额外磁盘**

## 🛠 环境安装
```bash
pip install streamlit PyPDF2 python-docx python-pptx chardet olefile
```

## 🚀 启动项目

```bash
streamlit run knowledge_base.py
```

默认访问地址：`http://127.0.0.1:8501`

## 📖 使用说明

### 1. 侧边栏 - 文档入库

- **文件上传**：选中文件→确认入库→自动备份原件

- **文件夹导入**：选择目录→批量解析，**无原文件备份**

- **删除文档**：下拉选择，一键删除库内选定文件相关所有记录

### 2. 浏览模式

- **文件模式**：分页列表，点击打开完整原文

- **条目模式**：文档自动分段，分页逐条查看

### 3. 顶部常驻检索

页面顶部搜索框，关键词全文检索，结果支持三种格式导出。

## 📂 项目目录

```Plain Text
LocalKB/
├─ knowledge_base.py    # 主程序源码
├─ LICENSE                       # MIT开源协议
├─ README.zh-CN.md     # 中文文档
├─ README.md                # 英文文档
├─ knowledge.db             # 自动生成数据库
└─ upload_files                 # 仅上传文件自动备份目录
```

## 📄 开源协议

本项目基于 **MIT License** 开源，可自由商用、修改、分发。

Copyright (c) 2026 CCking2022

## ⚠️ 功能限制

程序运行稳定，所有异常已捕获、不会崩溃，以下格式尚未完成适配：
1. PDF：加密PDF、图片扫描PDF、内嵌特殊字体PDF会出现乱码或提取失败，普通文字PDF解析正常；
2. Office：二进制老格式 .doc / .ppt（Office97~2003）暂不支持，现代格式 .docx / .pptx 解析完好。

## 📌 后续规划

- 优化PDF解析，修复字体乱码；
- 集成 antiword + catdoc 实现老doc、ppt解析；
- 接入Tesseract-OCR，适配扫描版PDF。
- Excel / EPUB 格式支持
- 本地大模型语义检索
- 文档标签分类
- 一键打包 EXE 免环境运行

⭐ 觉得好用欢迎 Star 支持！

