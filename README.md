# paopao

paopao helps create editable PowerPoint decks from PDFs, papers, reports, and reference materials inside your local AI workspace.

## Welcome

Welcome to paopao.

- Output is editable `.pptx`.
- Works locally in your AI workspace.
- Quality is being updated frequently; if a result looks imperfect, try again later or reduce the page count for a cleaner first pass.

## Install

Install this repository as a Codex plugin, then mention paopao when asking for a deck.

Example:

```text
Use paopao to make a 5-page English PPT from this PDF.
```

Before generation, paopao may ask you to confirm:

- Page count
- Language
- Focus, use case, or audience

## Update

```text
请运行 paopao_run.py update，然后继续做 PPT。
```

Or in English:

```text
Run paopao_run.py update, then continue making the PPT.
```

The update command only downloads files that changed — it takes a few seconds.

## Included

- Codex plugin manifest
- Paopao PPT skill entry
- Local runtime scripts
- Runtime template index

## Template Access

paopao uses a managed template service to prepare slide layouts while keeping source templates protected.

## Privacy & Data Security / 隐私与数据安全

paopao 的数据处理方式与您日常使用的 AI 工具（如 ChatGPT、Claude、Gemini）类似：生成过程中，部分信息需要经过云端服务处理。

**什么会离开您的电脑：**
- 每页幻灯片的摘要信息（标题、关键数据点、结论等即将出现在 PPT 上的内容）会发送至 paopao 设计服务器，用于版式匹配和页面设计。
- 访问验证信息会在需要时发送至 paopao 服务。

**什么不会离开您的电脑：**
- 您的原始文件（PDF、文档、Excel 等）始终保留在本地。
- 生成的 PPT 文件始终保留在本地。

**我们不会：**
- 存储您发送的摘要数据。
- 将您的数据用于模型训练或其他用途。

**安全建议：** 与使用任何 AI 工具一样，如果您的资料包含高度机密信息（如未公开财务数据、商业秘密、受保密协议约束的内容），请在使用前评估相关风险，或咨询您所在机构的信息安全部门。

---

paopao handles data similarly to AI tools you already use (ChatGPT, Claude, Gemini): some information is processed through a cloud service during generation.

**What leaves your computer:**
- Per-slide summary data (titles, key figures, conclusions - content that will appear on the slides) is sent to the paopao design server for layout matching.
- Access verification data is sent to the paopao service when needed.

**What stays on your computer:**
- Your source files (PDFs, documents, spreadsheets) never leave your machine.
- Your generated PPTX files never leave your machine.

**We do not:**
- Store your summary data after processing.
- Use your data for model training or any other purpose.

**Security advice:** As with any AI-powered tool, if your materials contain highly confidential information (e.g., unpublished financials, trade secrets, or NDA-protected content), please assess the risks before use, or consult your organization's information security team.

## Support

For feedback or access issues, contact WeChat: `sugarong_`.
