# paopao

paopao helps create editable PowerPoint decks from PDFs, papers, reports, and reference materials inside your local AI workspace.

## Public Preview

This public preview is intended for early feedback.

- Up to 10 slides per deck.
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

paopao is updated frequently during the public preview. If Codex says paopao needs an update, or if free generation unexpectedly asks for a license, ask Codex:

```text
请帮我更新 paopao 插件到最新版，然后重新开始生成 PPT。
```

After updating, start the deck request again. The free preview should allow up to 10 slides and access the free layout catalog without a license.

## Included

- Codex plugin manifest
- Paopao PPT skill entry
- Local runtime scripts
- Public preview template index

## Prompt Template Protection

Raw paopao prompt templates are not downloaded to your computer. The local runtime receives a template catalog with allowed layout IDs and fill zones, then calls the paopao design server to assemble a filled prompt for the selected slide. The server returns only the filled result for that slide, not the reusable source template.

## Privacy & Data Security / 隐私与数据安全

paopao 的数据处理方式与您日常使用的 AI 工具（如 ChatGPT、Claude、Gemini）类似：生成过程中，部分信息需要经过云端服务处理。

**什么会离开您的电脑：**
- 每页幻灯片的摘要信息（标题、关键数据点、结论等即将出现在 PPT 上的内容）会发送至 paopao 设计服务器，用于版式匹配和页面设计。
- 许可证验证信息（激活码）会发送至许可证服务器。

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
- License activation data is sent to the license server.

**What stays on your computer:**
- Your source files (PDFs, documents, spreadsheets) never leave your machine.
- Your generated PPTX files never leave your machine.

**We do not:**
- Store your summary data after processing.
- Use your data for model training or any other purpose.

**Security advice:** As with any AI-powered tool, if your materials contain highly confidential information (e.g., unpublished financials, trade secrets, or NDA-protected content), please assess the risks before use, or consult your organization's information security team.

## Limits

The public preview is intentionally limited. It may not match the full commercial version on page count, template variety, or reconstruction quality.

## Support

For feedback or access issues, contact WeChat: `sugarong_`.
