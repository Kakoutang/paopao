# paopao

paopao 是咨询级 PPT 生成工具，能将 PDF、报告、研究资料转化为可编辑的专业 PowerPoint。

## 当前状态

**paopao 目前处于内测阶段，生成功能尚未开放。**

你现在安装的是 paopao 的预览壳——它可以帮你整理需求（页数、语言、重点），但暂时还不能生成 PPT。生成引擎正在内测中，开放后会通过本插件自动启用。

如果你想参加内测，请联系 paopao 团队：kakoutang@gmail.com

**请注意：** 在生成引擎启用之前，paopao 不会生成任何 PPT 文件。如果你看到"runtime 未启用"的提示，说明一切正常——生成功能还没有开放到你的工作区。

---

paopao is a consulting-grade PPT generation tool that turns PDFs, reports, and research into editable PowerPoint decks.

**paopao is currently in closed beta. Generation is not yet available.**

What you are installing is the paopao preview shell. It can help you prepare your deck request (page count, language, focus), but it cannot generate PPTs yet. The generation engine is in closed beta and will be enabled through this plugin when it becomes available.

To join the beta, contact the paopao team: kakoutang@gmail.com

**Note:** Until the generation engine is enabled, paopao will not produce any PPT files. If you see a "runtime not enabled" message, that is expected — generation has not been opened for your workspace yet.

## 加入内测 / Join Beta

发邮件到 kakoutang@gmail.com，说明你的使用场景，我们会尽快开通。

Email kakoutang@gmail.com with your use case and we will enable access as soon as possible.

## 隐私 / Privacy

你的源文件始终留在你的本地环境中，不会上传到任何服务器。

Your source documents stay in your local workspace and are never uploaded.

## 授权 / License Activation

如果你已收到内测授权码：

```bash
python3 scripts/paopao_auth.py activate --code "你的授权码"
python3 scripts/paopao_auth.py status
```
