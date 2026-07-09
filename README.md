# 实训报告生成器

本地运行的实训报告生成工具。网页里拖入多个 TXT 笔记后，工具会按文件名中的“月-日”合并同一天内容，再基于 Word 模板生成实训报告、运行结果图和 `对照表.xlsx`。

## 功能

- 按 `月-日-次` 文件名分组，例如 `7-4-01.txt` 和 `7-4-02.txt` 会合并为同一天
- 自动生成课程目标、课程内容、课程内容详情、课程代码及执行过程
- TXT 有明确任务/练习时只提取明确任务；没有明确任务时，会从代码类语句、代码片段和代码操作步骤中提取任务；完全没有代码类内容才写提示语
- 本地规则加入段落顺序变化、句式变化、连接词变化、总结角度变化、问题分析角度变化、模块模板随机选择、个性化内容插入和 n-gram 重复率检测
- 自动生成 PyCharm 运行结果风格图片，并插入每个任务代码下方
- 生成 `对照表.xlsx`，三列为：项目、原txt文档、生成的word文档
- 可选通过后端真实调用 OpenAI-compatible API 进行 AI 增强
- API Key 可在网页输入，输入框会隐藏显示，并自动保存到本机 `.env`
- 新增 `/api/chat` 后端中转接口，可在网页里测试真实模型返回
- 网页采用 iOS 风格毛玻璃界面，TXT 输入区和产出日志区都支持内部滚动

## 启动

Windows 推荐双击：

```text
启动实训报告网页.bat
```

不要直接双击 `实训报告生成器.html`。如果浏览器地址是 `file:///.../实训报告生成器.html`，说明没有启动后端服务；正确地址应是 `http://127.0.0.1:8765/`。

手动启动：

```bash
pip install -r requirements.txt
python start_report_web.py
```

浏览器访问：

```text
http://127.0.0.1:8765/
```

## AI 配置

最方便的方式是在网页里勾选“AI 增强”，点击“套用接口预设”，然后把 API Key 粘贴到“API Key（隐藏输入）”框里。网页会先做格式检查，再把配置保存到本机 `.env`，后端会用这个配置真实调用 API。

也可以手动复制 `.env.example` 为 `.env`：

```text
AI_API_KEY=你的API密钥
AI_PROVIDER=deepseek
AI_BASE_URL=https://api.deepseek.com
AI_MODEL=deepseek-chat
AI_TEMPERATURE=0.6
AI_TIMEOUT=60
AI_PROXY_URL=
```

支持 DeepSeek、OpenRouter、Kimi、通义千问、智谱、硅基流动、豆包/火山方舟、MiniMax、零一万物以及自定义 OpenAI-compatible `/chat/completions` 接口。网页配置会自动写入 `.env`；点“清除 Key”会同时清空网页输入和 `.env` 里的 `AI_API_KEY`。

如果遇到 `WinError 10013`，先点“检测代理”，工具会自动寻找系统代理和常见 Clash/V2Ray 本地端口；再点“检测全部通道”，确认每个 API 地址的网络层是否可达。

## 测试 API

1. 网页勾选“AI 增强”。
2. 选择通道预设并点击“套用接口预设”。
3. 在“API Key（隐藏输入）”里粘贴真实 Key。
4. 点击“保存 AI 配置”或等待自动保存。
5. 如果网络受限，点击“检测代理”和“检测全部通道”。
6. 点击“测试 AI 连接”。
7. 在“AI 输入框”输入一句话，点击“发送到 AI”。
8. 如果成功，产出窗口会显示真实模型返回内容；生成报告时日志会显示是否真实使用 API。

## 主要文件

- `实训报告生成器.py`：核心 Word、内容、多样化和 API 调用逻辑
- `实训报告生成器网页.py`：本地网页后端，包含 `/generate`、`/test-ai`、`/api/chat`
- `实训报告生成器.html`：网页 UI
- `start_report_web.py`：ASCII 文件名启动中转，避免 bat 因中文文件名/编码失败
- `.env.example`：AI 环境变量示例
- `启动实训报告网页.bat`：Windows 双击启动脚本
