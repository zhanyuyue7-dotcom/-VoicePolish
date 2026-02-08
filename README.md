# VoicePolish — 语音输入 + AI 润色工具

按一个快捷键开始说话，再按一次自动润色，结果直接粘贴到你正在打字的地方。

## 功能

- **一键语音输入**：`Ctrl+Alt+V` 触发 Windows 语音识别，自动打开临时记事本接收语音文字
- **AI 自动润色**：去除语气词（嗯、呃、啊），整理成通顺的书面语，保留原意
- **自动粘贴**：润色结果自动粘贴到你之前的窗口，同时留在剪贴板

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 配置 API

复制 `config.example.json` 为 `config.json`，填入你的 API 信息：

```json
{
  "api_base": "https://ark.cn-beijing.volces.com/api/v3",
  "api_key": "你的API Key",
  "model": "doubao-seed-1-8-251228"
}
```

### 3. 运行

```bash
python voice_polish.py
```

## 配置说明

`config.json` 包含三个字段：

| 字段 | 说明 |
|------|------|
| `api_base` | AI 服务的 API 地址 |
| `api_key` | 你的 API 密钥 |
| `model` | 使用的模型名称 |
