"""
VoicePolish — 语音输入 + LLM 润色工具

用法：
  1. 运行此脚本（后台常驻）
  2. 按 Ctrl+Alt+V 开始语音输入（自动弹出临时记事本 + Win+H）
  3. 对着麦克风说话
  4. 再按 Ctrl+Alt+V 停止并自动润色
  5. 润色结果自动粘贴到你原来的窗口，并留在剪贴板
  6. 按 Ctrl+C 退出程序
"""

import json
import time
import threading
import subprocess
import tempfile
import sys
import os
import ctypes

import pyperclip
import pyautogui
import keyboard as kb
from openai import OpenAI

# --- Windows 窗口管理 ---
user32 = ctypes.windll.user32

# --- 配置 ---
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

with open(CONFIG_PATH, "r", encoding="utf-8") as f:
    config = json.load(f)

client = OpenAI(
    api_key=config["api_key"],
    base_url=config["api_base"],
)
MODEL = config["model"]

POLISH_PROMPT = """你是一个语音输入文本润色助手。用户通过语音输入了一段文字，请你：
1. 去除语气词（嗯、呃、啊、哦、额等），但保留有实际语义的词
2. 修复不自然的空格和间距
3. 将断断续续的片段整理成通顺的句子
4. 保持原意，不要添加或改变内容
5. 如果是代码相关的讨论，保留技术术语的准确性
6. 只输出润色后的文本，不要任何解释"""

# --- 状态 ---
class State:
    IDLE = "idle"
    RECORDING = "recording"
    POLISHING = "polishing"

state = State.IDLE
state_lock = threading.Lock()
recording_start_time = 0
target_hwnd = None       # 用户原来的窗口句柄
notepad_proc = None      # 临时记事本进程
notepad_hwnd = None      # 临时记事本窗口句柄
temp_file_path = None    # 临时文件路径


def get_foreground_window():
    """获取当前前台窗口句柄"""
    return user32.GetForegroundWindow()


def focus_window(hwnd):
    """将指定窗口带到前台"""
    if hwnd:
        user32.SetForegroundWindow(hwnd)
        time.sleep(0.3)


def get_window_pid(hwnd):
    """通过窗口句柄获取实际进程 PID"""
    pid = ctypes.c_ulong()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return pid.value


def wait_for_keys_release():
    """等待所有修饰键松开，避免与 pyautogui 模拟按键冲突"""
    modifiers = ["ctrl", "alt", "shift", "win"]
    for _ in range(50):  # 最多等 2.5 秒
        if not any(kb.is_pressed(m) for m in modifiers):
            return
        time.sleep(0.05)
    print("[VoicePolish] 警告: 修饰键未完全松开，继续执行...")


def polish_text(raw_text):
    """调用 LLM 润色文本"""
    raw_text = raw_text.strip()
    if not raw_text:
        return ""

    try:
        response = client.chat.completions.create(
            model=MODEL,
            messages=[
                {"role": "system", "content": POLISH_PROMPT},
                {"role": "user", "content": raw_text},
            ],
            temperature=0.3,
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"[VoicePolish] LLM 调用失败: {e}")
        return raw_text  # 失败时返回原文

def _clear_notepad_sessions():
    """清除 Windows 11 记事本的会话恢复数据，防止旧标签页弹出"""
    tab_state_dir = os.path.join(
        os.environ.get("LOCALAPPDATA", ""),
        "Packages", "Microsoft.WindowsNotepad_8wekyb3d8bbwe",
        "LocalState", "TabState"
    )
    if os.path.isdir(tab_state_dir):
        for f in os.listdir(tab_state_dir):
            try:
                os.unlink(os.path.join(tab_state_dir, f))
            except Exception:
                pass


def start_recording():
    """开始语音输入：打开临时记事本，触发 Win+H"""
    global state, recording_start_time, target_hwnd, notepad_proc, notepad_hwnd, temp_file_path
    print("[VoicePolish] 开始语音输入...")

    # 1. 记住用户当前的窗口（稍后润色结果要粘贴回这里）
    target_hwnd = get_foreground_window()
    print(f"[VoicePolish] 目标窗口句柄: {target_hwnd}")

    # 2. 清除记事本会话恢复数据，然后打开干净的临时记事本
    _clear_notepad_sessions()
    tmp = tempfile.NamedTemporaryFile(suffix=".txt", prefix="vp_", delete=False)
    temp_file_path = tmp.name
    tmp.close()
    notepad_proc = subprocess.Popen(["notepad.exe", temp_file_path])
    time.sleep(0.8)  # 等待记事本完全打开并获得焦点
    notepad_hwnd = get_foreground_window()  # 记住记事本的窗口句柄
    print(f"[VoicePolish] 记事本窗口句柄: {notepad_hwnd}")

    # 3. 触发 Win+H 开始语音输入（语音文字会打到记事本里）
    recording_start_time = time.time()
    wait_for_keys_release()
    pyautogui.hotkey("win", "h")
    time.sleep(0.5)

    state = State.RECORDING
    print("[VoicePolish] 请对着麦克风说话，说完后再按一次快捷键...")


def stop_and_polish():
    """停止语音输入：从记事本获取文字，润色后粘贴到原窗口"""
    global state, notepad_proc
    state = State.POLISHING
    print("[VoicePolish] 停止语音输入，开始润色...")

    # 确保至少录音了 0.5 秒
    elapsed = time.time() - recording_start_time
    if elapsed < 0.5:
        print("[VoicePolish] 录音时间太短，取消操作")
        _cleanup_notepad()
        state = State.IDLE
        return

    # 1. 关闭 Win+H
    wait_for_keys_release()
    pyautogui.hotkey("win", "h")
    time.sleep(0.8)

    # 2. 确保焦点在记事本上，然后获取文字（Ctrl+A 全选 → Ctrl+C 复制）
    focus_window(notepad_hwnd)
    wait_for_keys_release()
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.2)
    pyautogui.hotkey("ctrl", "c")
    time.sleep(0.2)
    raw_text = pyperclip.paste().strip()
    pyperclip.copy("")  # 立即清空剪贴板，避免原始文本残留
    print(f"[VoicePolish] 捕获到的文本长度: {len(raw_text)}, 内容: '{raw_text[:100]}'")

    if not raw_text:
        print("[VoicePolish] 没有检测到语音文字，保留记事本窗口供手动复制")
        state = State.IDLE
        return

    print(f"[VoicePolish] 原始文本: {raw_text}")

    # 3. 调用 LLM 润色
    polished = polish_text(raw_text)
    print(f"[VoicePolish] 润色结果: {polished}")

    # 4. 润色结果放入剪贴板
    pyperclip.copy(polished)

    # 5. 恢复焦点到用户原来的窗口，并粘贴
    focus_window(target_hwnd)
    time.sleep(0.2)
    wait_for_keys_release()
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.2)

    # 6. 再次确保剪贴板是润色结果（防止粘贴过程中被覆盖）
    pyperclip.copy(polished)

    # 7. 成功后才关闭记事本
    _cleanup_notepad()

    # 润色结果留在剪贴板，方便用户再次粘贴
    state = State.IDLE
    print("[VoicePolish] 完成! 润色结果已留在剪贴板")


def _cleanup_notepad():
    """强制关闭临时记事本并删除临时文件"""
    global notepad_proc, notepad_hwnd, temp_file_path

    # 通过窗口句柄获取真实 PID（Windows 11 UWP 记事本的 PID 和 Popen 的不同）
    real_pid = None
    if notepad_hwnd:
        try:
            real_pid = get_window_pid(notepad_hwnd)
        except Exception:
            pass

    # 方法1: 用真实 PID 强杀（最可靠）
    if real_pid:
        try:
            subprocess.run(
                ["taskkill", "/F", "/PID", str(real_pid)],
                capture_output=True, timeout=3
            )
        except Exception:
            pass

    # 方法2: 用 Popen 的 PID 强杀（兜底）
    if notepad_proc:
        try:
            subprocess.run(
                ["taskkill", "/F", "/PID", str(notepad_proc.pid), "/T"],
                capture_output=True, timeout=3
            )
        except Exception:
            pass
        try:
            notepad_proc.kill()
        except Exception:
            pass
        notepad_proc = None

    notepad_hwnd = None
    time.sleep(0.3)

    # 删除临时文件
    if temp_file_path:
        try:
            os.unlink(temp_file_path)
        except Exception:
            pass
        temp_file_path = None


def on_hotkey():
    """快捷键回调"""
    global state
    with state_lock:
        current_state = state

    if current_state == State.IDLE:
        threading.Thread(target=start_recording, daemon=True).start()
    elif current_state == State.RECORDING:
        threading.Thread(target=stop_and_polish, daemon=True).start()
    # POLISHING 状态下忽略按键


# --- 主程序 ---
def main():
    print("=" * 60)
    print("  VoicePolish — 语音输入 + LLM 润色")
    print("  快捷键: Ctrl+Alt+V")
    print("  按一次开始说话，再按一次自动润色")
    print("  按 Ctrl+C 退出")
    print("=" * 60)

    # 移除 suppress=True 以避免键盘冲突
    kb.add_hotkey("ctrl+alt+v", on_hotkey, suppress=False)
    print("[VoicePolish] 快捷键已注册，等待触发...")

    try:
        kb.wait()  # 阻塞主线程，直到程序退出
    except KeyboardInterrupt:
        print("\n[VoicePolish] 程序已退出")
        sys.exit(0)


if __name__ == "__main__":
    main()
