import tkinter as tk
from tkinter import ttk, messagebox
import mido
import threading
import time
import pandas as pd
import pygame.midi  # 播放声音

# 初始化 pygame.midi 播放器
pygame.midi.init()
player = pygame.midi.Output(0)  # 默认设备ID为0
player.set_instrument(0)  # 0 = Acoustic Grand Piano

# MIDI 按键映射
KEYS_TO_MONITOR = {note: f"按键{note - 47}" for note in range(48, 73)}
data_log = []

root = tk.Tk()
root.title("MIDI 上位机显示")
root.configure(bg="#1e1e1e")

FONT_LABEL = ("微软雅黑", 12, "bold")
FONT_ENTRY = ("Consolas", 12, "bold")

# 左侧日志区
frame_left = ttk.Frame(root)
frame_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

log_label = ttk.Label(frame_left, text="MIDI事件记录", font=FONT_LABEL)
log_label.pack()

log_text = tk.Text(frame_left, height=30, width=50, bg="#121212", fg="#e0e0e0", insertbackground="white",
                   font=FONT_ENTRY, relief=tk.FLAT)
log_text.pack()

# 右侧按键区
frame_right_outer = ttk.Frame(root)
frame_right_outer.pack(side=tk.RIGHT, fill=tk.Y, padx=10, pady=10)

canvas = tk.Canvas(frame_right_outer, width=400, height=600, bg="#1e1e1e", highlightthickness=0)
scrollable_frame = ttk.Frame(frame_right_outer)
scrollable_frame.pack(fill=tk.BOTH, expand=True)

status_vars = {}

def style_entry(entry, active=False):
    entry.configure(background="#444444", foreground="#cccccc", font=FONT_ENTRY)

cols = 5
for idx, (note, label) in enumerate(KEYS_TO_MONITOR.items()):
    row = idx // cols
    col = idx % cols
    lbl = tk.Label(scrollable_frame, text=label, font=FONT_LABEL, fg="#4caf50", bg="#1e1e1e")
    lbl.grid(row=row*2, column=col, padx=8, pady=2)
    var = tk.StringVar(value="0")
    entry = tk.Entry(scrollable_frame, textvariable=var, width=8, justify="center", relief=tk.FLAT)
    entry.grid(row=row*2+1, column=col, padx=8, pady=2)
    style_entry(entry, active=False)
    status_vars[note] = (var, entry)

def export_to_excel():
    if not data_log:
        messagebox.showinfo("提示", "没有数据可导出")
        return
    df = pd.DataFrame(data_log)
    df.to_excel("midi_output.xlsx", index=False)
    messagebox.showinfo("成功", "数据已导出为 midi_output.xlsx")

export_button = tk.Button(scrollable_frame, text="导出为Excel", font=FONT_LABEL, bg="#3f51b5", fg="white",
                          activebackground="#6573c3", relief=tk.FLAT, command=export_to_excel)
export_button.grid(row=(len(KEYS_TO_MONITOR)//cols + 1)*2, column=0, columnspan=cols, pady=20, sticky="ew")

def listen_to_midi():
    try:
        input_name = mido.get_input_names()[0]
        with mido.open_input(input_name) as inport:
            for msg in inport:
                if msg.type == 'note_on' and msg.velocity > 0:
                    timestamp = time.strftime("%H:%M:%S", time.localtime())
                    note = msg.note
                    velocity = msg.velocity

                    log_text.insert(tk.END, f"{timestamp} 按键{note - 47}: 力度 {velocity}\n")
                    log_text.see(tk.END)

                    data_log.append({
                        "时间": timestamp,
                        "按键": f"按键{note - 47}",
                        "力度": velocity
                    })

                    if note in status_vars:
                        var, entry = status_vars[note]
                        var.set(str(velocity))
                        style_entry(entry, active=True)

                    # 播放音符
                    player.note_on(note, 127)

                elif msg.type == 'note_off' or (msg.type == 'note_on' and msg.velocity == 0):
                    note = msg.note
                   # if note in status_vars:
                       # var, entry = status_vars[note]
                       # var.set("0")
                       # style_entry(entry, active=False)

                    # 停止音符
                    player.note_off(note, 0)

    except Exception as e:
        messagebox.showerror("错误", f"MIDI 打开失败: {e}")

# 清理资源
def on_closing():
    player.close()
    pygame.midi.quit()
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)

# 启动线程监听 MIDI
threading.Thread(target=listen_to_midi, daemon=True).start()

root.mainloop()
