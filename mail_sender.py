import csv
import json
import os
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import requests

# 1. 外部持久化配置路径（存放在 exe 所在文件夹）
if getattr(sys, "frozen", False):
  APP_DIR = os.path.dirname(sys.executable)
  # 打包后内部静态资源解压目录
  RESOURCE_DIR = getattr(sys, "_MEIPASS", APP_DIR)
else:
  APP_DIR = os.path.dirname(os.path.abspath(__file__))
  RESOURCE_DIR = APP_DIR

CONFIG_FILE = os.path.join(APP_DIR, "config.json")
ICON_PATH = os.path.join(RESOURCE_DIR, "app.ico")

# 接口端点映射字典
CHANNELS = {
    "别名发送 (Alias)": {
        "url": "https://ocloude.com/user/api/v1/alias/send",
        "label": "发件别名 (from):",
    },
    "Outlook.ba 发送 (BA)": {
        "url": "https://ocloude.com/user/api/v1/ba/send",
        "label": "Outlook.ba 邮箱 (from):",
    },
}


class MailSenderApp:

  def __init__(self, root):
    self.root = root
    self.root.title("邮件批量发送助手 - Win10版")
    self.root.geometry("820x760")
    self.root.minsize(720, 680)

    # 设置窗口左上角及任务栏图标
    if os.path.exists(ICON_PATH):
      try:
        self.root.iconbitmap(ICON_PATH)
      except Exception:
        pass

    self.is_sending = False
    self.stop_requested = False

    self._build_ui()
    self.load_config()

    # 退出程序时自动保存一次配置
    self.root.protocol("WM_DELETE_WINDOW", self.on_close)

  def _build_ui(self):
    # 顶部配置区
    config_frame = ttk.LabelFrame(self.root, text="API 与发件配置", padding=10)
    config_frame.pack(fill=tk.X, padx=10, pady=5)

    # 1. 发送渠道选择下拉框
    ttk.Label(config_frame, text="发送渠道:").grid(
        row=0, column=0, sticky=tk.W, pady=3
    )
    self.channel_var = tk.StringVar(value=list(CHANNELS.keys())[0])
    self.channel_combo = ttk.Combobox(
        config_frame,
        textvariable=self.channel_var,
        values=list(CHANNELS.keys()),
        state="readonly",
        width=30,
    )
    self.channel_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=3)
    self.channel_combo.bind("<<ComboboxSelected>>", self._on_channel_change)

    # 手动保存配置按钮
    save_cfg_btn = ttk.Button(
        config_frame, text="保存当前配置", command=self.save_config_manual
    )
    save_cfg_btn.grid(row=0, column=2, padx=5, pady=3, sticky=tk.E)

    # 2. Token
    ttk.Label(config_frame, text="API Token:").grid(
        row=1, column=0, sticky=tk.W, pady=3
    )
    self.token_entry = ttk.Entry(config_frame, width=55, show="*")
    self.token_entry.grid(
        row=1, column=1, columnspan=2, sticky=tk.EW, padx=5, pady=3
    )

    # 3. 发件人 (from)
    self.from_label = ttk.Label(
        config_frame, text=CHANNELS[self.channel_var.get()]["label"]
    )
    self.from_label.grid(row=2, column=0, sticky=tk.W, pady=3)
    self.from_entry = ttk.Entry(config_frame, width=55)
    self.from_entry.grid(
        row=2, column=1, columnspan=2, sticky=tk.EW, padx=5, pady=3
    )

    # 4. 可选的 Reply-To
    ttk.Label(config_frame, text="回复邮箱 (reply-to, 可选):").grid(
        row=3, column=0, sticky=tk.W, pady=3
    )
    self.reply_to_entry = ttk.Entry(config_frame, width=55)
    self.reply_to_entry.grid(
        row=3, column=1, columnspan=2, sticky=tk.EW, padx=5, pady=3
    )

    config_frame.columnconfigure(1, weight=1)

    # 邮件内容区
    content_frame = ttk.LabelFrame(self.root, text="邮件内容设置", padding=10)
    content_frame.pack(fill=tk.BOTH, expand=False, padx=10, pady=5)

    ttk.Label(content_frame, text="邮件主题:").grid(
        row=0, column=0, sticky=tk.W, pady=2
    )
    self.subject_entry = ttk.Entry(content_frame)
    self.subject_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=2)

    ttk.Label(content_frame, text="正文 (支持HTML):").grid(
        row=1, column=0, sticky=tk.NW, pady=2
    )
    self.body_text = tk.Text(content_frame, height=6)
    self.body_text.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=2)
    self.body_text.insert(tk.END, "<p>这是一封测试邮件。</p>")

    content_frame.columnconfigure(1, weight=1)

    # 收件人与发送策略控制区
    control_frame = ttk.LabelFrame(
        self.root, text="收件人与发送策略", padding=10
    )
    control_frame.pack(fill=tk.X, padx=10, pady=5)

    btn_box = ttk.Frame(control_frame)
    btn_box.pack(fill=tk.X, pady=2)

    ttk.Button(
        btn_box, text="导入收件人名单 (TXT/CSV)", command=self.load_recipients
    ).pack(side=tk.LEFT, padx=5)
    self.recipient_count_label = ttk.Label(btn_box, text="已加载: 0 个收件人")
    self.recipient_count_label.pack(side=tk.LEFT, padx=10)

    ttk.Label(btn_box, text="每封间隔(秒):").pack(side=tk.LEFT, padx=5)
    self.interval_spinbox = ttk.Spinbox(btn_box, from_=0.1, to=60.0, width=5)
    self.interval_spinbox.set(1.0)
    self.interval_spinbox.pack(side=tk.LEFT)

    self.start_btn = ttk.Button(
        btn_box, text="开始群发", command=self.start_sending
    )
    self.start_btn.pack(side=tk.RIGHT, padx=5)
    self.stop_btn = ttk.Button(
        btn_box, text="停止", command=self.stop_sending, state=tk.DISABLED
    )
    self.stop_btn.pack(side=tk.RIGHT, padx=5)

    # 进度条
    self.progress = ttk.Progressbar(control_frame, orient=tk.HORIZONTAL)
    self.progress.pack(fill=tk.X, pady=8)

    # 日志输出区
    log_frame = ttk.LabelFrame(self.root, text="运行日志", padding=10)
    log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    self.log_text = tk.Text(
        log_frame, state=tk.DISABLED, bg="#1e1e1e", fg="#fff"
    )
    self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar = ttk.Scrollbar(
        log_frame, orient=tk.VERTICAL, command=self.log_text.yview
    )
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    self.log_text.config(yscrollcommand=scrollbar.set)

    self.recipients = []

  def _on_channel_change(self, event=None):
    selected_channel = self.channel_var.get()
    label_text = CHANNELS[selected_channel]["label"]
    self.from_label.config(text=label_text)
    self.log(
        f"切换通道 -> {selected_channel} (Endpoint:"
        f" {CHANNELS[selected_channel]['url']})"
    )

  def log(self, message):
    self.log_text.config(state=tk.NORMAL)
    timestamp = time.strftime("[%H:%M:%S] ")
    self.log_text.insert(tk.END, timestamp + message + "\n")
    self.log_text.see(tk.END)
    self.log_text.config(state=tk.DISABLED)

  # ================= 配置读写逻辑 =================
  def load_config(self):
    """启动时从 config.json 读取配置并回显到界面"""
    if not os.path.exists(CONFIG_FILE):
      return

    try:
      with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)

      if "channel" in data and data["channel"] in CHANNELS:
        self.channel_var.set(data["channel"])
        self._on_channel_change()

      if "token" in data:
        self.token_entry.delete(0, tk.END)
        self.token_entry.insert(0, data["token"])

      if "from_alias" in data:
        self.from_entry.delete(0, tk.END)
        self.from_entry.insert(0, data["from_alias"])

      if "reply_to" in data:
        self.reply_to_entry.delete(0, tk.END)
        self.reply_to_entry.insert(0, data["reply_to"])

      if "interval" in data:
        self.interval_spinbox.set(str(data["interval"]))

      if "subject" in data:
        self.subject_entry.delete(0, tk.END)
        self.subject_entry.insert(0, data["subject"])

      self.log("已自动恢复上次保存的本地配置")
    except Exception as e:
      self.log(f"读取本地配置失败: {str(e)}")

  def save_config(self):
    """保存配置到 config.json"""
    data = {
        "channel": self.channel_var.get(),
        "token": self.token_entry.get().strip(),
        "from_alias": self.from_entry.get().strip(),
        "reply_to": self.reply_to_entry.get().strip(),
        "interval": self.interval_spinbox.get().strip(),
        "subject": self.subject_entry.get().strip(),
    }
    try:
      with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
      return True
    except Exception as e:
      self.log(f"保存配置异常: {str(e)}")
      return False

  def save_config_manual(self):
    if self.save_config():
      messagebox.showinfo("成功", "配置已成功保存到本地 (config.json)！")

  def on_close(self):
    """窗口关闭时的处理"""
    if self.is_sending:
      if not messagebox.askyesno(
          "退出确认", "群发任务正在执行中，确定要强行退出吗？"
      ):
        return
    self.save_config()
    self.root.destroy()

  # =================================================

  def load_recipients(self):
    filepath = filedialog.askopenfilename(
        title="选择收件人文件",
        filetypes=[
            ("Text/CSV Files", "*.txt *.csv"),
            ("All Files", "*.*"),
        ],
    )
    if not filepath:
      return

    emails = set()
    try:
      with open(filepath, "r", encoding="utf-8-sig") as f:
        if filepath.endswith(".csv"):
          reader = csv.reader(f)
          for row in reader:
            for item in row:
              val = item.strip()
              if "@" in val and "." in val:
                emails.add(val)
        else:
          for line in f:
            val = line.strip()
            if "@" in val and "." in val:
              emails.add(val)

      self.recipients = list(emails)
      self.recipient_count_label.config(
          text=f"已加载: {len(self.recipients)} 个有效收件人"
      )
      self.log(f"成功加载 {len(self.recipients)} 个收件人地址")
    except Exception as e:
      messagebox.showerror("读取错误", f"解析收件人文件失败: {str(e)}")

  def start_sending(self):
    channel = self.channel_var.get()
    endpoint = CHANNELS[channel]["url"]
    token = self.token_entry.get().strip()
    from_email = self.from_entry.get().strip()
    reply_to = self.reply_to_entry.get().strip()
    subject = self.subject_entry.get().strip()
    body = self.body_text.get("1.0", tk.END).strip()

    if not token or not from_email or not subject or not body:
      messagebox.showwarning(
          "配置缺失", "请先填齐 Token、发件邮箱/别名、邮件主题及正文！"
      )
      return

    if not self.recipients:
      messagebox.showwarning("缺少收件人", "请先导入收件人列表！")
      return

    try:
      interval = float(self.interval_spinbox.get())
    except ValueError:
      messagebox.showwarning("格式错误", "发送间隔必须为有效数字！")
      return

    # 点击开始群发时自动同步保存一次配置
    self.save_config()

    self.is_sending = True
    self.stop_requested = False
    self.start_btn.config(state=tk.DISABLED)
    self.stop_btn.config(state=tk.NORMAL)
    self.channel_combo.config(state=tk.DISABLED)
    self.progress["maximum"] = len(self.recipients)
    self.progress["value"] = 0

    thread = threading.Thread(
        target=self._worker,
        args=(
            endpoint,
            channel,
            token,
            from_email,
            reply_to,
            subject,
            body,
            interval,
        ),
        daemon=True,
    )
    thread.start()

  def stop_sending(self):
    if self.is_sending:
      self.stop_requested = True
      self.log("收到终止指令，将在当前邮件发送结束后停止...")

  def _worker(
      self,
      endpoint,
      channel_name,
      token,
      from_email,
      reply_to,
      subject,
      body,
      interval,
  ):
    total = len(self.recipients)
    success_count = 0
    fail_count = 0

    self.log(
        f"开始批量发送任务 [{channel_name}]，共 {total} 封..."
        + (f" [Reply-To: {reply_to}]" if reply_to else "")
    )

    for idx, to_email in enumerate(self.recipients, start=1):
      if self.stop_requested:
        self.log(f"发送任务已手动中断。完成: {idx-1}/{total}")
        break

      payload = {
          "token": token,
          "from": from_email,
          "to": to_email,
          "subject": subject,
          "body": body,
      }
      if reply_to:
        payload["reply_to"] = reply_to

      headers = {"Content-Type": "application/x-www-form-urlencoded"}

      try:
        resp = requests.post(endpoint, data=payload, headers=headers, timeout=15)
        res_json = resp.json()

        if resp.status_code == 200 and res_json.get("ret") == 1:
          success_count += 1
          self.log(f"[{idx}/{total}] 成功 -> {to_email}")
        else:
          fail_count += 1
          error_msg = res_json.get("msg", "未知错误")
          self.log(
              f"[{idx}/{total}] 失败 -> {to_email} | HTTP {resp.status_code}:"
              f" {error_msg}"
          )

          # 关键错误自动停机保护
          if resp.status_code in [400, 401, 403]:
            self.log(f"严重错误: 触发保护停机 (HTTP {resp.status_code})")
            break

      except requests.RequestException as e:
        fail_count += 1
        self.log(f"[{idx}/{total}] 请求网络异常 -> {to_email} | {str(e)}")

      self.progress["value"] = idx
      if idx < total:
        time.sleep(interval)

    self.log(f"任务结束！成功: {success_count} 封，失败: {fail_count} 封。")
    self.is_sending = False
    self.start_btn.config(state=tk.NORMAL)
    self.stop_btn.config(state=tk.DISABLED)
    self.channel_combo.config(state="readonly")


if __name__ == "__main__":
  root = tk.Tk()
  app = MailSenderApp(root)
  root.mainloop()