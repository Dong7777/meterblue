import asyncio
import serial
import serial.tools.list_ports
import tkinter as tk
from tkinter import messagebox
from bleak import BleakClient, BleakScanner
import threading
import warnings
import traceback
import json
import os
import sys
import datetime

# =============================
# 🔧 基础配置与全局变量
# =============================
warnings.filterwarnings(
    "ignore",
    message=".*BLEDevice.rssi is deprecated.*",
    category=FutureWarning,
)

DEFAULT_CONFIG = {
    "serial_port": "COM9",
    "baudrate": 9600,
    "mac": "",
}

CONFIG_FILE = os.path.join(os.path.expanduser("~"), "ble_serial_bridge_config.json")
LOG_FILE = os.path.join(os.path.expanduser("~"), "ble_serial_bridge.log")

BLE_NOTIFY_UUID = "0000fff1-0000-1000-8000-00805f9b34fb"
BLE_WRITE_UUID = "0000fff2-0000-1000-8000-00805f9b34fb"

# 全局状态
ble_client = None
serial_handle = None
bridge_loop = None
bridge_tasks = []
stop_event = threading.Event()
ble_fail_count = 0
BLE_FAIL_THRESHOLD = 3
is_connecting = False
window = None
status_var = None
status_colors = {
    "未连接": "gray",
    "连接中": "orange",
    "已连接": "green",
    "异常": "red"
}

# =============================
# 📝 工具函数
# =============================
def load_config():
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            config = json.load(f)
        for key in DEFAULT_CONFIG:
            if key not in config:
                config[key] = DEFAULT_CONFIG[key]
        return config
    except:
        return DEFAULT_CONFIG.copy()

def save_config(config):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        log_message(f"⚠️ 配置保存失败: {e}")
        return False

def log_message(msg):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_line = f"[{timestamp}] {msg}"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(log_line + "\n")
    except:
        print(f"日志写入失败: {msg}")

    def _update_log():
        if 'log_widget' in globals() and log_widget:
            log_widget.insert(tk.END, log_line + "\n")
            log_widget.yview(tk.END)
            # 限制日志行数
            line_count = int(log_widget.index('end-1c').split('.')[0])
            if line_count > 1000:
                log_widget.delete(1.0, 2.0)
    if window:
        window.after(0, _update_log)

def clear_log():
    """清空日志（GUI + 文件）"""
    global log_widget
    try:
        log_widget.delete(1.0, tk.END)
    except:
        pass
    try:
        if os.path.exists(LOG_FILE):
            os.remove(LOG_FILE)
    except Exception as e:
        print(f"⚠️ 日志文件清理失败: {e}")

def get_available_ports():
    try:
        ports = [port.device for port in serial.tools.list_ports.comports()]
        return ports if ports else ["COM1", "COM2", "COM3"]
    except:
        return ["COM1", "COM2", "COM3"]

# =============================
# 🔁 BLE ↔ 串口桥接逻辑
# =============================
async def ble_notify_loop(client, ser):
    global ble_fail_count
    def handler(sender, data):
        global ble_fail_count
        if stop_event.is_set():
            return
        try:
            ser.write(data)
            ble_fail_count = 0
            log_message(f"[BLE→SER] {data.hex()}")
        except Exception as e:
            ble_fail_count += 1
            log_message(f"❌ BLE→SER 写失败 ({ble_fail_count}/{BLE_FAIL_THRESHOLD}): {e}")
            if ble_fail_count >= BLE_FAIL_THRESHOLD:
                log_message("❌ 连续通信失败，触发断开")
                stop_event.set()

    try:
        await client.start_notify(BLE_NOTIFY_UUID, handler)
        while not stop_event.is_set():
            await asyncio.sleep(0.1)
    except Exception as e:
        ble_fail_count += 1
        log_message(f"❌ Notify 异常 ({ble_fail_count}/{BLE_FAIL_THRESHOLD}): {e}")
        if ble_fail_count >= BLE_FAIL_THRESHOLD:
            stop_event.set()
    finally:
        try: await client.stop_notify(BLE_NOTIFY_UUID)
        except: pass

async def serial_to_ble(client, ser):
    global ble_fail_count
    try:
        while not stop_event.is_set():
            await asyncio.sleep(0.01)
            if ser.in_waiting:
                data = ser.read(ser.in_waiting)
                try:
                    await client.write_gatt_char(BLE_WRITE_UUID, data, response=False)
                    ble_fail_count = 0
                    log_message(f"[SER→BLE] {data.hex()}")
                except Exception as e:
                    ble_fail_count += 1
                    log_message(f"❌ SER→BLE 写失败 ({ble_fail_count}/{BLE_FAIL_THRESHOLD}): {e}")
                    if ble_fail_count >= BLE_FAIL_THRESHOLD:
                        log_message("❌ 连续通信失败，触发断开")
                        stop_event.set()
    except Exception as e:
        log_message(f"❌ 串口读取异常: {e}")
        stop_event.set()

async def start_bridge_async(config):
    global ble_client, serial_handle, bridge_tasks
    stop_event.clear()
    bridge_tasks = []

    # 打开串口
    try:
        serial_handle = serial.Serial(config["serial_port"], config["baudrate"], timeout=1)
        log_message(f"✅ 串口 {config['serial_port']} 已打开")
    except Exception as e:
        log_message(f"❌ 串口初始化失败: {e}")
        window.after(0, lambda: [
            messagebox.showerror("串口错误", f"无法打开串口：{e}"),
            status_var.set("异常"),
            status_label.config(fg=status_colors["异常"])
        ])
        stop_event.set()
        return

    # 连接 BLE
    try:
        ble_client = BleakClient(config["mac"])
        await ble_client.connect()
        log_message("🔗 蓝牙已连接")
    except Exception as e:
        log_message(f"❌ 蓝牙连接失败: {e}")
        window.after(0, lambda: [
            messagebox.showerror("蓝牙错误", f"无法连接蓝牙：{e}"),
            status_var.set("异常"),
            status_label.config(fg=status_colors["异常"])
        ])
        if serial_handle and serial_handle.is_open:
            serial_handle.close()
        stop_event.set()
        return

    # 更新状态
    window.after(0, lambda: [
        status_var.set("已连接"),
        status_label.config(fg=status_colors["已连接"])
    ])
    log_message("✅ 桥接已启动")

    # 启动任务
    tasks = [
        asyncio.create_task(ble_notify_loop(ble_client, serial_handle)),
        asyncio.create_task(serial_to_ble(ble_client, serial_handle)),
    ]
    bridge_tasks = tasks

    try:
        while not stop_event.is_set():
            await asyncio.sleep(0.2)
    finally:
        log_message("🧹 正在清理资源...")
        for t in tasks:
            t.cancel()
        await asyncio.gather(*tasks, return_exceptions=True)
        # 关闭 BLE
        try:
            if ble_client and ble_client.is_connected:
                await ble_client.disconnect()
                log_message("🔌 蓝牙已断开")
        except:
            pass
        ble_client = None
        # 关闭串口
        try:
            if serial_handle and serial_handle.is_open:
                serial_handle.close()
                log_message("🔌 串口已关闭")
        except:
            pass
        serial_handle = None

        window.after(0, lambda: [
            status_var.set("未连接"),
            status_label.config(fg=status_colors["未连接"])
        ])
        stop_event.clear()
        clear_log()
        log_message("✅ 桥接任务已结束，日志已清空")

# =============================
# 🧵 线程封装与操作函数
# =============================
def hard_reset_bridge():
    global ble_client, serial_handle, bridge_loop, bridge_tasks, ble_fail_count, stop_event
    log_message("♻ 执行硬复位...")
    stop_event.set()

    # 取消任务
    if bridge_tasks:
        for t in bridge_tasks:
            t.cancel()
        bridge_tasks.clear()

    # BLE 断开
    if ble_client:
        try:
            if ble_client.is_connected:
                asyncio.run(ble_client.disconnect())
                log_message("🔌 BLE 已断开")
        except: pass
    ble_client = None

    # 串口关闭
    if serial_handle:
        try:
            if serial_handle.is_open:
                serial_handle.close()
                log_message("🔌 串口已关闭")
        except: pass
    serial_handle = None

    # 事件循环关闭
    if bridge_loop:
        try:
            pending = asyncio.all_tasks(bridge_loop)
            for t in pending:
                t.cancel()
            bridge_loop.run_until_complete(asyncio.gather(*pending, return_exceptions=True))
        except: pass
        try:
            bridge_loop.close()
        except: pass
    bridge_loop = None

    ble_fail_count = 0
    stop_event.clear()
    status_var.set("未连接")
    status_label.config(fg=status_colors["未连接"])
    clear_log()
    log_message("✅ 硬复位完成，日志已清空")

def disconnect_bridge():
    global is_connecting
    if not is_connecting:
        log_message("⚠️ 当前未连接")
        return
    log_message("⏹ 请求断开连接")
    stop_event.set()
    status_var.set("未连接")
    status_label.config(fg=status_colors["未连接"])
    clear_log()
    log_message("✅ 已断开连接，日志已清空")

def start_bridge():
    global bridge_loop, bridge_tasks, is_connecting, ble_fail_count
    if is_connecting:
        log_message("⚠️ 已在连接中")
        return

    hard_reset_bridge()

    config = {
        "serial_port": serial_var.get(),
        "baudrate": int(baud_entry.get()),
        "mac": mac_entry.get().strip(),
    }
    if not config["mac"]:
        messagebox.showwarning("配置错误", "请先选择或输入蓝牙MAC地址")
        return

    save_config(config)
    status_var.set("连接中")
    status_label.config(fg=status_colors["连接中"])
    is_connecting = True
    ble_fail_count = 0

    def _run():
        global bridge_loop, bridge_tasks, is_connecting
        try:
            bridge_loop = asyncio.new_event_loop()
            asyncio.set_event_loop(bridge_loop)
            bridge_loop.run_until_complete(start_bridge_async(config))
        except Exception as e:
            log_message(f"❌ 桥接线程异常: {e}")
            log_message(traceback.format_exc())
        finally:
            hard_reset_bridge()
            is_connecting = False
            log_message("✅ 桥接线程已完全退出，可重新启动")

    threading.Thread(target=_run, daemon=True).start()

# =============================
# 🔍 BLE 扫描功能
# =============================
async def scan_ble_devices():
    return await BleakScanner.discover(timeout=5)

def scan_button_click():
    device_listbox.delete(0, tk.END)
    log_message("🔍 开始扫描蓝牙设备...")

    def _scan():
        try:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            devices = loop.run_until_complete(scan_ble_devices())
            loop.close()

            def _update_list():
                for d in devices:
                    name = d.name or "Unknown"
                    rssi = getattr(d, "rssi", "N/A")
                    device_listbox.insert(tk.END, f"{name} | {d.address} | RSSI={rssi}")
                log_message(f"✅ 扫描完成，共发现 {len(devices)} 个设备")
            window.after(0, _update_list)
        except Exception as e:
            log_message(f"❌ 扫描失败: {str(e)}")

    threading.Thread(target=_scan, daemon=True).start()

def on_device_select(event):
    sel = device_listbox.curselection()
    if not sel:
        return
    try:
        mac = device_listbox.get(sel[0]).split("|")[1].strip()
        mac_entry.delete(0, tk.END)
        mac_entry.insert(0, mac)
    except IndexError:
        log_message("⚠️ 设备信息解析失败")

# =============================
# 🖥 GUI初始化
# =============================
def init_gui():
    global window, log_widget, device_listbox, serial_var, serial_entry
    global baud_entry, mac_entry, status_label, status_var

    config = load_config()
    window = tk.Tk()
    window.title("蓝牙串口桥接工具 v1.0")
    window.geometry("780x720")
    window.resizable(True, True)

    status_var = tk.StringVar(value="未连接")
    status_label = tk.Label(window, textvariable=status_var, fg=status_colors["未连接"], font=("Arial", 12, "bold"))
    status_label.pack(pady=2)

    cfg = tk.LabelFrame(window, text="基础配置")
    cfg.pack(fill="x", padx=10, pady=5)

    tk.Label(cfg, text="串口").grid(row=0, column=0, padx=5, pady=3)
    serial_var = tk.StringVar(value=config["serial_port"])
    serial_entry = tk.OptionMenu(cfg, serial_var, *get_available_ports())
    serial_entry.grid(row=0, column=1, padx=5, pady=3)

    tk.Label(cfg, text="波特率").grid(row=0, column=3, padx=5, pady=3)
    baud_entry = tk.Entry(cfg, width=8)
    baud_entry.insert(0, config["baudrate"])
    baud_entry.grid(row=0, column=4, padx=5, pady=3)

    tk.Label(cfg, text="蓝牙MAC").grid(row=1, column=0, padx=5, pady=3)
    mac_entry = tk.Entry(cfg, width=25)
    mac_entry.insert(0, config["mac"])
    mac_entry.grid(row=1, column=1, columnspan=2, padx=5, pady=3)

    # 扫描区域
    scan_frame = tk.LabelFrame(window, text="蓝牙设备扫描")
    scan_frame.pack(fill="both", expand=True, padx=10, pady=5)
    device_listbox = tk.Listbox(scan_frame, height=8)
    device_listbox.pack(fill="both", expand=True, padx=5, pady=5)
    device_listbox.bind("<<ListboxSelect>>", on_device_select)
    tk.Button(scan_frame, text="🔍 扫描蓝牙设备", command=scan_button_click).pack(pady=5)

    actions = tk.Frame(window)
    actions.pack(pady=8)
    tk.Button(actions, text="启动桥接", command=start_bridge, bg="#4CAF50", fg="white").pack(side=tk.LEFT, padx=4)
    tk.Button(actions, text="断开连接", command=disconnect_bridge, bg="#FF9800", fg="white").pack(side=tk.LEFT, padx=4)
    tk.Button(actions, text="♻ 初始化", command=hard_reset_bridge, bg="#F44336", fg="white").pack(side=tk.LEFT, padx=4)

    log_frame = tk.LabelFrame(window, text="运行日志")
    log_frame.pack(fill="both", expand=True, padx=10, pady=5)
    log_widget = tk.Text(log_frame, height=12, font=("Consolas", 9))
    log_widget.pack(fill="both", expand=True, padx=5, pady=5, side=tk.LEFT)
    log_scroll = tk.Scrollbar(log_frame, command=log_widget.yview)
    log_scroll.pack(fill="y", side=tk.RIGHT)
    log_widget.config(yscrollcommand=log_scroll.set)

    def on_closing():
        hard_reset_bridge()
        window.destroy()

    window.protocol("WM_DELETE_WINDOW", on_closing)

    log_message("✅ 程序启动完成")
    log_message(f"ℹ️ 配置文件路径: {CONFIG_FILE}")
    log_message(f"ℹ️ 日志文件路径: {LOG_FILE}")

    window.mainloop()

# =============================
# 🚀 程序入口
# =============================
if __name__ == "__main__":
    if sys.platform == 'win32':
        import ctypes
        def is_admin():
            try:
                return ctypes.windll.shell32.IsUserAnAdmin()
            except: return False
        if not is_admin():
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
            sys.exit(0)
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

    init_gui()
