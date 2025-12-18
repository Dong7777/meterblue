# 在原有 blueexe.py 基础上增加：
# 1️⃣ 自动扫描蓝牙设备
# 2️⃣ 设备信息显示框（名称 / MAC / RSSI）

import asyncio
import serial
import tkinter as tk
from tkinter import messagebox, simpledialog
from bleak import BleakClient, BleakError, BleakScanner
import threading
import warnings

# 关闭 Bleak RSSI 废弃警告（不影响功能）
warnings.filterwarnings(
    "ignore",
    message=".*BLEDevice.rssi is deprecated.*",
    category=FutureWarning,
)

# =============================
# 🔧 默认配置
# =============================
SERIAL_PORT = "COM9"
SERIAL_BAUDRATE = 9600
SERIAL_TIMEOUT = 1
TARGET_MAC = ""
BLE_PIN = "111111"
BLE_NOTIFY_UUID = "0000fff1-0000-1000-8000-00805f9b34fb"
BLE_WRITE_UUID  = "0000fff2-0000-1000-8000-00805f9b34fb"

# =============================
# 📡 蓝牙扫描
# =============================
async def scan_ble_devices():
    devices = await BleakScanner.discover(timeout=5.0)
    return devices

def scan_button_click():
    device_listbox.delete(0, tk.END)

    def _scan():
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        devices = loop.run_until_complete(scan_ble_devices())
        for d in devices:
            name = d.name or "Unknown"
            # Bleak 新版已废弃 d.rssi，这里仅用于显示（已屏蔽警告）
            rssi = getattr(d, 'rssi', None)
            line = f"{name} | {d.address} | RSSI={rssi if rssi is not None else 'N/A'}"
            device_listbox.insert(tk.END, line)

    threading.Thread(target=_scan, daemon=True).start()

# =============================
# 🔗 桥接逻辑（已修复语法错误）
# =============================
ble_client = None
serial_handle = None
stop_event = threading.Event()
bridge_loop = None  # asyncio 事件循环引用

async def ble_notify_loop(client, ser):
    def notification_handler(sender, data):
        if stop_event.is_set():
            return
        log_message(f"[BLE→Serial] {data.hex()}")
        try:
            ser.write(data)
        except Exception as e:
            log_message(f"⚠️ 串口写入失败: {e}")

    await client.start_notify(BLE_NOTIFY_UUID, notification_handler)

    while not stop_event.is_set() and client.is_connected:
        await asyncio.sleep(0.1)

    # ⚠️ Windows + Bleak 下 stop_notify 可能抛 KeyError，需保护
    try:
        if client.is_connected:
            await client.stop_notify(BLE_NOTIFY_UUID)
    except Exception:
        pass


async def serial_to_ble(client, ser):
    while not stop_event.is_set():
        await asyncio.sleep(0.01)
        if ser.in_waiting:
            data = ser.read(ser.in_waiting)
            log_message(f"[Serial→BLE] {data.hex()}")
            try:
                await client.write_gatt_char(BLE_WRITE_UUID, data, response=False)
            except Exception as e:
                log_message(f"⚠️ 蓝牙写入失败: {e}")


async def start_bridge_async():
    global ble_client, serial_handle
    stop_event.clear()

    serial_handle = serial.Serial(SERIAL_PORT, SERIAL_BAUDRATE, timeout=SERIAL_TIMEOUT)
    ble_client = BleakClient(TARGET_MAC)

    await ble_client.connect()
    log_message("🔗 已连接蓝牙，尝试配对...")

    try:
        paired = await ble_client.pair()
    except Exception:
        paired = False

    if not paired:
        pin = simpledialog.askstring("蓝牙配对", "设备需要配对密码，请输入 PIN：", show='*')
        if not pin:
            log_message("❌ 用户取消配对")
            return
        try:
            await ble_client.pair(pin=pin)
            log_message("✅ 蓝牙配对成功")
        except Exception as e:
            log_message(f"❌ 蓝牙配对失败: {e}")
            return

    log_message("✅ 蓝牙已连接并完成配对")

    try:
        await asyncio.gather(
            ble_notify_loop(ble_client, serial_handle),
            serial_to_ble(ble_client, serial_handle),
        )
    finally:
        # 确保退出时真正断开 BLE，避免 pending task
        try:
            if ble_client and ble_client.is_connected:
                await ble_client.disconnect()
                log_message("🔌 蓝牙已断开")
        except Exception as e:
            log_message(f"⚠️ 蓝牙断开异常: {e}")

        # 关闭串口并记录日志
        try:
            if serial_handle and serial_handle.is_open:
                serial_handle.close()
                log_message("🔌 串口已关闭")
        except Exception as e:
            log_message(f"⚠️ 串口关闭异常: {e}")

# =============================
# 🧵 线程封装

# =============================
def start_bridge():
    global bridge_loop
    # 以 bridge_loop 作为唯一“是否在连接中”的判断
    if bridge_loop is not None:
        log_message("⚠️ 已经在连接中")
        return

    def _run():
        global bridge_loop, ble_client, serial_handle
        loop = asyncio.new_event_loop()
        bridge_loop = loop
        asyncio.set_event_loop(loop)
        try:
            loop.run_until_complete(start_bridge_async())
        finally:
            # 在 loop 关闭前，取消所有未完成任务
            pending = asyncio.all_tasks(loop)
            for task in pending:
                task.cancel()
            try:
                loop.run_until_complete(asyncio.gather(*pending, return_exceptions=True))
            except Exception:
                pass
            loop.close()
            # ⭐ 关键：线程结束时统一清状态，允许再次连接
            bridge_loop = None
            ble_client = None
            serial_handle = None
            stop_event.clear()
            log_message("✅ 断开连接完成，可重新连接")
    threading.Thread(target=_run, daemon=True).start()


def disconnect_bridge():
    global bridge_loop

    log_message("⏹ 正在断开连接...")
    # 只发停止信号，真正释放由后台线程统一完成
    stop_event.set()

    # 轻触后台 loop，让其尽快从 await 中返回
    if bridge_loop:
        try:
            bridge_loop.call_soon_threadsafe(lambda: None)
        except Exception:
            pass

# =============================
# 🧰 GUI 工具函数
# =============================
def log_message(msg):
    log_widget.insert(tk.END, msg + "\n")
    log_widget.yview(tk.END)


def apply_config():
    global SERIAL_PORT, SERIAL_BAUDRATE, TARGET_MAC, BLE_PIN

    SERIAL_PORT = serial_entry.get()
    SERIAL_BAUDRATE = int(baud_entry.get())
    TARGET_MAC = mac_entry.get()
    BLE_PIN = pin_entry.get()

    log_message("✅ 配置已应用")


def on_device_select(event):
    # ⚠️ 可能在列表刷新 / 断开过程中触发空选择，需保护
    sel = device_listbox.curselection()
    if not sel:
        return

    selection = device_listbox.get(sel[0])
    parts = selection.split("|")
    if len(parts) < 2:
        return

    mac = parts[1].strip()
    mac_entry.delete(0, tk.END)
    mac_entry.insert(0, mac)

# =============================
# 🖥 GUI
# =============================
window = tk.Tk()
window.title("蓝牙串口桥接（带自动扫描）")
window.geometry("750x650")

cfg = tk.LabelFrame(window, text="配置")
cfg.pack(fill="x", padx=10, pady=5)

tk.Label(cfg, text="串口").grid(row=0, column=0)
serial_entry = tk.Entry(cfg)
serial_entry.insert(0, SERIAL_PORT)
serial_entry.grid(row=0, column=1)

tk.Label(cfg, text="波特率").grid(row=0, column=2)
baud_entry = tk.Entry(cfg, width=8)
baud_entry.insert(0, SERIAL_BAUDRATE)
baud_entry.grid(row=0, column=3)

tk.Label(cfg, text="蓝牙 MAC").grid(row=1, column=0)
mac_entry = tk.Entry(cfg, width=25)
mac_entry.grid(row=1, column=1, columnspan=2)

tk.Label(cfg, text="PIN").grid(row=1, column=3)
pin_entry = tk.Entry(cfg, width=8)
pin_entry.insert(0, BLE_PIN)
pin_entry.grid(row=1, column=4)

# 扫描区
scan_frame = tk.LabelFrame(window, text="扫描到的蓝牙设备")
scan_frame.pack(fill="both", expand=True, padx=10, pady=5)

device_listbox = tk.Listbox(scan_frame, height=8)
device_listbox.pack(fill="both", expand=True)
device_listbox.bind("<<ListboxSelect>>", on_device_select)

scan_btn = tk.Button(scan_frame, text="🔍 扫描蓝牙设备", command=scan_button_click)
scan_btn.pack(pady=5)

# 操作按钮
actions = tk.Frame(window)
actions.pack(pady=5)

tk.Button(actions, text="应用配置", command=apply_config).pack(side=tk.LEFT, padx=5)
tk.Button(actions, text="启动桥接", command=start_bridge).pack(side=tk.LEFT, padx=5)
tk.Button(actions, text="断开连接", command=disconnect_bridge).pack(side=tk.LEFT, padx=5)

# 日志
log_widget = tk.Text(window, height=10)
log_widget.pack(fill="both", padx=10, pady=5)

window.mainloop()
