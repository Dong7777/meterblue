import asyncio
import serial
import serial.tools.list_ports
import tkinter as tk
from tkinter import simpledialog, messagebox
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

# 默认配置
DEFAULT_CONFIG = {
    "serial_port": "COM9",
    "baudrate": 9600,
    "mac": "",
    "pin": "111111"
}

# 文件路径配置
CONFIG_FILE = os.path.join(os.path.expanduser("~"), "ble_serial_bridge_config.json")
LOG_FILE = os.path.join(os.path.expanduser("~"), "ble_serial_bridge.log")

# BLE配置
BLE_NOTIFY_UUID = "0000fff1-0000-1000-8000-00805f9b34fb"
BLE_WRITE_UUID = "0000fff2-0000-1000-8000-00805f9b34fb"

# 全局状态（仅定义变量名，不初始化tkinter对象）
ble_client = None
serial_handle = None
bridge_loop = None
stop_event = threading.Event()
is_connecting = False
window = None
status_var = None  # 延迟初始化
status_colors = {
    "未连接": "gray",
    "连接中": "orange",
    "已连接": "green",
    "异常": "red"
}


# =============================
# 📝 工具函数（配置/日志/权限）
# =============================
def load_config():
    """加载持久化配置"""
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            config = json.load(f)
        # 兼容旧配置，补充缺失字段
        for key in DEFAULT_CONFIG:
            if key not in config:
                config[key] = DEFAULT_CONFIG[key]
        return config
    except (FileNotFoundError, json.JSONDecodeError):
        return DEFAULT_CONFIG.copy()


def save_config(config):
    """保存配置到文件"""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        log_message(f"⚠️ 配置保存失败: {e}")
        return False


def log_message(msg):
    """线程安全的日志输出（GUI+文件）"""
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_line = f"[{timestamp}] {msg}"

    # 写入日志文件
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(log_line + "\n")
    except Exception as e:
        print(f"日志文件写入失败: {e}")

    # 线程安全更新GUI日志
    def _update_log():
        if 'log_widget' in globals() and log_widget:
            log_widget.insert(tk.END, log_line + "\n")
            log_widget.yview(tk.END)
            # 限制日志行数，避免卡顿
            line_count = int(log_widget.index('end-1c').split('.')[0])
            if line_count > 1000:
                log_widget.delete(1.0, 2.0)

    if window:
        window.after(0, _update_log)


def get_available_ports():
    """获取可用串口列表"""
    try:
        ports = [port.device for port in serial.tools.list_ports.comports()]
        return ports if ports else ["COM1", "COM2", "COM3"]
    except Exception:
        return ["COM1", "COM2", "COM3"]


def is_admin():
    """检查是否管理员权限（Windows）"""
    if sys.platform != 'win32':
        return True
    try:
        import ctypes
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def elevate_admin():
    """提权到管理员（Windows）"""
    if sys.platform != 'win32' or is_admin():
        return True
    try:
        import ctypes
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, " ".join(sys.argv), None, 1
        )
        sys.exit(0)
    except Exception as e:
        messagebox.showerror("权限不足", f"需要管理员权限运行：{e}")
        return False


# =============================
# 📡 BLE 相关功能
# =============================
async def scan_ble_devices():
    """扫描BLE设备"""
    return await BleakScanner.discover(timeout=5)


def scan_button_click():
    """扫描蓝牙设备按钮点击事件"""
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
                    device_listbox.insert(
                        tk.END, f"{name} | {d.address} | RSSI={rssi}"
                    )
                log_message(f"✅ 扫描完成，共发现 {len(devices)} 个设备")

            window.after(0, _update_list)
        except Exception as e:
            log_message(f"❌ 扫描失败: {str(e)}")

    threading.Thread(target=_scan, daemon=True).start()


def on_device_select(event):
    """选择蓝牙设备后自动填充MAC"""
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
# 🔁 桥接核心逻辑
# =============================
async def ble_notify_loop(client, ser):
    """BLE → 串口数据转发"""

    def handler(sender, data):
        if stop_event.is_set():
            return
        try:
            ser.write(data)
            log_message(f"[BLE→SER] {data.hex()}")
        except Exception as e:
            log_message(f"❌ 串口写失败: {e}")
            stop_event.set()

    try:
        await client.start_notify(BLE_NOTIFY_UUID, handler)
        while not stop_event.is_set():
            await asyncio.sleep(0.1)
    except Exception as e:
        log_message(f"❌ Notify 异常: {e}")
        stop_event.set()
    finally:
        try:
            await client.stop_notify(BLE_NOTIFY_UUID)
        except Exception:
            pass


async def serial_to_ble(client, ser):
    """串口 → BLE数据转发"""
    try:
        while not stop_event.is_set():
            await asyncio.sleep(0.01)
            if ser.in_waiting:
                data = ser.read(ser.in_waiting)
                log_message(f"[SER→BLE] {data.hex()}")
                await client.write_gatt_char(
                    BLE_WRITE_UUID, data, response=False
                )
    except Exception as e:
        log_message(f"❌ BLE 写失败: {e}")
        stop_event.set()


async def ble_watchdog(client):
    """BLE连接看门狗"""
    while not stop_event.is_set():
        await asyncio.sleep(1)
        if not client.is_connected:
            log_message("❌ BLE 异常断开")
            stop_event.set()
            window.after(0, lambda: [
                status_var.set("异常"),
                status_label.config(fg=status_colors["异常"])
            ])
            break


async def start_bridge_async(config):
    """桥接主逻辑"""
    global ble_client, serial_handle
    stop_event.clear()
    pairing_success = False

    # 1. 初始化串口
    try:
        serial_handle = serial.Serial(
            config["serial_port"], config["baudrate"], timeout=1
        )
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

    # 2. 连接蓝牙
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

    # 3. 蓝牙配对
    try:
        await ble_client.pair(protection_level=2)
        pairing_success = True
        log_message("✅ 蓝牙配对完成（无需PIN）")
    except Exception as e:
        log_message(f"ℹ️ 需要PIN码配对: {e}")
        pin_result = None

        def _get_pin():
            nonlocal pin_result
            pin_result = simpledialog.askstring("蓝牙配对", "请输入 PIN 码：", show="*")

        window.after(0, _get_pin)

        # 等待PIN输入（最多60秒）
        wait_time = 0
        while pin_result is None and wait_time < 600:
            await asyncio.sleep(0.1)
            wait_time += 1

        if pin_result is None:
            log_message("❌ 用户取消PIN输入")
            stop_event.set()
            return

        try:
            await ble_client.pair(pin=pin_result, protection_level=2)
            pairing_success = True
            log_message("✅ 蓝牙配对完成（使用PIN）")
        except Exception as e:
            log_message(f"❌ PIN配对失败: {e}")
            window.after(0, lambda: [
                messagebox.showerror("配对失败", f"PIN码错误或配对失败：{e}"),
                status_var.set("异常"),
                status_label.config(fg=status_colors["异常"])
            ])
            stop_event.set()
            return

    if not pairing_success:
        log_message("❌ 蓝牙配对未完成")
        stop_event.set()
        return

    # 4. 更新状态并启动桥接任务
    window.after(0, lambda: [
        status_var.set("已连接"),
        status_label.config(fg=status_colors["已连接"])
    ])
    log_message("✅ 桥接已启动")

    tasks = [
        asyncio.create_task(ble_notify_loop(ble_client, serial_handle)),
        asyncio.create_task(serial_to_ble(ble_client, serial_handle)),
        asyncio.create_task(ble_watchdog(ble_client)),
    ]

    try:
        while not stop_event.is_set():
            await asyncio.sleep(0.2)
    finally:
        log_message("🧹 正在清理资源...")

        # 取消所有异步任务
        for t in tasks:
            t.cancel()
        await asyncio.gather(*tasks, return_exceptions=True)

        # 关闭蓝牙连接
        try:
            if ble_client and ble_client.is_connected:
                await ble_client.disconnect()
                log_message("🔌 蓝牙已断开")
        except Exception as e:
            log_message(f"⚠️ 蓝牙断开失败: {e}")

        # 关闭串口
        try:
            if serial_handle and serial_handle.is_open:
                serial_handle.close()
                log_message("🔌 串口已关闭")
        except Exception as e:
            log_message(f"⚠️ 串口关闭失败: {e}")

        # 重置状态
        window.after(0, lambda: [
            status_var.set("未连接"),
            status_label.config(fg=status_colors["未连接"])
        ])


# =============================
# 🧵 线程封装与操作函数
# =============================
def start_bridge():
    """启动桥接（按钮点击）"""
    global is_connecting
    if is_connecting:
        log_message("⚠️ 已在连接中")
        return

    # 获取当前配置
    config = {
        "serial_port": serial_var.get(),
        "baudrate": int(baud_entry.get().strip()),
        "mac": mac_entry.get().strip(),
        "pin": pin_entry.get().strip()
    }

    # 校验配置
    if not config["mac"].strip():
        messagebox.showwarning("配置错误", "请先选择或输入蓝牙MAC地址")
        return

    # 保存配置并更新状态
    save_config(config)
    status_var.set("连接中")
    status_label.config(fg=status_colors["连接中"])
    is_connecting = True

    def _run():
        global bridge_loop, is_connecting
        try:
            # 适配Windows异步循环
            if sys.platform == 'win32':
                asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
            bridge_loop = asyncio.new_event_loop()
            asyncio.set_event_loop(bridge_loop)
            bridge_loop.run_until_complete(start_bridge_async(config))
        except Exception as e:
            log_message(f"❌ 桥接线程异常: {str(e)}")
            log_message(f"📝 详细错误: {traceback.format_exc()}")
            window.after(0, lambda: [
                status_var.set("异常"),
                status_label.config(fg=status_colors["异常"])
            ])
        finally:
            # 清理资源
            pending = asyncio.all_tasks(bridge_loop) if bridge_loop else []
            for t in pending:
                t.cancel()
            if bridge_loop:
                try:
                    bridge_loop.run_until_complete(
                        asyncio.gather(*pending, return_exceptions=True)
                    )
                except Exception:
                    pass
                bridge_loop.close()

            # 重置全局状态
            global ble_client, serial_handle
            bridge_loop = None
            ble_client = None
            serial_handle = None
            is_connecting = False
            stop_event.clear()

            log_message("✅ 连接线程已完全退出，可重新启动")

    threading.Thread(target=_run, daemon=True).start()


def disconnect_bridge():
    """断开桥接"""
    if not is_connecting:
        log_message("⚠️ 当前未连接")
        return

    log_message("⏹ 请求断开连接")
    stop_event.set()
    status_var.set("未连接")
    status_label.config(fg=status_colors["未连接"])

    if bridge_loop:
        try:
            bridge_loop.call_soon_threadsafe(lambda: None)
        except Exception as e:
            log_message(f"⚠️ 停止信号发送失败: {e}")


def hard_reset_bridge():
    """硬复位"""
    log_message("♻ 初始化（系统 RESET）")
    stop_event.set()

    # 强制清理资源
    global ble_client, serial_handle, bridge_loop
    try:
        if serial_handle and serial_handle.is_open:
            serial_handle.close()
    except Exception:
        pass
    serial_handle = None

    if bridge_loop:
        try:
            bridge_loop.call_soon_threadsafe(lambda: None)
        except Exception:
            pass
    bridge_loop = None
    ble_client = None
    is_connecting = False
    status_var.set("未连接")
    status_label.config(fg=status_colors["未连接"])

    log_message("✅ 硬复位完成")


def clear_log():
    """清空日志"""
    log_widget.delete(1.0, tk.END)
    try:
        os.remove(LOG_FILE)
        log_message("✅ 日志已清空")
    except Exception as e:
        log_message(f"⚠️ 日志文件清理失败: {e}")


def refresh_ports():
    """刷新串口列表"""
    menu = serial_entry["menu"]
    menu.delete(0, "end")
    ports = get_available_ports()
    for port in ports:
        menu.add_command(label=port, command=tk._setit(serial_var, port))
    log_message(f"✅ 串口列表已刷新，当前可用：{ports}")


# =============================
# 🖥 GUI初始化
# =============================
def init_gui():
    global window, log_widget, device_listbox, serial_var, serial_entry
    global baud_entry, mac_entry, pin_entry, status_label, status_var

    # 加载配置
    config = load_config()

    # 第一步：创建主窗口（必须先创建窗口，再初始化tkinter变量）
    window = tk.Tk()
    window.title("蓝牙串口桥接工具 v1.0")
    window.geometry("780x720")
    window.resizable(True, True)

    # 第二步：初始化tkinter变量（此时已有根窗口）
    status_var = tk.StringVar(value="未连接")

    # 状态显示栏
    status_label = tk.Label(
        window, textvariable=status_var, fg=status_colors["未连接"],
        font=("Arial", 12, "bold")
    )
    status_label.pack(pady=2)

    # 配置区域
    cfg = tk.LabelFrame(window, text="基础配置")
    cfg.pack(fill="x", padx=10, pady=5)

    # 串口选择（下拉框+刷新）
    tk.Label(cfg, text="串口").grid(row=0, column=0, padx=5, pady=3)
    serial_var = tk.StringVar(value=config["serial_port"])
    serial_entry = tk.OptionMenu(cfg, serial_var, *get_available_ports())
    serial_entry.grid(row=0, column=1, padx=5, pady=3)
    tk.Button(cfg, text="刷新串口", command=refresh_ports).grid(row=0, column=2, padx=5, pady=3)

    # 波特率
    tk.Label(cfg, text="波特率").grid(row=0, column=3, padx=5, pady=3)
    baud_entry = tk.Entry(cfg, width=8)
    baud_entry.insert(0, config["baudrate"])
    baud_entry.grid(row=0, column=4, padx=5, pady=3)

    # 蓝牙MAC
    tk.Label(cfg, text="蓝牙MAC").grid(row=1, column=0, padx=5, pady=3)
    mac_entry = tk.Entry(cfg, width=25)
    mac_entry.insert(0, config["mac"])
    mac_entry.grid(row=1, column=1, columnspan=2, padx=5, pady=3)

    # PIN码
    tk.Label(cfg, text="PIN码").grid(row=1, column=3, padx=5, pady=3)
    pin_entry = tk.Entry(cfg, width=8)
    pin_entry.insert(0, config["pin"])
    pin_entry.grid(row=1, column=4, padx=5, pady=3)

    # 蓝牙设备扫描区域
    scan_frame = tk.LabelFrame(window, text="蓝牙设备扫描")
    scan_frame.pack(fill="both", expand=True, padx=10, pady=5)

    device_listbox = tk.Listbox(scan_frame, height=8)
    device_listbox.pack(fill="both", expand=True, padx=5, pady=5)
    device_listbox.bind("<<ListboxSelect>>", on_device_select)

    tk.Button(scan_frame, text="🔍 扫描蓝牙设备", command=scan_button_click).pack(pady=5)

    # 操作按钮区域
    actions = tk.Frame(window)
    actions.pack(pady=8)

    tk.Button(actions, text="启动桥接", command=start_bridge, bg="#4CAF50", fg="white").pack(side=tk.LEFT, padx=4)
    tk.Button(actions, text="断开连接", command=disconnect_bridge, bg="#FF9800", fg="white").pack(side=tk.LEFT, padx=4)
    tk.Button(actions, text="♻ 初始化", command=hard_reset_bridge, bg="#F44336", fg="white").pack(side=tk.LEFT, padx=4)
    tk.Button(actions, text="清空日志", command=clear_log).pack(side=tk.LEFT, padx=4)

    # 日志区域
    log_frame = tk.LabelFrame(window, text="运行日志")
    log_frame.pack(fill="both", expand=True, padx=10, pady=5)

    log_widget = tk.Text(log_frame, height=12, font=("Consolas", 9))
    log_widget.pack(fill="both", expand=True, padx=5, pady=5, side=tk.LEFT)

    # 日志滚动条
    log_scroll = tk.Scrollbar(log_frame, command=log_widget.yview)
    log_scroll.pack(fill="y", side=tk.RIGHT)
    log_widget.config(yscrollcommand=log_scroll.set)

    # 窗口关闭处理
    def on_closing():
        hard_reset_bridge()
        window.destroy()

    window.protocol("WM_DELETE_WINDOW", on_closing)

    # 初始化日志
    log_message("✅ 程序启动完成")
    log_message(f"ℹ️ 配置文件路径: {CONFIG_FILE}")
    log_message(f"ℹ️ 日志文件路径: {LOG_FILE}")

    window.mainloop()


# =============================
# 🚀 程序入口
# =============================
if __name__ == "__main__":
    # 检查管理员权限（Windows）
    if sys.platform == 'win32' and not elevate_admin():
        sys.exit(1)

    # 适配异步循环
    if sys.platform == 'win32':
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

    # 启动GUI
    init_gui()