import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, Frame
import serial
import serial.tools.list_ports
from threading import Thread, Event
import queue
from datetime import datetime

def calculate_crc8(data, poly=0x07, init_value=0x00, final_xor=0x00):
    """CRC-8校验算法"""
    crc = init_value
    for byte in data:
        crc ^= byte
        for _ in range(8):
            if crc & 0x80:
                crc = (crc << 1) ^ poly
            else:
                crc <<= 1
            crc &= 0xFF
    return crc ^ final_xor

class SerialApp:
    def __init__(self, master):
        self.master = master
        self.ser = None
        self.stop_event = Event()
        self.recv_queue = queue.Queue()
        self.setup_ui()
        self.update_ports()
        self.hex_mode = True  # 默认十六进制显示
        self.receive_count = 0  # 接收计数器

    def setup_ui(self):
        self.master.title("串口调试工具")
        
        # 串口配置
        config_frame = ttk.Frame(self.master)
        config_frame.pack(padx=10, pady=5, fill=tk.X)

        ttk.Label(config_frame, text="端口:").grid(row=0, column=0)
        self.port_cb = ttk.Combobox(config_frame, width=15)
        self.port_cb.grid(row=0, column=1, padx=5)
        ttk.Button(config_frame, text="刷新", command=self.update_ports).grid(row=0, column=2)

        ttk.Label(config_frame, text="波特率:").grid(row=0, column=3)
        self.baud_cb = ttk.Combobox(config_frame, values=['9600', '19200', '38400', '57600', '115200'], width=8)
        self.baud_cb.set('115200')
        self.baud_cb.grid(row=0, column=4, padx=5)

        self.connect_btn = ttk.Button(config_frame, text="打开", command=self.toggle_connection)
        self.connect_btn.grid(row=0, column=5)

        # 数据帧设置
        data_frame = ttk.LabelFrame(self.master, text="数据帧配置")
        data_frame.pack(padx=10, pady=5, fill=tk.X)

        self.entries = []
        for i in range(14):
            row = i // 7
            col = i % 7
            ttk.Label(data_frame, text=f"D{i}").grid(row=row, column=col*2, padx=2)
            entry = ttk.Entry(data_frame, width=4)
            entry.grid(row=row, column=col*2+1, padx=2)
            if i == 12:  # CRC字段设为只读
                entry.config(state='readonly')
            self.entries.append(entry)

        # 发送控制
        self.send_btn = ttk.Button(self.master, text="发送数据", command=self.send_data)
        self.send_btn.pack(pady=5)

        # 接收显示控制
        recv_control = Frame(self.master)
        recv_control.pack(fill=tk.X, padx=10)
        ttk.Button(recv_control, text="清空接收", command=self.clear_received).pack(side=tk.LEFT)
        self.display_mode = ttk.Combobox(recv_control, values=['Hex', 'ASCII'], width=6)
        self.display_mode.set('Hex')
        self.display_mode.pack(side=tk.LEFT, padx=5)
        self.display_mode.bind('<<ComboboxSelected>>', self.change_display_mode)
        
        # 接收显示
        self.recv_area = scrolledtext.ScrolledText(self.master, wrap=tk.WORD, state='disabled')
        self.recv_area.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
        
        # 状态栏
        self.status = ttk.Label(self.master, text="就绪", anchor=tk.W)
        self.status.pack(fill=tk.X, padx=10)

    def update_ports(self):
        ports = [port.device for port in serial.tools.list_ports.comports()]
        self.port_cb['values'] = ports
        if ports:
            self.port_cb.set(ports[0])

    def toggle_connection(self):
        if self.ser and self.ser.is_open:
            self.close_serial()
        else:
            self.open_serial()

    def open_serial(self):
        try:
            self.ser = serial.Serial(
                port=self.port_cb.get(),
                baudrate=int(self.baud_cb.get()),
                timeout=1
            )
            self.connect_btn.config(text="关闭")
            Thread(target=self.read_serial, daemon=True).start()
            self.update_status(f"已连接 {self.port_cb.get()} @ {self.baud_cb.get()}")
        except Exception as e:
            messagebox.showerror("错误", f"打开串口失败: {e}")

    def close_serial(self):
        self.stop_event.set()
        if self.ser:
            self.ser.close()
        self.connect_btn.config(text="打开")
        self.stop_event.clear()
        self.update_status("已断开连接")

    def send_data(self):
        if not self.ser or not self.ser.is_open:
            messagebox.showerror("错误", "请先打开串口")
            return

        try:
            data_bytes = []
            for i in range(14):
                if i == 12: continue
                value = self.entries[i].get().strip()
                if not value:
                    raise ValueError(f"D{i} 不能为空")
                data_bytes.append(int(value, 16))

            crc_data = data_bytes[1:12]
            crc = calculate_crc8(crc_data)
            
            data_bytes.insert(12, crc)
            final_data = bytes(data_bytes)
            
            self.entries[12].config(state='normal')
            self.entries[12].delete(0, tk.END)
            self.entries[12].insert(0, f"{crc:02X}")
            self.entries[12].config(state='readonly')

            self.ser.write(final_data)
            self.update_status(f"已发送 {len(final_data)} 字节")
        except ValueError as e:
            messagebox.showerror("输入错误", f"无效输入: {e}")
        except Exception as e:
            messagebox.showerror("发送错误", f"发送失败: {e}")

    def read_serial(self):
        while not self.stop_event.is_set() and self.ser.is_open:
            try:
                if self.ser.in_waiting:
                    data = self.ser.read(self.ser.in_waiting)
                    timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
                    self.recv_queue.put((timestamp, data))
                    self.master.after(10, self.process_received_data)
            except Exception as e:
                self.recv_queue.put(("ERROR", f"接收错误: {e}"))
                break

    def process_received_data(self):
        while not self.recv_queue.empty():
            item = self.recv_queue.get()
            if item[0] == "ERROR":
                messagebox.showerror("错误", item[1])
                self.close_serial()
                return
            
            timestamp, data = item
            self.receive_count += len(data)
            self.update_status(f"已接收 {self.receive_count} 字节")
            
            hex_str = ' '.join(f"{b:02X}" for b in data)
            ascii_str = ''.join([chr(b) if 32 <= b <= 126 else '.' for b in data])
            
            self.recv_area.config(state='normal')
            self.recv_area.insert(tk.END, f"[{timestamp}] ")
            if self.hex_mode:
                self.recv_area.insert(tk.END, hex_str)
            else:
                self.recv_area.insert(tk.END, ascii_str)
            self.recv_area.insert(tk.END, '\n')
            self.recv_area.see(tk.END)
            self.recv_area.config(state='disabled')

    def change_display_mode(self, event):
        self.hex_mode = (self.display_mode.get() == 'Hex')
        self.recv_area.config(state='normal')
        self.recv_area.delete(1.0, tk.END)
        self.recv_area.config(state='disabled')
        self.receive_count = 0

    def clear_received(self):
        self.recv_area.config(state='normal')
        self.recv_area.delete(1.0, tk.END)
        self.recv_area.config(state='disabled')
        self.update_status("接收区已清空")

    def update_status(self, message):
        self.status.config(text=message)

if __name__ == "__main__":
    root = tk.Tk()
    app = SerialApp(root)
    root.mainloop()