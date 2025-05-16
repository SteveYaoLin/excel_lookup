# filepath: d:\Working\JIELI\repo\modbus_crc_verilog\crc_python\main.py
# -*- coding: utf-8 -*-
import struct

def calculate_modbus_crc(data):
    crc = 0xFFFF
    for byte in data:
        crc ^= byte
        for _ in range(8):
            if crc & 0x0001:
                crc = (crc >> 1) ^ 0xA001
            else:
                crc >>= 1
    return crc

def main():
    a = [0xf0, 0x03, 0x00, 0x01, 0x00, 0x01]
    b = [0x01]
    crc16 = calculate_modbus_crc(a)
    print(hex(crc16).lower())
    crc16 = calculate_modbus_crc(b)
    print(hex(crc16).lower())
if __name__ == '__main__':
    print("hello world\n")
    main()