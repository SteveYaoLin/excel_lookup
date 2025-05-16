def calculate_modbus_crc(data):
    crc = 0xFFFF
    print(f"初始CRC: 0x{crc:04x}")
    for byte_idx, byte in enumerate(data):
        print(f"\n处理第 {byte_idx} 个字节: 0x{byte:02x}")
        crc ^= byte
        print(f"  异或字节后 CRC = 0x{crc:04x}")
        for bit in range(8):  # 处理每个bit (从LSB到MSB)
            print(f"  处理字节位 [bit {bit}]：")
            print(f"    当前CRC: 0x{crc:04x} (二进制: {crc:016b})")
            if crc & 0x0001:
                print("    最低位为1 → 右移并异或多项式 0xA001")
                crc = (crc >> 1) ^ 0xA001
            else:
                print("    最低位为0 → 仅右移")
                crc >>= 1
            print(f"    处理后CRC: 0x{crc:04x}\n{'-'*40}")
    print("\n最终CRC: 0x{:04x}".format(crc))
    return crc

def main():
    a = [0xf0, 0x03, 0x00, 0x01, 0x00, 0x01]
    b = [0xf0]
    
    print("===== 计算数据a的CRC =====")
    crc_a = calculate_modbus_crc(a)
    print("结果: " + hex(crc_a).lower() + "\n")
    
    print("\n===== 计算数据b的CRC =====")
    crc_b = calculate_modbus_crc(b)
    print("结果: " + hex(crc_b).lower())

if __name__ == '__main__':
    main()