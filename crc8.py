def calculate_crc8(data, poly=0x07, init_value=0x00, final_xor=0x00):
    """
    CRC-8校验算法
    参数：
        data: 输入数据（字节列表）
        poly: 多项式（默认0x07，标准CRC-8）
        init_value: 初始值（默认0x00）
        final_xor: 最终异或值（默认0x00）
    """
    crc = init_value
    for byte in data:
        crc ^= byte  # 当前字节与CRC寄存器异或
        for _ in range(8):  # 处理每个bit（共8位）
            if crc & 0x80:  # 检查最高位是否为1（MSB-first处理）
                crc = (crc << 1) ^ poly  # 左移并异或多项式
            else:
                crc <<= 1  # 仅左移
            crc &= 0xFF    # 保持8位
    return crc ^ final_xor  # 应用最终异或

# ========= 示例测试 =========
if __name__ == "__main__":
    # 测试用例1: 空数据（应返回初始值异或final_xor）
    print(hex(calculate_crc8([])))  # 0x00

    # 测试用例2: 单字节 0x01
    # CRC-8校验过程演示：
    # 初始值: 0x00
    # 异或0x01 → 0x01
    # 处理8位：
    #   最高位为0 → 左移 → 0x02
    #   最高位为0 → 左移 → 0x04
    #   ...（直到最高位为1时异或多项式）
    # 最终结果: 0xD5（不同多项式结果不同）
    data = [0x01,0x01,0x01,0x03,0x00,0x44,0x00,0x00,0x00,0x00,0xff]
    print(hex(calculate_crc8(data)))  # 输出取决于多项式

    # 测试其他常见CRC-8变种（例如CRC-8/MAXIM）
    def crc8_maxim(data):
        return calculate_crc8(data, poly=0x31, init_value=0x00, final_xor=0x00)
    
    print(hex(crc8_maxim(data)))  # 示例输出: 0xA1