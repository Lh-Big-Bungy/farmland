import cn2an

def number_to_chinese(amount):
    """将数字金额转换为中文大写金额"""
    amount_str = f"{amount:.2f}"  # 保留两位小数
    return cn2an.an2cn(amount_str, "rmb")  # 转换为人民币大写

# 示例
num = 411764.14
print(number_to_chinese(num))  # 输出：肆拾壹万壹仟柒佰陆拾肆元壹角肆分
