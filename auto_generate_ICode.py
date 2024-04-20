import itertools

# 字符集
chars = "0123456789abcdef"

# 读取文件
with open('加密注册码.txt', 'r') as file:
    lines = file.readlines()

# 处理每一行
final_results = []
for line in lines:
    line = line.strip()
    # 计算星号的数量
    star_count = line.count('*')
    if star_count >= 2:
        # 获取所有可能的替换
        indices = [i for i, char in enumerate(line) if char == '*']
        for replacement in itertools.product(chars, repeat=star_count):
            # 创建一个字符列表以便可以修改
            temp_list = list(line)
            # 替换所有星号
            for idx, char in zip(indices, replacement):
                temp_list[idx] = char
            # 将列表转回字符串
            final_results.append(''.join(temp_list))

# 存储结果到新文件
with open('注册码.txt', 'w') as file:
    for result in final_results:
        file.write(result + '\n')

