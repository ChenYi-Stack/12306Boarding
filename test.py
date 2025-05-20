import re
#正则表达式测试


# 测试文本（含订单号12306）
test_text = """
尊敬的先生： 
您好！您于2025年05月05日在中国铁路客户服务中心网站(12306.cn) 成功购买了1张车票，
票款共计23.50元。订单号码12306。所购车票信息如下： 
1.，2025年05月05日16:20开，金华南站-缙云西站，G7347次列车，6车13C号，
二等座，成人票，票价23.5元，电子客票。 

尊敬的先生： 
您好！您于2025年04月17日在中国铁路客户服务中心网站(12306.cn) 成功购买了1张车票，
票款共计90.00元。订单号码12307。所购车票信息如下： 
1.，2025年04月17日10:01开，垫江站-万州北站，G1478次列车，1车6F号，
一等座，成人票，票价90.0元，电子客票。 
"""

# 修复后的正则表达式
pattern = r"(?:车次|次列车)[：:\s\u3000]*([GDKTZ]\d{1,4})(?=\s*(次|次列车|班次|运行线|$))|([GDKTZ]\d{1,4}|1[3-6]\d{3})"

def find_train_numbers(text, regex_pattern):
    try:
        regex = re.compile(regex_pattern, re.IGNORECASE)
        matches = regex.findall(text)
        results = [group1 or group3 for group1, _, group3 in matches]
        return results
    except re.error as e:
        return f"正则表达式错误: {e}"

if __name__ == "__main__":
    results = find_train_numbers(test_text, pattern)
    print("=== 车次号匹配结果 ===")
    if isinstance(results, list) and results:
        for i, number in enumerate(results, 1):
            print(f"{i}. {number}")
    else:
        print("未找到匹配的车次号")