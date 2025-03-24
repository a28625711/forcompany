def main():
    print("请输入多行文本，按两次回车结束输入：")
    lines = []
    while True:
        line = input()
        if line == "":  # 检测到空行
            if not lines or lines[-1] == "":  # 如果上一行也是空行，则结束输入
                break
            else:
                lines.append("")
        else:
            lines.append(line)

    # 将换行符替换为空格
    result = " ".join(line for line in lines if line.strip())  # 跳过空行
    print("\n处理后的文本：")
    print(result)


if __name__ == "__main__":
    main()
