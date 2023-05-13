def a1_notation(row: int, col_a1_notation: int):
    alphabet = ["Z", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y"]
    digit = []
    while col_a1_notation > 26:
        digit.append(col_a1_notation % 26)
        col_a1_notation = col_a1_notation // 26
    digit.append(col_a1_notation % 26)
    digit_alphabet = [alphabet[mod] for mod in reversed(digit)]
    output = "".join(digit_alphabet)+str(row)
    return output


def list_differ(a: list, b: list):
    set_a = set(a)
    set_b = set(b)
    return list(set_a.symmetric_difference(set_b))


def ws_cell(ws, row_num_cell, col_num_cell):
    cell = ws.cell(row=row_num_cell, column=col_num_cell)
    return cell.value


def identify_element_from_path(code_path: list):
    for path in reversed(code_path):
        if "上海" in path or "上交所" in path:
            return "上海"
        elif "深圳" in path or "深交所" in path:
            return "深圳"
        elif "中金所" in path:
            return "中金所"
    return None


def identify_bank_element_from_path(code_path: list):
    bank_element = None
    for i, element in enumerate(reversed(code_path)):
        if '银行' in element:
            bank_element = element
            break
    return bank_element
