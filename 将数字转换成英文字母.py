

def changeNumToChar(toSmallChar=None, toBigChar=None):
    # 把数字转换成相应的字符,1-->'A' 27-->'AA'
    init_number = 0
    increment = 0
    res_char = ''
    if not toSmallChar and not toBigChar:
       return ''
    else:
        if toSmallChar:
            init_number = toSmallChar
            increment = ord('a') - 1
        else:
            init_number = toBigChar
            increment = ord('A') - 1
    shang,yu = divmod(init_number, 26)
    if shang > 0:
        char1 = chr(shang + increment)
    else:
        char1 = ''
    char2 = chr(yu + increment)
    res_char = char1 + char2
    return res_char