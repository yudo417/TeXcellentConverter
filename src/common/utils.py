from PyQt5.QtGui import QColor


def color_to_tikz_rgb(color):
    """QColorオブジェクトをTikZ互換のRGB形式に変換する"""
    # red, green, blueの一般的な色名はそのまま使用
    if color == QColor('red'):
        return 'red'
    elif color == QColor('green'):
        return 'green'
    elif color == QColor('blue'):
        return 'blue'
    elif color == QColor('black'):
        return 'black'
    elif color == QColor('yellow'):
        return 'yellow'
    elif color == QColor('cyan'):
        return 'cyan'
    elif color == QColor('magenta'):
        return 'magenta'
    elif color == QColor('orange'):
        return 'orange'
    elif color == QColor('purple'):
        return 'purple'
    elif color == QColor('brown'):
        return 'brown'
    elif color == QColor('gray'):
        return 'gray'
    # それ以外の色はRGB値として変換（0-255から0-1へ）
    else:
        r = color.red() / 255.0
        g = color.green() / 255.0
        b = color.blue() / 255.0
        return f"color = {{rgb,255:red,{color.red()};green,{color.green()};blue,{color.blue()}}}"


def safe_float_conversion(value):
    """安全に値をfloatに変換する"""
    if value is None or value == '':
        return None
    try:
        return float(value)
    except (ValueError, TypeError):
        return None


def escape_latex_special_chars(text):
    """LaTeX特殊文字をエスケープする"""
    if not isinstance(text, str):
        text = str(text)
    
    for char in ['&', '%', '$', '#', '_', '{', '}', '~', '^', '\\']:
        if char in text:
            text = text.replace(char, '\\' + char)
    return text 