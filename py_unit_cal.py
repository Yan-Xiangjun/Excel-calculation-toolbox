import sympy as sp
import re
import sys


def d_cal(d_formu, unit_list):
    for u in unit_list[0]:  # 生成符号变量
        exec("%s=sp.Symbol('%s',positive=True)" % (u, u))
    for i in range(1, len(unit_list)):
        d_formu = d_formu.replace(unit_list[i][0], unit_list[i][1])

    d_formu = d_formu.replace('·', '*')
    d_formu = d_formu.replace('^', '**')
    d_formu = d_formu.replace(',', '+')
    d_formu = re.sub(r'(?<!\*\*\()-', '+', d_formu)  # 防止量纲相减之后抵消，变成0
    d_formu = re.sub(r'(?<!\*\*)0', '1', d_formu)
    out = eval('sp.expand(%s)' % d_formu)
    out = str(out)
    out = re.sub(r'sqrt\([0-9.]+\)', '1', out)  # 将sqrt(纯数字)替换成1
    out = eval('sp.expand(%s)' % out)
    out = str(out)
    # 如果out是纯数字，说明为无量纲
    if re.search(r'^[0-9.]+$', out):
        return '1'
    out = out.replace('**', '^')
    out = re.sub(r'[*/][0-9.]+|[0-9.]+[*]', '', out)
    out = re.sub(r'[0-9.]+/(?P<unit>.+)\^*(?P<num>[0-9.]+)*', to_exponent, out)
    out = out.replace('*', '·')
    return out


def to_exponent(match):
    if match.group('num') == None:
        return "%s^(-1)" % match.group('unit')
    return "%s^(-%s)" % (match.group('unit'), match.group('num'))


unit_list = [['kN', 'm', 'N', 'mm', 'K'], ['(mm)/10^3', '(m)'], ['(m)*10^3', '(mm)'], ['(kN)*10^3', '(N)'],
             ['(kN·m)*10^6', '(N·mm)'], ['(kN/m^2)/10^3', '(N/mm^2)']]
try:
    print(d_cal(sys.argv[1], unit_list))
except Exception:
    print('【Error】')
