# @Author: YiLong Li
# @Date: 2026/11/13 10:14
# @File: main.py
# @Software: PyCharm


"""
项目名称：加班数据核对工作台（OvertimeKit）

项目介绍：
将网页模板、Excel 解析库、业务脚本及 Logo 封装为独立 HTML 工作台。
生成后无需持续运行 Python，可单独复制 HTML 离线使用。
导入加班明细与考勤汇总表，按员工编码累计四类加班时数，
集中核对差异及异常，人工确认后导出结果，不修改原始文件。

主要功能：
双表导入、工作表选择、字段映射、分页预览、四类加班汇总、
差异核对、人工确认、进度保存与恢复、Excel 结果导出。

目录说明：
assets/         网页模板、业务脚本、Excel 解析库、Logo 及许可证
help/           中文操作手册（Markdown）
output/         生成的独立 HTML；不存在时自动创建

运行方式：
python main.py
python main.py --no-browser
uv run main.py

构建仅使用 Python 标准库，无需安装额外依赖。
默认生成 output/加班数据核对工作台.html。
"""


import argparse
import base64
from pathlib import Path
import webbrowser

BASE = Path(__file__).resolve().parent

def build(output):
    assets = BASE / 'assets'
    text = (assets / 'dashboard_template.html').read_text(encoding='utf-8')
    # 在嵌入脚本前写入许可，避免误替换第三方脚本字符串中的 HTML 标签。
    license_text = (assets / 'SheetJS-LICENSE.txt').read_text(encoding='utf-8').replace('--', '—')
    text = text.replace('</head>', '<!--\n' + license_text + '\n--></head>', 1)
    for name in ('xlsx.full.min.js', 'core.js', 'app.js'):
        script = (assets / name).read_text(encoding='utf-8').replace('</script', r'<\/script')
        text = text.replace(f'<script src="{name}"></script>', '<script>' + script + '</script>')
    logo = base64.b64encode((assets / 'cml_logo_white.svg').read_bytes()).decode()
    text = text.replace('src="cml_logo_white.svg"', f'src="data:image/svg+xml;base64,{logo}"')
    output.mkdir(parents=True, exist_ok=True)
    path = output / '加班数据核对工作台.html'
    path.write_text(text, encoding='utf-8')
    return path.resolve()

def main():
    parser = argparse.ArgumentParser(description='生成加班数据核对工作台')
    parser.add_argument('--no-browser', action='store_true')
    parser.add_argument('--output', type=Path, default=BASE / 'output')
    args = parser.parse_args()
    try:
        path = build(args.output)
    except OSError as exc:
        parser.exit(1, f'生成失败，请完整解压项目并检查读写权限：{exc}\n')
    print(f'已生成：{path}')
    if not args.no_browser:
        webbrowser.open(path.as_uri())

if __name__ == '__main__':
    main()