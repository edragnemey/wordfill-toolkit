"""同一模板生成两种交付方案；运行位置不限，输出默认放到项目 output 目录。"""
from pathlib import Path
import argparse
import json
import sys

HERE = Path(__file__).resolve().parent
TOOL = HERE.parents[1]
sys.path.insert(0, str(TOOL))
from wordfill.engine import fill_template


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-dir', type=Path, default=TOOL.parent / 'output' / '智能办公空间')
    args = parser.parse_args()
    for version in ('轻量版', '旗舰版'):
        data = json.loads((HERE / f'{version}数据.json').read_text(encoding='utf-8'))
        expected = sum(item['数量'] * item['单价'] for item in data['设备'])
        if any(item['小计'] != item['数量'] * item['单价'] for item in data['设备']):
            raise ValueError(f'{version}：请先修正设备小计')
        if data['预算']['设备'] != expected or data['预算']['合计'] != expected + data['预算']['服务']:
            raise ValueError(f'{version}：请先修正预算合计')
        output = args.output_dir.resolve() / f'{version}交付方案.docx'
        result = fill_template(str(HERE / '交付方案模板.docx'), data, str(output), strict=True)
        if result.leftover or result.warnings:
            raise RuntimeError(result.summary())
        print(f'已生成：{output}')
        print(result.summary())
    return 0


if __name__ == '__main__':
    sys.exit(main())
