import os
import sys
import json
import argparse
from pathlib import Path
from config_parser import ConfigParser
from excel_generator import ExcelGenerator


def main():
    parser = argparse.ArgumentParser(description='Excel结构生成工具')
    parser.add_argument('config', help='配置文件路径 (JSON格式)')
    parser.add_argument('-o', '--output', help='输出文件路径 (可选，默认使用配置文件中的output_file)')
    parser.add_argument('-d', '--output-dir', help='输出目录 (可选，默认为target目录)')

    args = parser.parse_args()

    config_path = Path(args.config)

    if not config_path.exists():
        print(f"错误: 配置文件不存在: {config_path}")
        sys.exit(1)

    print(f"读取配置文件: {config_path}")

    try:
        config_parser = ConfigParser()
        config = config_parser.parse_file(str(config_path))

        print("验证配置...")
        errors = config_parser.validate(config)
        if errors:
            print("配置验证失败:")
            for error in errors:
                print(f"  - {error}")
            sys.exit(1)

        print("配置验证通过")

        output_dir = args.output_dir or 'target'
        os.makedirs(output_dir, exist_ok=True)

        output_file = args.output or config.output_file
        if not os.path.isabs(output_file):
            output_path = os.path.join(output_dir, output_file)
        else:
            output_path = output_file

        print(f"生成Excel文件: {output_path}")

        generator = ExcelGenerator()
        generator.generate(config, output_path)

        print(f"Excel文件生成成功: {output_path}")

    except json.JSONDecodeError as e:
        print(f"错误: 配置文件JSON格式错误: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
