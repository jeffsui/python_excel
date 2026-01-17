# excel_tool.py
import argparse
import sys
from pathlib import Path

import pandas as pd


def merge_excels(file_paths, output_path, merge_by="rows", sheet_name=None):
    """
    合并多个 Excel 文件
    :param file_paths: Excel 文件路径列表
    :param output_path: 输出文件路径
    :param merge_by: "rows"（纵向拼接）或 "columns"（横向拼接）
    :param sheet_name: 指定读取的 sheet 名（默认第一个 sheet）
    """
    dfs = []
    for fp in file_paths:
        
        try:
            # 读取 Excel 文件，获取第一个 sheet
            data = pd.read_excel(fp, sheet_name=sheet_name, engine="openpyxl")

            # 处理返回类型：可能是 DataFrame 或 dict（多 sheet 情况）
            if isinstance(data, dict):
                # 取第一个 sheet 的数据
                sheet_names = list(data.keys())
                if sheet_names:
                    df = data[sheet_names[0]]
                    if not isinstance(df, pd.DataFrame):
                        print(f"⚠️  文件格式异常 {fp}: 第一个 sheet 不是 DataFrame")
                        continue
                else:
                    print(f"⚠️  文件为空 {fp}")
                    continue
            elif isinstance(data, pd.DataFrame):
                df = data
            else:
                print(f"⚠️  文件格式异常 {fp}: 返回类型 {type(data)}")
                continue

            df["_source_file"] = Path(fp).name  # 可选：标记来源文件
            dfs.append(df)
            print(f"✅ 已加载: {fp} ({df.shape[0]} 行 × {df.shape[1]} 列)")
        except Exception as e:
            print(f"⚠️  加载失败 {fp}: {e}")
            continue

    if not dfs:
        print("❌ 无有效数据，合并终止。")
        return

    if merge_by == "rows":  # 行合并
        result = pd.concat(dfs, ignore_index=True)
    elif merge_by == "columns":  # 列合并
        result = pd.concat(dfs, axis=1)
    else:
        raise ValueError("merge_by 必须是 'rows' 或 'columns'")

    result.to_excel(output_path, index=False)
    print(f"✅ 合并完成！结果已保存至: {output_path}")


def statistics_excel(file_path, group_cols=None, agg_cols=None, output_path=None):
    """
    对 Excel 做统计分析
    :param file_path: 输入 Excel 路径
    :param group_cols: 分组列（list 或 None）
    :param agg_cols: 待统计的数值列（list）
    :param output_path: 输出路径（默认原文件名 + _stats.xlsx）
    """
    try:
        df = pd.read_excel(file_path, engine="openpyxl")
        print(f"📊 加载数据: {df.shape[0]} 行 × {df.shape[1]} 列")
        print(f"📋 所有列名: {', '.join(df.columns.tolist())}")

        if output_path is None:
            stem = Path(file_path).stem
            output_path = f"{stem}_stats.xlsx"

        with pd.ExcelWriter(output_path) as writer:
            # 基础描述性统计
            desc = df.describe(include="all").T
            desc.to_excel(writer, sheet_name="summary")

            # 分组统计（如指定）
            if group_cols or agg_cols:
                # 检查列名是否存在
                all_cols = df.columns.tolist()
                missing_cols = []

                if group_cols:
                    missing_cols.extend(
                        [col for col in group_cols if col not in all_cols]
                    )
                if agg_cols:
                    missing_cols.extend(
                        [col for col in agg_cols if col not in all_cols]
                    )

                if missing_cols:
                    print(f"⚠️  以下列名未找到: {', '.join(missing_cols)}")
                    print(f"   可用的列名: {', '.join(all_cols)}")
                else:
                    # 只有当所有列名都存在时才进行分组统计
                    if group_cols and agg_cols:
                        try:
                            grouped = (
                                df.groupby(group_cols)[agg_cols]
                                .agg(["sum", "mean", "count"])
                                .reset_index()
                            )
                            # 处理 MultiIndex 列名
                            if isinstance(grouped.columns, pd.MultiIndex):
                                # 将 MultiIndex 转换为单层列名
                                grouped.columns = [
                                    "_".join(col).strip()
                                    for col in grouped.columns.values
                                ]
                            grouped.to_excel(
                                writer, sheet_name="grouped_stats", index=False
                            )
                            print("✅ 分组统计已生成")
                        except Exception as e:
                            print(f"⚠️  分组统计失败: {e}")

            # 计数（各列非空值数量）
            counts = df.count().to_frame("non_null_count")
            counts.to_excel(writer, sheet_name="counts")

        print(f"✅ 统计完成！结果已保存至: {output_path}")

    except Exception as e:
        print(f"❌ 统计失败: {e}")


def main():
    parser = argparse.ArgumentParser(
        description="📊 Excel 处理工具 — 合并 & 统计",
        epilog="示例：\n"
        "  合并：python excel_tool.py merge file1.xlsx file2.xlsx -o merged.xlsx\n"
        "  统计：python excel_tool.py stats data.xlsx --group_cols=部门 --agg_cols=销售额,利润",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    subparsers = parser.add_subparsers(dest="command", help="子命令")

    # 合并子命令
    merge_parser = subparsers.add_parser("merge", help="合并多个 Excel 文件")
    merge_parser.add_argument("files", nargs="+", help="Excel 文件路径（至少两个）")
    merge_parser.add_argument(
        "-o", "--output", default="merged_output.xlsx", help="输出文件路径"
    )
    merge_parser.add_argument(
        "--by",
        choices=["rows", "columns"],
        default="rows",
        help="合并方式：rows（默认）或 columns",
    )
    merge_parser.add_argument("--sheet", help="指定 sheet 名（默认读取第一个）")

    # 统计子命令
    stats_parser = subparsers.add_parser("stats", help="对 Excel 做统计分析")
    stats_parser.add_argument("file", help="输入 Excel 文件路径")
    stats_parser.add_argument("--group_cols", nargs="*", help="分组列名（空格分隔）")
    stats_parser.add_argument(
        "--agg_cols", nargs="*", help="待聚合的数值列（空格分隔）"
    )
    stats_parser.add_argument("-o", "--output", help="输出文件路径（默认自动命名）")

    # 预处理参数：支持逗号分隔
    if len(sys.argv) > 1:
        for i, arg in enumerate(sys.argv):
            if arg.startswith("--group_cols="):
                sys.argv[i] = "--group_cols"
                sys.argv.insert(i + 1, *(arg.split("=")[1].split(",")))
            elif arg.startswith("--agg_cols="):
                sys.argv[i] = "--agg_cols"
                sys.argv.insert(i + 1, *(arg.split("=")[1].split(",")))

    if len(sys.argv) == 1:
        parser.print_help()
        return

    args = parser.parse_args()

    if args.command == "merge":
        if len(args.files) < 2:
            print("❌ 至少需要两个 Excel 文件进行合并！")
            return
        merge_excels(args.files, args.output, merge_by=args.by, sheet_name=args.sheet)

    elif args.command == "stats":
        statistics_excel(
            file_path=args.file,
            group_cols=args.group_cols,
            agg_cols=args.agg_cols,
            output_path=args.output,
        )

    else:
        parser.print_help()


if __name__ == "__main__":
    main()
