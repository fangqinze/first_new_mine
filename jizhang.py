import json
import os
from datetime import datetime
from collections import defaultdict


class SimpleFinanceTracker:
    def __init__(self, filename="finance_data.json"):
        """初始化记账程序"""
        self.filename = filename
        self.data = self.load_data()

    def load_data(self):
        """从文件加载数据"""
        if os.path.exists(self.filename):
            try:
                with open(self.filename, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except (json.JSONDecodeError, FileNotFoundError):
                return {"records": [], "categories": {"收入": [], "支出": []}}
        return {"records": [], "categories": {"收入": [], "支出": []}}

#records 记录  categories 类别
    def save_data(self):
        """保存数据到文件"""
        with open(self.filename, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    def get_current_date(self):
        """获取当前日期字符串"""
        return datetime.now().strftime("%Y-%m-%d")

    def get_current_month(self):
        """获取当前月份字符串"""
        return datetime.now().strftime("%Y-%m")

    def add_record(self, record_type, amount, description, category=None, date=None):
        """添加一条记录"""
        if record_type not in ["收入", "支出"]:
            print("错误: 类型必须是'收入'或'支出'")
            return False

        if not isinstance(amount, (int, float)) or amount <= 0:
            print("错误: 金额必须是正数")
            return False

        if not description.strip():
            print("错误: 描述不能为空")
            return False

        if date is None:
            date = self.get_current_date()

        if category is None:
            category = input(f"请输入{record_type}类别 (按回车跳过): ").strip() or "未分类"

        # 如果类别不存在，添加到类别列表
        if category not in self.data["categories"][record_type]:
            self.data["categories"][record_type].append(category)

        record = {
            "date": date,
            "type": record_type,
            "amount": amount,
            "description": description,
            "category": category
        }

        self.data["records"].append(record)
        self.save_data()
        print(f"✓ 已添加{record_type}记录: {date} - {description} - ¥{amount:.2f} ({category})")
        return True

    def show_records(self, month=None, limit=50):
        """显示记录"""
        records = self.data["records"]

        if not records:
            print("暂无记录")
            return

        # 筛选月份
        filtered_records = records
        if month:
            filtered_records = [r for r in records if r["date"].startswith(month)]

        if not filtered_records:
            print(f"{month}月份无记录")
            return

        # 显示记录
        print(f"\n{'=' * 60}")
        print(f"{'日期':<12} {'类型':<6} {'金额':<12} {'描述':<20} {'类别':<10}")
        print(f"{'-' * 60}")

        # 限制显示数量
        display_records = filtered_records[-limit:]  # 显示最新的记录

        for i, record in enumerate(reversed(display_records), 1):
            print(f"{record['date']:<12} {record['type']:<6} ¥{record['amount']:<10.2f} "
                  f"{record['description'][:18]:<20} {record['category'][:8]:<10}")

        print(f"{'=' * 60}")

        # 显示统计
        if month:
            self.show_monthly_summary(month)

    def show_monthly_summary(self, month=None):
        """显示月度统计"""
        if month is None:
            month = self.get_current_month()

        records = self.data["records"]
        month_records = [r for r in records if r["date"].startswith(month)]

        if not month_records:
            print(f"{month}月份无记录")
            return

        # 计算统计
        income_total = sum(r["amount"] for r in month_records if r["type"] == "收入")
        expense_total = sum(r["amount"] for r in month_records if r["type"] == "支出")
        balance = income_total - expense_total

        # 按类别统计
        income_by_category = defaultdict(float)
        expense_by_category = defaultdict(float)

        for record in month_records:
            if record["type"] == "收入":
                income_by_category[record["category"]] += record["amount"]
            else:
                expense_by_category[record["category"]] += record["amount"]

        # 显示统计
        print(f"\n📊 {month}月份统计")
        print(f"{'=' * 40}")
        print(f"总收入: ¥{income_total:.2f}")
        if income_by_category:
            print("  收入分类明细:")
            for category, amount in sorted(income_by_category.items(), key=lambda x: x[1], reverse=True):
                print(f"    {category}: ¥{amount:.2f}")

        print(f"\n总支出: ¥{expense_total:.2f}")
        if expense_by_category:
            print("  支出分类明细:")
            for category, amount in sorted(expense_by_category.items(), key=lambda x: x[1], reverse=True):
                print(f"    {category}: ¥{amount:.2f}")

        print(f"\n{'-' * 40}")
        print(f"月度结余: ¥{balance:.2f}")

        if balance > 0:
            print("💹 财务状况: 良好 (有结余)")
        elif balance == 0:
            print("⚖️ 财务状况: 平衡 (收支平衡)")
        else:
            print("🔻 财务状况: 斥资 (支出大于收入)")

    def show_yearly_summary(self, year=None):
        """显示年度统计"""
        if year is None:
            year = datetime.now().strftime("%Y")

        records = self.data["records"]
        year_records = [r for r in records if r["date"].startswith(str(year))]

        if not year_records:
            print(f"{year}年度无记录")
            return

        # 月度统计
        monthly_stats = defaultdict(lambda: {"收入": 0, "支出": 0})

        for record in year_records:
            month = record["date"][:7]  # 获取年月
            monthly_stats[month][record["type"]] += record["amount"]

        # 年度总计
        year_income = sum(stats["收入"] for stats in monthly_stats.values())
        year_expense = sum(stats["支出"] for stats in monthly_stats.values())
        year_balance = year_income - year_expense

        # 显示年度统计
        print(f"\n📅 {year}年度统计")
        print(f"{'=' * 60}")
        print(f"{'月份':<10} {'收入':<12} {'支出':<12} {'结余':<12} {'状态':<8}")
        print(f"{'-' * 60}")

        for month in sorted(monthly_stats.keys()):
            income = monthly_stats[month]["收入"]
            expense = monthly_stats[month]["支出"]
            balance = income - expense

            if balance > 0:
                status = "盈余"
            elif balance < 0:
                status = "赤字"
            else:
                status = "平衡"

            print(f"{month:<10} ¥{income:<10.2f} ¥{expense:<10.2f} ¥{balance:<10.2f} {status:<8}")

        print(f"{'=' * 60}")
        print(f"年度总收入: ¥{year_income:.2f}")
        print(f"年度总支出: ¥{year_expense:.2f}")
        print(f"年度总结余: ¥{year_balance:.2f}")

        if year_balance > 0:
            print(f"💹 年度财务状况: 良好，结余 ¥{year_balance:.2f}")
        else:
            print(f"⚠️ 年度财务状况: 需注意，斥资 ¥{-year_balance:.2f}")

    def export_to_text(self, filename=None):
        """导出数据到文本文件"""
        if filename is None:
            filename = f"finance_export_{self.get_current_date()}.txt"

        with open(filename, 'w', encoding='utf-8') as f:
            f.write("个人财务记录\n")
            f.write("=" * 50 + "\n\n")

            # 写入所有记录
            f.write("所有交易记录:\n")
            f.write("-" * 50 + "\n")
            for record in self.data["records"]:
                f.write(f"{record['date']} | {record['type']} | ¥{record['amount']:.2f} | "
                        f"{record['description']} | {record['category']}\n")

            # 写入月度统计
            f.write("\n\n月度统计:\n")
            f.write("=" * 50 + "\n")

            # 按月份分组
            monthly_records = defaultdict(list)
            for record in self.data["records"]:
                month = record["date"][:7]  # 获取年月
                monthly_records[month].append(record)

            for month in sorted(monthly_records.keys()):
                month_income = sum(r["amount"] for r in monthly_records[month] if r["type"] == "收入")
                month_expense = sum(r["amount"] for r in monthly_records[month] if r["type"] == "支出")
                month_balance = month_income - month_expense

                f.write(f"\n{month}:\n")
                f.write(f"  收入: ¥{month_income:.2f}\n")
                f.write(f"  支出: ¥{month_expense:.2f}\n")
                f.write(f"  结余: ¥{month_balance:.2f}\n")

        print(f"✓ 数据已导出到 {filename}")

    def delete_record(self, index):
        """删除指定记录"""
        if 0 <= index < len(self.data["records"]):
            deleted = self.data["records"].pop(index)
            self.save_data()
            print(f"✓ 已删除记录: {deleted['date']} - {deleted['description']}")
            return True
        else:
            print("错误: 无效的记录索引")
            return False


def main():
    """主程序"""
    tracker = SimpleFinanceTracker()

    while True:
        print("\n" + "=" * 50)
        print("个人记账系统")
        print("=" * 50)
        print("1. 添加收入记录")
        print("2. 添加支出记录")
        print("3. 查看所有记录")
        print("4. 查看本月记录")
        print("5. 查看本月统计")
        print("6. 查看年度统计")
        print("7. 导出数据到文件")
        print("8. 查看类别列表")
        print("9. 删除最新记录")
        print("0. 退出程序")
        print("-" * 50)

        try:
            choice = input("请选择操作 (0-9): ").strip()

            if choice == "0":
                print("感谢使用，再见！")
                break

            elif choice == "1":
                print("\n添加收入记录")
                print("-" * 30)
                try:
                    amount = float(input("收入金额: "))
                    description = input("收入描述 (如: 工资/奖金/兼职): ")
                    date = input("日期 (YYYY-MM-DD, 直接回车使用今天): ")
                    if not date:
                        date = None
                    tracker.add_record("收入", amount, description, date=date)
                except ValueError:
                    print("错误: 金额必须为数字")

            elif choice == "2":
                print("\n添加支出记录")
                print("-" * 30)
                try:
                    amount = float(input("支出金额: "))
                    description = input("支出描述 (如: 餐饮/购物/交通): ")
                    date = input("日期 (YYYY-MM-DD, 直接回车使用今天): ")
                    if not date:
                        date = None
                    tracker.add_record("支出", amount, description, date=date)
                except ValueError:
                    print("错误: 金额必须为数字")

            elif choice == "3":
                print("\n所有记录 (最近50条)")
                tracker.show_records()

            elif choice == "4":
                month = input("请输入月份 (YYYY-MM, 直接回车查看本月): ").strip()
                if not month:
                    month = tracker.get_current_month()
                tracker.show_records(month=month)

            elif choice == "5":
                month = input("请输入月份 (YYYY-MM, 直接回车查看本月): ").strip()
                if not month:
                    month = tracker.get_current_month()
                tracker.show_monthly_summary(month)

            elif choice == "6":
                year = input("请输入年份 (YYYY, 直接回车查看今年): ").strip()
                if not year:
                    year = datetime.now().strftime("%Y")
                tracker.show_yearly_summary(year)

            elif choice == "7":
                tracker.export_to_text()

            elif choice == "8":
                print("\n收入类别:", ", ".join(tracker.data["categories"]["收入"]) or "暂无")
                print("支出类别:", ", ".join(tracker.data["categories"]["支出"]) or "暂无")

            elif choice == "9":
                if tracker.data["records"]:
                    print(f"确定要删除最新记录吗?")
                    print(f"记录: {tracker.data['records'][-1]}")
                    confirm = input("输入 'y' 确认删除: ").lower()
                    if confirm == 'y':
                        tracker.delete_record(-1)
                else:
                    print("暂无记录可删除")

            else:
                print("无效选择，请重新输入")

        except KeyboardInterrupt:
            print("\n\n程序被用户中断")
            break
        except Exception as e:
            print(f"发生错误: {e}")


if __name__ == "__main__":
    # 程序入口
    print("正在启动记账程序...")
    print("数据将保存在 finance_data.json 文件中")
    main()