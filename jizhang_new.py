import json
import os
import hashlib
import base64
from datetime import datetime, timedelta
from collections import defaultdict
import csv
from pathlib import Path
import warnings

warnings.filterwarnings('ignore')

# 尝试导入可选的第三方库
try:
    import pandas as pd

    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False
    print("警告: 未安装pandas，部分高级功能将不可用")

try:
    import matplotlib.pyplot as plt

    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False
    print("警告: 未安装matplotlib，图表功能将不可用")

try:
    from openpyxl import Workbook

    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    print("警告: 未安装openpyxl，Excel导出功能将不可用")

try:
    from jinja2 import Template

    JINJA2_AVAILABLE = True
except ImportError:
    JINJA2_AVAILABLE = False
    print("警告: 未安装Jinja2，HTML报告功能将不可用")

# 添加中文字体支持
try:
    if MATPLOTLIB_AVAILABLE:
        plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
        plt.rcParams['axes.unicode_minus'] = False
except:
    pass


class SimpleFinanceTracker:
    def __init__(self, filename="finance_data.json", user_id="default"):
        """初始化记账程序，添加用户支持和加密选项"""
        self.filename = filename
        self.user_id = user_id
        self.data = self.load_data()
        # 初始化预算数据
        if "budgets" not in self.data:
            self.data["budgets"] = {}
        # 初始化提醒设置
        if "reminders" not in self.data:
            self.data["reminders"] = {}

#budgets 预算  reminders 提醒
    def load_data(self):
        """从文件加载数据"""
        if os.path.exists(self.filename):
            try:
                with open(self.filename, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    # 确保数据包含必要字段
                    if "records" not in data:
                        data["records"] = []
                    if "categories" not in data:
                        data["categories"] = {"收入": [], "支出": []}
                    if "budgets" not in data:
                        data["budgets"] = {}
                    if "reminders" not in data:
                        data["reminders"] = {}
                    return data
            except (json.JSONDecodeError, FileNotFoundError):
                return {"records": [], "categories": {"收入": [], "支出": []}, "budgets": {}, "reminders": {}}
        return {"records": [], "categories": {"收入": [], "支出": []}, "budgets": {}, "reminders": {}}

    def save_data(self):
        """保存数据到文件"""
        try:
            with open(self.filename, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"保存数据时出错: {e}")
            return False

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

        # 验证日期格式
        try:
            datetime.strptime(date, "%Y-%m-%d")
        except ValueError:
            print("错误: 日期格式应为 YYYY-MM-DD")
            return False

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
        if self.save_data():
            print(f"✓ 已添加{record_type}记录: {date} - {description} - ¥{amount:.2f} ({category})")
            return True
        return False

    def show_records(self, month=None, limit=50):
        """显示记录"""
        records = self.data.get("records", [])

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

        records = self.data.get("records", [])
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

        records = self.data.get("records", [])
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
                status = "斥资"
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

        try:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write("个人财务记录\n")
                f.write("=" * 50 + "\n\n")

                # 写入所有记录
                f.write("所有交易记录:\n")
                f.write("-" * 50 + "\n")
                records = self.data.get("records", [])
                for record in records:
                    f.write(f"{record['date']} | {record['type']} | ¥{record['amount']:.2f} | "
                            f"{record['description']} | {record['category']}\n")

                # 写入月度统计
                f.write("\n\n月度统计:\n")
                f.write("=" * 50 + "\n")

                # 按月份分组
                monthly_records = defaultdict(list)
                for record in records:
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
            return True
        except Exception as e:
            print(f"导出到文本文件时出错: {e}")
            return False

    def delete_record(self, index):
        """删除指定记录"""
        records = self.data.get("records", [])
        if 0 <= index < len(records):
            deleted = records.pop(index)
            if self.save_data():
                print(f"✓ 已删除记录: {deleted['date']} - {deleted['description']}")
                return True
            else:
                print("错误: 保存数据失败")
                return False
        elif index < 0 and abs(index) <= len(records):
            # 支持负索引
            deleted = records.pop(index)
            if self.save_data():
                print(f"✓ 已删除记录: {deleted['date']} - {deleted['description']}")
                return True
            else:
                print("错误: 保存数据失败")
                return False
        else:
            print("错误: 无效的记录索引")
            return False

    def analyze_spending_patterns(self, months=6):
        """分析消费模式"""
        if not MATPLOTLIB_AVAILABLE or not PANDAS_AVAILABLE:
            print("请先安装pandas和matplotlib库: pip install pandas matplotlib")
            return

        records = self.data.get("records", [])
        if len(records) < 5:
            print("数据不足，无法进行模式分析")
            return

        try:
            # 转换为DataFrame
            df = pd.DataFrame(records)
            df['date'] = pd.to_datetime(df['date'])
            df['month'] = df['date'].dt.to_period('M')

            # 筛选最近几个月的数据
            recent_months = df['month'].unique()[-months:]
            df_recent = df[df['month'].isin(recent_months)]

            # 支出分析
            expenses = df_recent[df_recent['type'] == '支出']

            if len(expenses) == 0:
                print("没有找到支出数据")
                return

            # 创建图表
            fig, axes = plt.subplots(2, 2, figsize=(12, 10))
            fig.suptitle('消费模式分析', fontsize=16)

            # 1. 月度支出趋势
            monthly_expense = expenses.groupby('month').agg({'amount': 'sum'}).reset_index()
            monthly_expense['month'] = monthly_expense['month'].astype(str)

            axes[0, 0].bar(monthly_expense['month'], monthly_expense['amount'])
            axes[0, 0].set_title('月度支出趋势')
            axes[0, 0].set_xlabel('月份')
            axes[0, 0].set_ylabel('金额')
            axes[0, 0].tick_params(axis='x', rotation=45)

            # 2. 支出类别分布
            category_expense = expenses.groupby('category').agg({'amount': 'sum'}).reset_index()
            axes[0, 1].pie(category_expense['amount'], labels=category_expense['category'], autopct='%1.1f%%')
            axes[0, 1].set_title('支出类别分布')

            # 3. 月度平均每日支出
            expenses['day_of_week'] = expenses['date'].dt.dayofweek
            daily_avg = expenses.groupby('day_of_week').agg({'amount': 'mean'}).reset_index()
            days = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']
            daily_avg['day_name'] = [days[int(i)] for i in daily_avg['day_of_week']]

            axes[1, 0].plot(daily_avg['day_name'], daily_avg['amount'], marker='o')
            axes[1, 0].set_title('平均每日支出（按星期）')
            axes[1, 0].set_xlabel('星期')
            axes[1, 0].set_ylabel('平均金额')

            # 4. 月度支出统计
            monthly_stats = df_recent[df_recent['type'] == '支出'].groupby('month').agg({
                'amount': ['sum', 'mean', 'count']
            }).reset_index()

            axes[1, 1].bar(range(len(monthly_stats)), monthly_stats[('amount', 'sum')], alpha=0.5, label='总额')
            axes[1, 1].plot(range(len(monthly_stats)), monthly_stats[('amount', 'mean')], 'r-', marker='o',
                            label='平均值')
            axes[1, 1].set_title('月度支出统计')
            axes[1, 1].set_xlabel('月份')
            axes[1, 1].set_ylabel('金额')
            axes[1, 1].legend()
            axes[1, 1].set_xticks(range(len(monthly_stats)))
            axes[1, 1].set_xticklabels([str(m) for m in monthly_stats['month']], rotation=45)

            plt.tight_layout()
            plt.show()

            # 关闭图形以释放内存
            plt.close()

            # 打印分析结果
            print("\n📈 消费模式分析结果:")
            print(f"分析周期: 最近{months}个月")
            print(f"总支出记录数: {len(expenses)}")
            print(f"平均每月支出: ¥{expenses['amount'].mean():.2f}")
            print(f"最大单笔支出: ¥{expenses['amount'].max():.2f}")

            if not expenses['category'].mode().empty:
                print(f"最常见支出类别: {expenses['category'].mode().iloc[0]}")
            else:
                print("最常见支出类别: 无")

        except Exception as e:
            print(f"分析消费模式时出错: {e}")

    def set_budget(self, category, monthly_budget, year_month=None):
        """设置月度预算"""
        if year_month is None:
            year_month = self.get_current_month()

        if monthly_budget <= 0:
            print("错误: 预算金额必须大于0")
            return False

        if "budgets" not in self.data:
            self.data["budgets"] = {}

        if year_month not in self.data["budgets"]:
            self.data["budgets"][year_month] = {}

        self.data["budgets"][year_month][category] = monthly_budget
        if self.save_data():
            print(f"✓ 已为{year_month}的{category}类别设置预算: ¥{monthly_budget:.2f}")
            return True
        return False

    def check_budget(self, year_month=None):
        """检查预算使用情况"""
        if year_month is None:
            year_month = self.get_current_month()

        if "budgets" not in self.data or year_month not in self.data["budgets"]:
            print(f"{year_month}没有设置预算")
            return

        budgets = self.data["budgets"][year_month]
        records = [r for r in self.data.get("records", [])
                   if r["date"].startswith(year_month) and r["type"] == "支出"]

        if not records:
            print(f"{year_month}没有支出记录")
            return

        # 按类别统计支出
        expenses_by_category = defaultdict(float)
        for record in records:
            if record["type"] == "支出":
                expenses_by_category[record["category"]] += record["amount"]

        print(f"\n💰 {year_month}预算检查")
        print("=" * 60)
        print(f"{'类别':<15} {'预算':<12} {'实际支出':<12} {'剩余':<12} {'使用率':<10}")
        print("-" * 60)

        total_budget = 0
        total_spent = 0

        for category, budget in budgets.items():
            spent = expenses_by_category.get(category, 0)
            remaining = budget - spent
            usage_rate = (spent / budget * 100) if budget > 0 else 0

            # 使用率颜色指示
            if usage_rate < 70:
                indicator = "🟢"
            elif usage_rate < 90:
                indicator = "🟡"
            else:
                indicator = "🔴"

            print(f"{category:<15} ¥{budget:<10.2f} ¥{spent:<10.2f} "
                  f"¥{remaining:<10.2f} {usage_rate:<8.1f}% {indicator}")

            total_budget += budget
            total_spent += spent

        print("-" * 60)
        total_remaining = total_budget - total_spent
        total_usage = (total_spent / total_budget * 100) if total_budget > 0 else 0

        print(f"{'总计':<15} ¥{total_budget:<10.2f} ¥{total_spent:<10.2f} "
              f"¥{total_remaining:<10.2f} {total_usage:<8.1f}%")

        if total_usage > 100:
            print(f"⚠️  警告: 总支出已超出预算 ¥{-total_remaining:.2f}")

    def search_records(self, keyword=None, category=None,
                       min_amount=None, max_amount=None,
                       start_date=None, end_date=None):
        """高级搜索功能"""
        results = self.data.get("records", [])

        # 应用筛选条件
        if keyword:
            results = [r for r in results if keyword.lower() in r["description"].lower()]

        if category:
            results = [r for r in results if r["category"] == category]

        if min_amount is not None:
            results = [r for r in results if r["amount"] >= min_amount]

        if max_amount is not None:
            results = [r for r in results if r["amount"] <= max_amount]

        if start_date:
            results = [r for r in results if r["date"] >= start_date]

        if end_date:
            results = [r for r in results if r["date"] <= end_date]

        if not results:
            print("没有找到匹配的记录")
            return []

        # 显示结果
        print(f"\n🔍 搜索到 {len(results)} 条记录")
        print("=" * 60)
        print(f"{'日期':<12} {'类型':<6} {'金额':<12} {'描述':<20} {'类别':<10}")
        print("-" * 60)

        # 显示所有结果，而不仅仅是最后50条
        for record in results[:50]:  # 最多显示50条
            print(f"{record['date']:<12} {record['type']:<6} ¥{record['amount']:<10.2f} "
                  f"{record['description'][:18]:<20} {record['category'][:8]:<10}")

        if len(results) > 50:
            print(f"... 还有 {len(results) - 50} 条记录未显示")

        print("=" * 60)

        # 统计搜索结果
        income_total = sum(r["amount"] for r in results if r["type"] == "收入")
        expense_total = sum(r["amount"] for r in results if r["type"] == "支出")
        net_balance = income_total - expense_total

        print("搜索结果统计:")
        print(f"收入总计: ¥{income_total:.2f}")
        print(f"支出总计: ¥{expense_total:.2f}")
        print(f"净额: ¥{net_balance:.2f}")

        return results

    def predict_future_expenses(self, months=3):
        """简单预测未来支出"""
        if not PANDAS_AVAILABLE:
            print("请先安装pandas库: pip install pandas")
            return

        records = self.data.get("records", [])
        if len(records) < 3:
            print("数据不足，无法进行预测")
            return

        # 提取历史支出数据
        expenses = [r for r in records if r["type"] == "支出"]
        if len(expenses) < 3:
            print("支出数据不足")
            return

        try:
            # 按月份统计
            df = pd.DataFrame(expenses)
            df['date'] = pd.to_datetime(df['date'])
            df['month'] = df['date'].dt.to_period('M')

            monthly_expenses = df.groupby('month').agg({'amount': 'sum'}).reset_index()
            monthly_expenses['month'] = monthly_expenses['month'].astype(str)

            if len(monthly_expenses) < 3:
                print("月度数据不足")
                return

            # 简单移动平均预测
            recent_months = monthly_expenses['amount'].tail(3).tolist()
            avg_expense = sum(recent_months) / len(recent_months)

            # 计算增长率
            growth_rates = []
            for i in range(1, len(monthly_expenses)):
                prev_amount = monthly_expenses.iloc[i - 1]['amount']
                curr_amount = monthly_expenses.iloc[i]['amount']
                if prev_amount > 0:
                    growth = (curr_amount - prev_amount) / prev_amount
                    growth_rates.append(growth)

            avg_growth = sum(growth_rates) / len(growth_rates) if growth_rates else 0

            print(f"\n🔮 未来{months}个月支出预测")
            print("=" * 50)
            print(f"基于最近{len(recent_months)}个月的数据:")
            for i, amount in enumerate(recent_months, 1):
                print(f"  前{len(recent_months) - i + 1}个月: ¥{amount:.2f}")
            print(f"  平均月支出: ¥{avg_expense:.2f}")
            if growth_rates:
                print(f"  平均增长率: {avg_growth * 100:.1f}%")
            else:
                print(f"  平均增长率: 0.0% (无历史增长数据)")

            # 预测未来月份
            current_month = pd.Period(self.get_current_month(), freq='M')
            last_amount = recent_months[-1]

            print(f"\n预测结果:")
            predicted_amounts = []
            for i in range(1, months + 1):
                next_month = current_month + i
                # 使用加权平均预测
                predicted = last_amount * (1 + avg_growth * 0.5)  # 减缓增长率
                predicted_amounts.append(predicted)
                print(f"  {next_month}: ¥{predicted:.2f}")
                last_amount = predicted  # 修复：更新last_amount为预测值

            # 提供建议
            print(f"\n💡 建议:")
            if avg_growth > 0.1:  # 增长率超过10%
                print("  ⚠️  支出增长较快，建议审视消费习惯")
            elif avg_growth < -0.1:  # 负增长超过10%
                print("  ✅ 支出呈下降趋势，继续保持")
            else:
                print("  📊 支出相对稳定")

        except Exception as e:
            print(f"预测未来支出时出错: {e}")

    def export_to_excel(self, filename=None):
        """导出到Excel文件"""
        if not PANDAS_AVAILABLE or not OPENPYXL_AVAILABLE:
            print("请先安装pandas和openpyxl库: pip install pandas openpyxl")
            return False

        if filename is None:
            filename = f"finance_export_{self.get_current_date()}.xlsx"

        try:
            # 创建DataFrame
            records = self.data.get("records", [])
            if not records:
                print("没有数据可导出")
                return False

            df = pd.DataFrame(records)

            # 添加年月列用于分组
            df['date'] = pd.to_datetime(df['date'])
            df['year'] = df['date'].dt.year
            df['month'] = df['date'].dt.month

            # 创建Excel写入器
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                # 原始数据
                df.to_excel(writer, sheet_name='所有记录', index=False)

                # 月度统计
                if len(df) > 0:
                    monthly_summary = df.groupby(['year', 'month', 'type']).agg({
                        'amount': ['sum', 'count', 'mean', 'max', 'min']
                    }).round(2)
                    if not monthly_summary.empty:
                        monthly_summary.to_excel(writer, sheet_name='月度统计')

                # 类别统计
                category_summary = df.groupby(['type', 'category']).agg({
                    'amount': ['sum', 'count']
                }).round(2)
                if not category_summary.empty:
                    category_summary.to_excel(writer, sheet_name='类别统计')

                # 年度总结
                if len(df) > 0:
                    year_summary = df.groupby(['year', 'type']).agg({
                        'amount': 'sum'
                    }).unstack(fill_value=0).round(2)
                    if not year_summary.empty:
                        year_summary.columns = year_summary.columns.droplevel(0)
                        year_summary['结余'] = year_summary.get('收入', 0) - year_summary.get('支出', 0)
                        year_summary.to_excel(writer, sheet_name='年度总结')

            print(f"✓ 数据已导出到 {filename}")
            print("  包含以下工作表:")
            print("  - 所有记录: 所有原始交易记录")
            print("  - 月度统计: 按月度和类型的统计")
            print("  - 类别统计: 按类别的统计")
            print("  - 年度总结: 年度收支总结")
            return True

        except Exception as e:
            print(f"导出到Excel时出错: {e}")
            return False

    def auto_backup(self, backup_dir="backups"):
        """自动备份数据"""
        import shutil

        # 检查源文件是否存在
        if not os.path.exists(self.filename):
            print(f"错误: 源文件 {self.filename} 不存在")
            return None

        try:
            # 创建备份目录
            os.makedirs(backup_dir, exist_ok=True)

            # 生成备份文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_file = os.path.join(backup_dir, f"finance_backup_{timestamp}.json")

            # 备份数据
            shutil.copy2(self.filename, backup_file)

            # 清理旧备份（保留最近7个）
            try:
                backup_files = []
                for f in os.listdir(backup_dir):
                    if f.endswith('.json') and f.startswith('finance_backup_'):
                        filepath = os.path.join(backup_dir, f)
                        if os.path.isfile(filepath):
                            backup_files.append((filepath, os.path.getmtime(filepath)))

                # 按修改时间排序
                backup_files.sort(key=lambda x: x[1])

                # 删除除最近7个外的所有备份
                if len(backup_files) > 7:
                    for filepath, _ in backup_files[:-7]:
                        try:
                            os.remove(filepath)
                        except:
                            pass
            except Exception as e:
                print(f"清理旧备份时出错: {e}")

            print(f"✓ 数据已备份到: {backup_file}")
            return backup_file

        except Exception as e:
            print(f"自动备份时出错: {e}")
            return None

    def set_reminder(self, reminder_type, time, category=None, amount=None):
        """设置提醒"""
        if reminder_type not in ["daily", "weekly", "monthly"]:
            print("提醒类型必须是 daily, weekly, 或 monthly")
            return False

        # 验证时间格式
        try:
            if ":" in time:
                datetime.strptime(time, "%H:%M")
        except ValueError:
            print("错误: 时间格式应为 HH:MM")
            return False

        reminder = {
            "type": reminder_type,
            "time": time,
            "category": category,
            "amount": amount,
            "enabled": True
        }

        self.data["reminders"][reminder_type] = reminder
        if self.save_data():
            print(f"✓ 已设置{reminder_type}提醒，时间: {time}")
            return True
        return False

    def check_reminders(self):
        """检查提醒"""
        reminders = self.data.get("reminders", {})
        if not reminders:
            return

        today = datetime.now().strftime("%Y-%m-%d")
        today_records = [r for r in self.data.get("records", []) if r.get("date") == today]

        print("\n🔔 今日提醒:")
        print("=" * 50)

        today_income = sum(r["amount"] for r in today_records if r["type"] == "收入")
        today_expense = sum(r["amount"] for r in today_records if r["type"] == "支出")

        print(f"今日已记录: {len(today_records)} 笔交易")
        print(f"今日收入: ¥{today_income:.2f}")
        print(f"今日支出: ¥{today_expense:.2f}")

        if len(today_records) == 0:
            print("💡 提醒: 今天还没有记录任何交易，记得记账哦！")

        # 检查预算提醒
        current_month = self.get_current_month()
        if "budgets" in self.data and current_month in self.data["budgets"]:
            budgets = self.data["budgets"][current_month]
            month_records = [r for r in self.data.get("records", [])
                             if r["date"].startswith(current_month) and r["type"] == "支出"]

            for category, budget in budgets.items():
                spent = sum(r["amount"] for r in month_records if r.get("category") == category)
                usage_rate = (spent / budget * 100) if budget > 0 else 0

                if usage_rate > 80:
                    print(f"⚠️  {category}类别预算已使用 {usage_rate:.1f}%")

    def generate_report(self, year_month=None, output_format="text"):
        """生成详细报告"""
        if year_month is None:
            year_month = self.get_current_month()

        records = [r for r in self.data.get("records", []) if r["date"].startswith(year_month)]

        if not records:
            print(f"{year_month}没有数据")
            return None

        if output_format == "html":
            return self._generate_html_report(year_month, records)
        else:
            return self._generate_text_report(year_month, records)

    def _generate_text_report(self, year_month, records):
        """生成文本报告"""
        filename = f"finance_report_{year_month}.txt"

        try:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(f"个人财务报告 - {year_month}\n")
                f.write("=" * 60 + "\n\n")

                # 总体统计
                income_total = sum(r["amount"] for r in records if r["type"] == "收入")
                expense_total = sum(r["amount"] for r in records if r["type"] == "支出")
                balance = income_total - expense_total

                f.write("📊 总体统计\n")
                f.write("-" * 40 + "\n")
                f.write(f"总收入: ¥{income_total:.2f}\n")
                f.write(f"总支出: ¥{expense_total:.2f}\n")
                f.write(f"收支结余: ¥{balance:.2f}\n\n")

                # 收入详情
                income_records = [r for r in records if r["type"] == "收入"]
                if income_records:
                    f.write("💰 收入详情\n")
                    f.write("-" * 40 + "\n")
                    income_by_category = defaultdict(float)
                    for record in income_records:
                        income_by_category[record["category"]] += record["amount"]
                        f.write(
                            f"{record['date']} | {record['description']:<20} | ¥{record['amount']:>10.2f} | {record['category']}\n")

                    f.write(f"\n收入类别统计:\n")
                    for category, amount in sorted(income_by_category.items(), key=lambda x: x[1], reverse=True):
                        f.write(f"  {category}: ¥{amount:.2f}\n")
                    f.write("\n")

                # 支出详情
                expense_records = [r for r in records if r["type"] == "支出"]
                if expense_records:
                    f.write("💸 支出详情\n")
                    f.write("-" * 40 + "\n")
                    expense_by_category = defaultdict(float)
                    for record in expense_records:
                        expense_by_category[record["category"]] += record["amount"]
                        f.write(
                            f"{record['date']} | {record['description']:<20} | ¥{record['amount']:>10.2f} | {record['category']}\n")

                    f.write(f"\n支出类别统计:\n")
                    for category, amount in sorted(expense_by_category.items(), key=lambda x: x[1], reverse=True):
                        percentage = (amount / expense_total * 100) if expense_total > 0 else 0
                        f.write(f"  {category}: ¥{amount:.2f} ({percentage:.1f}%)\n")
                    f.write("\n")

                # 预算检查
                if "budgets" in self.data and year_month in self.data["budgets"]:
                    f.write("📋 预算执行情况\n")
                    f.write("-" * 40 + "\n")
                    budgets = self.data["budgets"][year_month]
                    for category, budget in budgets.items():
                        spent = expense_by_category.get(category, 0)
                        remaining = budget - spent
                        usage_rate = (spent / budget * 100) if budget > 0 else 0
                        status = "✅ 正常" if usage_rate <= 100 else "❌ 超支"
                        f.write(f"{category}: 预算 ¥{budget:.2f}, 实际 ¥{spent:.2f}, "
                                f"剩余 ¥{remaining:.2f}, 使用率 {usage_rate:.1f}% {status}\n")

            print(f"✓ 报告已生成: {filename}")
            return filename

        except Exception as e:
            print(f"生成文本报告时出错: {e}")
            return None

    def _generate_html_report(self, year_month, records):
        """生成HTML报告（基础版）"""
        if not JINJA2_AVAILABLE:
            print("请先安装Jinja2库: pip install Jinja2")
            return self._generate_text_report(year_month, records)

        try:
            # 统计计算
            income_total = sum(r["amount"] for r in records if r["type"] == "收入")
            expense_total = sum(r["amount"] for r in records if r["type"] == "支出")
            balance = income_total - expense_total

            income_by_category = defaultdict(float)
            expense_by_category = defaultdict(float)

            for record in records:
                if record["type"] == "收入":
                    income_by_category[record["category"]] += record["amount"]
                else:
                    expense_by_category[record["category"]] += record["amount"]

            # HTML模板
            html_template = """
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="UTF-8">
                <title>财务报告 - {{ year_month }}</title>
                <style>
                    body { font-family: Arial, sans-serif; margin: 40px; }
                    .header { background-color: #f0f0f0; padding: 20px; border-radius: 5px; }
                    .summary { margin: 20px 0; padding: 20px; border: 1px solid #ddd; border-radius: 5px; }
                    .positive { color: green; font-weight: bold; }
                    .negative { color: red; font-weight: bold; }
                    table { width: 100%; border-collapse: collapse; margin: 20px 0; }
                    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
                    th { background-color: #f2f2f2; }
                    tr:nth-child(even) { background-color: #f9f9f9; }
                    .budget-ok { background-color: #d4edda; }
                    .budget-warn { background-color: #fff3cd; }
                    .budget-danger { background-color: #f8d7da; }
                </style>
            </head>
            <body>
                <div class="header">
                    <h1>个人财务报告 - {{ year_month }}</h1>
                    <p>生成时间: {{ generated_time }}</p>
                </div>

                <div class="summary">
                    <h2>总体统计</h2>
                    <p>总收入: ¥{{ "%.2f"|format(income_total) }}</p>
                    <p>总支出: ¥{{ "%.2f"|format(expense_total) }}</p>
                    <p>收支结余: 
                        <span class="{{ 'positive' if balance >= 0 else 'negative' }}">
                            ¥{{ "%.2f"|format(balance) }}
                        </span>
                    </p>
                </div>

                {% if income_by_category %}
                <div>
                    <h2>收入类别统计</h2>
                    <table>
                        <tr><th>类别</th><th>金额</th><th>占比</th></tr>
                        {% for category, amount in income_by_category.items()|sort(attribute='1', reverse=true) %}
                        <tr>
                            <td>{{ category }}</td>
                            <td>¥{{ "%.2f"|format(amount) }}</td>
                            <td>{{ "%.1f"|format(amount/income_total*100) if income_total > 0 else 0 }}%</td>
                        </tr>
                        {% endfor %}
                    </table>
                </div>
                {% endif %}

                {% if expense_by_category %}
                <div>
                    <h2>支出类别统计</h2>
                    <table>
                        <tr><th>类别</th><th>金额</th><th>占比</th></tr>
                        {% for category, amount in expense_by_category.items()|sort(attribute='1', reverse=true) %}
                        <tr>
                            <td>{{ category }}</td>
                            <td>¥{{ "%.2f"|format(amount) }}</td>
                            <td>{{ "%.1f"|format(amount/expense_total*100) if expense_total > 0 else 0 }}%</td>
                        </tr>
                        {% endfor %}
                    </table>
                </div>
                {% endif %}
            </body>
            </html>
            """

            # 渲染模板
            template = Template(html_template)
            html_content = template.render(
                year_month=year_month,
                generated_time=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                income_total=income_total,
                expense_total=expense_total,
                balance=balance,
                income_by_category=dict(income_by_category),
                expense_by_category=dict(expense_by_category)
            )

            # 保存文件
            filename = f"finance_report_{year_month}.html"
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(html_content)

            print(f"✓ HTML报告已生成: {filename}")
            return filename

        except Exception as e:
            print(f"生成HTML报告时出错: {e}")
            return None


def main():
    """主程序"""
    # 检查依赖库
    print("正在启动记账程序 - 增强版...")
    print("数据将保存在 finance_data.json 文件中")

    # 列出可选的库状态
    optional_libs = []
    if not PANDAS_AVAILABLE:
        optional_libs.append("pandas")
    if not MATPLOTLIB_AVAILABLE:
        optional_libs.append("matplotlib")
    if not OPENPYXL_AVAILABLE:
        optional_libs.append("openpyxl")
    if not JINJA2_AVAILABLE:
        optional_libs.append("Jinja2")

    if optional_libs:
        print("=" * 50)
        print("注意: 以下可选功能需要安装额外的库:")
        for lib in optional_libs:
            print(f"  - {lib}")
        print("您可以通过以下命令安装: pip install " + " ".join(optional_libs))
        print("部分高级功能将不可用，但基本功能正常。")

    print("=" * 50)

    tracker = SimpleFinanceTracker()

    while True:
        print("\n" + "=" * 50)
        print("个人记账系统 - 增强版")
        print("=" * 50)
        print("基本功能:")
        print(" 1. 添加收入记录")
        print(" 2. 添加支出记录")
        print(" 3. 查看所有记录")
        print(" 4. 查看本月记录")
        print(" 5. 查看本月统计")
        print(" 6. 查看年度统计")
        print(" 7. 导出数据到文本文件")
        print(" 8. 查看类别列表")
        print(" 9. 删除最新记录")

        print("\n高级功能:")
        print("10. 消费模式分析" + (
            " (需要pandas, matplotlib)" if not (PANDAS_AVAILABLE and MATPLOTLIB_AVAILABLE) else ""))
        print("11. 设置预算")
        print("12. 检查预算")
        print("13. 高级搜索")
        print("14. 未来支出预测" + (" (需要pandas)" if not PANDAS_AVAILABLE else ""))
        print("15. 导出到Excel" + (" (需要pandas, openpyxl)" if not (PANDAS_AVAILABLE and OPENPYXL_AVAILABLE) else ""))
        print("16. 自动备份数据")
        print("17. 设置提醒")
        print("18. 检查提醒")
        print("19. 生成报告")
        print(" 0. 退出程序")
        print("-" * 50)

        try:
            choice = input("请选择操作 (0-19): ").strip()

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
                records = tracker.data.get("records", [])
                if records:
                    print(f"确定要删除最新记录吗?")
                    print(f"记录: {records[-1]}")
                    confirm = input("输入 'y' 确认删除: ").lower()
                    if confirm == 'y':
                        tracker.delete_record(-1)
                else:
                    print("暂无记录可删除")

            elif choice == "10":
                if PANDAS_AVAILABLE and MATPLOTLIB_AVAILABLE:
                    try:
                        months = input("分析最近几个月的数据? (默认6): ").strip()
                        months = int(months) if months.isdigit() else 6
                        tracker.analyze_spending_patterns(months)
                    except Exception as e:
                        print(f"分析消费模式时出错: {e}")
                else:
                    print("此功能需要pandas和matplotlib库，请先安装")

            elif choice == "11":
                print("\n设置月度预算")
                print("-" * 30)
                category = input("类别名称: ").strip()
                try:
                    budget = float(input("月度预算金额: "))
                    year_month = input("年月 (YYYY-MM, 直接回车使用当前月): ").strip()
                    if not year_month:
                        year_month = None
                    tracker.set_budget(category, budget, year_month)
                except ValueError:
                    print("错误: 预算金额必须为数字")

            elif choice == "12":
                year_month = input("检查哪个月的预算? (YYYY-MM, 直接回车使用当前月): ").strip()
                if not year_month:
                    year_month = None
                tracker.check_budget(year_month)

            elif choice == "13":
                print("\n高级搜索")
                print("-" * 30)
                print("提示: 直接按回车跳过筛选条件")
                keyword = input("关键词搜索: ").strip() or None
                category = input("类别筛选: ").strip() or None

                min_amount_input = input("最小金额: ").strip()
                min_amount = float(min_amount_input) if min_amount_input else None

                max_amount_input = input("最大金额: ").strip()
                max_amount = float(max_amount_input) if max_amount_input else None

                start_date = input("开始日期 (YYYY-MM-DD): ").strip() or None
                end_date = input("结束日期 (YYYY-MM-DD): ").strip() or None

                tracker.search_records(
                    keyword=keyword,
                    category=category,
                    min_amount=min_amount,
                    max_amount=max_amount,
                    start_date=start_date,
                    end_date=end_date
                )

            elif choice == "14":
                if PANDAS_AVAILABLE:
                    try:
                        months = input("预测未来几个月? (默认3): ").strip()
                        months = int(months) if months.isdigit() else 3
                        tracker.predict_future_expenses(months)
                    except Exception as e:
                        print(f"预测未来支出时出错: {e}")
                else:
                    print("此功能需要pandas库，请先安装")

            elif choice == "15":
                if PANDAS_AVAILABLE and OPENPYXL_AVAILABLE:
                    filename = input("导出文件名 (默认: finance_export_YYYY-MM-DD.xlsx): ").strip()
                    if not filename:
                        filename = None
                    tracker.export_to_excel(filename)
                else:
                    print("此功能需要pandas和openpyxl库，请先安装")

            elif choice == "16":
                backup_dir = input("备份目录 (默认: backups): ").strip()
                if not backup_dir:
                    backup_dir = "backups"
                tracker.auto_backup(backup_dir)

            elif choice == "17":
                print("\n设置提醒")
                print("-" * 30)
                reminder_type = input("提醒类型 (daily/weekly/monthly): ").strip().lower()
                time = input("提醒时间 (HH:MM): ").strip()
                category = input("类别 (可选, 直接回车跳过): ").strip() or None

                amount_input = input("金额阈值 (可选, 直接回车跳过): ").strip()
                amount = float(amount_input) if amount_input else None

                tracker.set_reminder(reminder_type, time, category, amount)

            elif choice == "18":
                tracker.check_reminders()

            elif choice == "19":
                print("\n生成报告")
                print("-" * 30)
                year_month = input("年月 (YYYY-MM, 直接回车使用当前月): ").strip()
                if not year_month:
                    year_month = None

                format_choice = input("报告格式 (text/html, 默认text): ").strip().lower()
                if format_choice not in ["text", "html"]:
                    format_choice = "text"

                tracker.generate_report(year_month, format_choice)

            else:
                print("无效选择，请重新输入")

        except KeyboardInterrupt:
            print("\n\n程序被用户中断")
            break
        except Exception as e:
            print(f"发生错误: {e}")


if __name__ == "__main__":
    main()