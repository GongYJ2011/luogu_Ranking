import time
import requests
import pandas as pd
import json
import threading
import os
from tkinter import *
from tkinter import ttk, messagebox, filedialog
from pandastable import Table

# 提前设定输出格式
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('max_colwidth', 50)


class LuoguRankingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("洛谷比赛排行榜获取工具")
        self.root.geometry("800x650")
        self.root.resizable(True, True)

        # 创建主框架
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=BOTH, expand=True)

        # 创建输入区域
        self.create_input_area()

        # 创建按钮区域
        self.create_button_area()

        # 创建结果区域
        self.create_result_area()

        # 创建状态栏
        self.status_var = StringVar()
        self.status_bar = ttk.Label(root, textvariable=self.status_var, relief=SUNKEN, anchor=W)
        self.status_bar.pack(side=BOTTOM, fill=X)
        self.update_status("就绪")

        # 初始化变量
        self.df_all = pd.DataFrame()
        self.df_user = pd.DataFrame()
        self.stop_requested = False  # 添加停止标志

    def create_input_area(self):
        """创建输入区域"""
        input_frame = ttk.LabelFrame(self.main_frame, text="参数设置", padding="10")
        input_frame.pack(fill=X, padx=5, pady=5)

        # 比赛编号
        ttk.Label(input_frame, text="比赛编号:").grid(row=0, column=0, sticky=W, padx=5, pady=5)
        self.contest_entry = ttk.Entry(input_frame, width=20)
        self.contest_entry.grid(row=0, column=1, sticky=W, padx=5, pady=5)

        # 用户名
        ttk.Label(input_frame, text="查找用户名:").grid(row=1, column=0, sticky=W, padx=5, pady=5)
        self.username_entry = ttk.Entry(input_frame, width=20)
        self.username_entry.grid(row=1, column=1, sticky=W, padx=5, pady=5)

        # 起始页码
        ttk.Label(input_frame, text="起始页码:").grid(row=2, column=0, sticky=W, padx=5, pady=5)
        self.page_start_entry = ttk.Entry(input_frame, width=5)
        self.page_start_entry.insert(0, "1")
        self.page_start_entry.grid(row=2, column=1, sticky=W, padx=5, pady=5)

        # 结束页码
        ttk.Label(input_frame, text="结束页码:").grid(row=2, column=2, sticky=W, padx=5, pady=5)
        self.page_end_entry = ttk.Entry(input_frame, width=5)
        self.page_end_entry.grid(row=2, column=3, sticky=W, padx=5, pady=5)

        # 添加示例文本
        example_frame = ttk.Frame(input_frame)
        example_frame.grid(row=4, column=0, columnspan=3, sticky=W, padx=5, pady=5)
        ttk.Label(example_frame, text="示例: 比赛编号:123456, 用户名:chen_zhe, 起始页:1, 结束页:5").pack(anchor=W)

    def create_button_area(self):
        """创建按钮区域"""
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(fill=X, padx=5, pady=5)

        # 获取数据按钮
        self.fetch_button = ttk.Button(button_frame, text="获取排行榜", command=self.start_fetching)
        self.fetch_button.pack(side=LEFT, padx=5)

        # 停止按钮
        self.stop_button = ttk.Button(button_frame, text="停止获取", command=self.stop_fetching, state=DISABLED)
        self.stop_button.pack(side=LEFT, padx=5)

        # 保存数据按钮
        self.save_button = ttk.Button(button_frame, text="保存结果", command=self.save_results, state=DISABLED)
        self.save_button.pack(side=LEFT, padx=5)

        # 清除结果按钮
        self.clear_button = ttk.Button(button_frame, text="清除结果", command=self.clear_results)
        self.clear_button.pack(side=LEFT, padx=5)

        # 退出按钮
        self.quit_button = ttk.Button(button_frame, text="退出", command=self.root.quit)
        self.quit_button.pack(side=RIGHT, padx=5)

    def create_result_area(self):
        """创建结果显示区域"""
        # 创建Notebook用于切换不同的结果视图
        self.result_notebook = ttk.Notebook(self.main_frame)
        self.result_notebook.pack(fill=BOTH, expand=True, padx=5, pady=5)

        # 完整排行榜标签页
        self.full_ranking_frame = ttk.Frame(self.result_notebook)
        self.result_notebook.add(self.full_ranking_frame, text="完整排行榜")

        # 用户排名标签页
        self.user_ranking_frame = ttk.Frame(self.result_notebook)
        self.result_notebook.add(self.user_ranking_frame, text="用户排名")

        # 日志标签页
        self.log_frame = ttk.Frame(self.result_notebook)
        self.result_notebook.add(self.log_frame, text="日志")

        # 添加日志文本框
        self.log_text = Text(self.log_frame, wrap=WORD)
        self.log_scroll = ttk.Scrollbar(self.log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=self.log_scroll.set)

        self.log_scroll.pack(side=RIGHT, fill=Y)
        self.log_text.pack(fill=BOTH, expand=True)

        # 禁用日志编辑
        self.log_text.config(state=DISABLED)

    def update_status(self, message):
        """更新状态栏"""
        self.status_var.set(message)
        self.root.update()

    def log_message(self, message):
        """向日志框添加消息"""
        self.log_text.config(state=NORMAL)
        self.log_text.insert(END, message + "\n")
        self.log_text.see(END)  # 滚动到底部
        self.log_text.config(state=DISABLED)
        self.update_status(message)

    def clear_results(self):
        """清除结果"""
        self.df_all = pd.DataFrame()
        self.df_user = pd.DataFrame()
        self.log_text.config(state=NORMAL)
        self.log_text.delete(1.0, END)
        self.log_text.config(state=DISABLED)

        # 清除表格视图
        for widget in self.full_ranking_frame.winfo_children():
            widget.destroy()

        for widget in self.user_ranking_frame.winfo_children():
            widget.destroy()

        self.save_button.config(state=DISABLED)
        self.update_status("结果已清除")

    def stop_fetching(self):
        """停止获取数据"""
        self.stop_requested = True
        self.log_message("正在停止获取数据...")
        self.stop_button.config(state=DISABLED)

    def start_fetching(self):
        """开始获取数据（在新线程中）"""
        # 获取输入参数
        contest_id = self.contest_entry.get().strip()
        username = self.username_entry.get().strip()
        page_start = self.page_start_entry.get().strip()
        page_end = self.page_end_entry.get().strip()

        # 验证输入
        if not contest_id:
            messagebox.showerror("输入错误", "请输入比赛编号")
            return
        if not username:
            messagebox.showerror("输入错误", "请输入要查找的用户名")
            return

        # 验证页码输入
        if not page_start.isdigit() or int(page_start) <= 0:
            messagebox.showerror("输入错误", "起始页码必须为正整数")
            return

        # 验证结束页码
        if page_end and (not page_end.isdigit() or int(page_end) <= 0):
            messagebox.showerror("输入错误", "结束页码必须为正整数")
            return

        # 验证结束页码是否大于起始页码
        if page_end and int(page_end) < int(page_start):
            messagebox.showerror("输入错误", "结束页码不能小于起始页码")
            return

        # 重置停止标志
        self.stop_requested = False

        # 禁用按钮
        self.fetch_button.config(state=DISABLED)
        self.save_button.config(state=DISABLED)
        self.stop_button.config(state=NORMAL)

        # 清空日志
        self.log_text.config(state=NORMAL)
        self.log_text.delete(1.0, END)
        self.log_text.config(state=DISABLED)

        # 在新线程中运行获取过程
        self.log_message(f"开始获取比赛 {contest_id} 的排行榜...")
        self.log_message(f"查找用户: {username}")
        self.log_message(f"起始页码: {page_start}, 结束页码: {page_end if page_end else '无限制'}")

        threading.Thread(
            target=self.fetch_data,
            args=(contest_id, username, int(page_start), "0000000000000000000000000000000000000000"),
            daemon=True
        ).start()

    def fetch_data(self, contest_id, username, start_page, end_page):
        """获取数据的主函数"""
        try:
            page = start_page
            df_all = pd.DataFrame()
            df_user = pd.DataFrame()

            # 创建cookie字典
            cookies = {"__client_id": "0000000000000000000000000000000000000000"}

            # 最大尝试次数
            max_attempts = 3

            # 将结束页码转换为整数（如果提供）
            end_page = int(end_page) if end_page and end_page.isdigit() else 0

            while not self.stop_requested:
                # 检查结束页码限制
                if (end_page != 0) and page > end_page:
                    self.log_message(f"已达到指定的结束页码 {end_page}")
                    break

                self.log_message(f"正在获取第 {page} 页数据...")

                # 尝试多次请求
                data = None
                for attempt in range(max_attempts):
                    data = self.get_data(contest_id, cookies, page)
                    if data:
                        break
                    elif attempt < max_attempts - 1:
                        self.log_message(f"第 {attempt + 1} 次尝试失败，5秒后重试...")
                        time.sleep(5)

                if not data:
                    self.log_message(f"第 {page} 页数据获取失败，跳过该页")
                    page += 1
                    continue

                # 检查API返回的状态码
                if 'code' in data and data['code'] != 200:
                    error_msg = data.get('errorMessage', '未知错误')
                    self.log_message(f"API返回错误: {error_msg} (代码: {data['code']})")
                    break

                # 检查是否有scoreboard数据
                if 'scoreboard' not in data:
                    self.log_message("返回数据中未找到scoreboard信息")
                    self.log_message(f"返回数据: {json.dumps(data, ensure_ascii=False, indent=2)}")
                    break

                # 处理数据
                df_page, found_user = self.process_data(data, page, username)

                # 添加到总数据框
                df_all = pd.concat([df_all, df_page])

                # 如果找到用户
                if not found_user.empty:
                    df_user = found_user

                # 检查是否还有下一页
                # 修复'totalPage'键不存在的问题
                total_pages = data['scoreboard'].get('totalPage')

                if total_pages:
                    if page >= total_pages:
                        self.log_message(f"已到达最后一页（共 {total_pages} 页）")
                        break
                else:
                    # 如果没有totalPage信息，检查当前页是否有数据
                    current_data = data['scoreboard'].get('result', [])
                    if not current_data:
                        self.log_message("当前页无数据，可能是最后一页")
                        break
                    else:
                        # 如果不知道总页数，默认最多获取100页
                        if page - start_page >= 100:
                            self.log_message("已达到最大页数限制（100页）")
                            break

                page += 1
                time.sleep(1)  # 避免请求过快

            # 更新结果
            self.df_all = df_all
            self.df_user = df_user

            # 显示结果
            self.root.after(0, self.show_results)

            # 启用保存按钮
            self.root.after(0, lambda: self.save_button.config(state=NORMAL))

            if not df_all.empty:
                self.log_message("排行榜获取完成！")
                self.log_message(f"共获取 {len(df_all)} 条记录")
            else:
                self.log_message("未找到任何排行榜数据")

            if df_user.empty:
                self.log_message(f"未找到用户 {username} 的成绩")
            else:
                self.log_message(f"已找到用户 {username} 的成绩")

        except Exception as e:
            self.log_message(f"发生错误: {str(e)}")
            import traceback
            self.log_message(traceback.format_exc())
        finally:
            # 重新启用获取按钮
            self.root.after(0, lambda: self.fetch_button.config(state=NORMAL))
            self.root.after(0, lambda: self.stop_button.config(state=DISABLED))

    def get_data(self, contest_id, cookies, page):
        """请求数据"""
        api_url = f'https://www.luogu.com.cn/fe/api/contest/scoreboard/{contest_id}?page={page}'
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
            'Referer': f'https://www.luogu.com.cn/contest/{contest_id}',
            'X-Requested-With': 'XMLHttpRequest'
        }

        try:
            response = requests.get(api_url, headers=headers, cookies=cookies, timeout=15)

            if response.status_code == 200:
                return response.json()
            else:
                self.log_message(f"请求失败，状态码: {response.status_code}")
                # 尝试获取错误信息
                try:
                    error_data = response.json()
                    self.log_message(f"错误信息: {error_data.get('errorMessage', '未知错误')}")
                except:
                    self.log_message(f"响应内容: {response.text[:200]}")
                return None
        except Exception as e:
            self.log_message(f"请求异常: {str(e)}")
            return None

    def process_data(self, data, page, username):
        """处理JSON数据并转换为DataFrame"""
        # 检查是否有结果数据
        if 'result' not in data['scoreboard']:
            self.log_message(f"第 {page} 页无结果数据")
            return pd.DataFrame(), pd.DataFrame()

        results = data['scoreboard']['result']

        # 创建当前页的数据框
        df_page = pd.DataFrame([
            {
                '排名': idx + 1 + (page - 1) * 50,
                '用户名': result['user']['name'],
                '总分': result['score'],
                **{f'题目 {problem}': details['score'] for problem, details in result['details'].items()}
            }
            for idx, result in enumerate(results)
        ])

        # 查找用户
        df_user = pd.DataFrame()
        for i in range(len(df_page)):
            if df_page.iloc[i]['用户名'] == username:
                df_user = df_page.iloc[[i]]
                break

        return df_page, df_user

    def show_results(self):
        """显示结果"""
        # 显示完整排行榜
        if not self.df_all.empty:
            for widget in self.full_ranking_frame.winfo_children():
                widget.destroy()

            # 创建带滚动条的框架
            frame = ttk.Frame(self.full_ranking_frame)
            frame.pack(fill=BOTH, expand=True)

            # 使用pandastable显示表格
            pt = Table(frame, dataframe=self.df_all, showtoolbar=True, showstatusbar=True)
            pt.show()

        # 显示用户排名
        if not self.df_user.empty:
            for widget in self.user_ranking_frame.winfo_children():
                widget.destroy()

            frame = ttk.Frame(self.user_ranking_frame)
            frame.pack(fill=BOTH, expand=True)

            pt = Table(frame, dataframe=self.df_user, showtoolbar=True, showstatusbar=True)
            pt.show()

    def save_results(self):
        """保存结果到文件"""
        if self.df_all.empty:
            messagebox.showwarning("保存失败", "没有可保存的数据")
            return

        # 获取比赛ID
        contest_id = self.contest_entry.get().strip()
        if not contest_id:
            contest_id = "luogu_contest"

        # 选择保存位置
        filename = filedialog.asksaveasfilename(
            initialfile=f"luogu_contest_{contest_id}",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )

        if not filename:
            return

        try:
            if filename.endswith('.xlsx'):
                # 保存为Excel，包含两个sheet
                with pd.ExcelWriter(filename) as writer:
                    self.df_all.to_excel(writer, sheet_name='完整排行榜', index=False)
                    if not self.df_user.empty:
                        self.df_user.to_excel(writer, sheet_name='用户排名', index=False)
                self.log_message(f"结果已保存到Excel文件: {filename}")

            elif filename.endswith('.csv'):
                # 保存为CSV
                self.df_all.to_csv(filename, index=False)
                self.log_message(f"完整排行榜已保存到CSV文件: {filename}")

                # 保存用户数据到单独文件
                if not self.df_user.empty:
                    user_file = os.path.splitext(filename)[0] + "_用户排名.csv"
                    self.df_user.to_csv(user_file, index=False)
                    self.log_message(f"用户排名已保存到: {user_file}")

            messagebox.showinfo("保存成功", "数据已成功保存！")

        except Exception as e:
            messagebox.showerror("保存失败", f"保存文件时出错: {str(e)}")
            self.log_message(f"保存失败: {str(e)}")


if __name__ == "__main__":
    root = Tk()
    app = LuoguRankingApp(root)
    root.mainloop()
# pyinstaller --onefile --windowed --icon=app.ico --name "洛谷完整比赛排行榜获取 & 查找指定用户排名" __init__.py
