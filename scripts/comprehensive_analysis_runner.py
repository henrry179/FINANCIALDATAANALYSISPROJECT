#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
综合分析运行器
一键运行所有深度分析工具，生成交互式网页和PDF报告
"""

import os
import sys
import time
from datetime import datetime
import subprocess

# ====== 配置区域 ======

EXCEL_FILE = "数据合并结果_20250601_1703.xlsx"
ANALYSIS_TOOLS = [
    ("advanced_data_analyzer.py", "高级数据深度分析"),
    ("interactive_visualization.py", "交互式可视化网页"),
    ("pdf_report_generator.py", "专业PDF分析报告")
]

# ====== 配置区域结束 ======


class ComprehensiveAnalysisRunner:
    """综合分析运行器"""
    
    def __init__(self):
        self.start_time = datetime.now()
        self.results = {}
        
    def check_dependencies(self):
        """检查依赖库"""
        print("🔍 检查分析环境...")
        
        required_packages = [
            ('pandas', 'pip install pandas'),
            ('numpy', 'pip install numpy'),
            ('scipy', 'pip install scipy'),
            ('sklearn', 'pip install scikit-learn'),
            ('plotly', 'pip install plotly'),
            ('matplotlib', 'pip install matplotlib'),
            ('seaborn', 'pip install seaborn'),
            ('reportlab', 'pip install reportlab')
        ]
        
        missing_packages = []
        
        for package, install_cmd in required_packages:
            try:
                __import__(package)
                print(f"   ✅ {package}")
            except ImportError:
                print(f"   ❌ {package} - {install_cmd}")
                missing_packages.append((package, install_cmd))
        
        if missing_packages:
            print(f"\n⚠️ 发现 {len(missing_packages)} 个缺失的依赖包")
            print("请先安装缺失的包:")
            for package, cmd in missing_packages:
                print(f"   {cmd}")
            return False
        
        print("✅ 所有依赖包检查完成")
        return True
    
    def check_data_file(self):
        """检查数据文件"""
        print(f"\n📊 检查数据文件: {EXCEL_FILE}")
        
        if not os.path.exists(EXCEL_FILE):
            print(f"❌ 数据文件不存在: {EXCEL_FILE}")
            print("请确保数据合并文件存在于当前目录")
            return False
        
        file_size = os.path.getsize(EXCEL_FILE) / 1024 / 1024
        print(f"✅ 数据文件检查完成 ({file_size:.1f} MB)")
        return True
    
    def run_analysis_tool(self, tool_script, tool_name):
        """运行单个分析工具"""
        print(f"\n🚀 运行 {tool_name}...")
        print("-" * 50)
        
        start_time = time.time()
        
        try:
            # 运行Python脚本
            result = subprocess.run([
                sys.executable, tool_script
            ], capture_output=True, text=True, encoding='utf-8')
            
            end_time = time.time()
            duration = end_time - start_time
            
            if result.returncode == 0:
                print(f"✅ {tool_name} 完成 (耗时: {duration:.1f}秒)")
                
                # 显示输出的关键信息
                if result.stdout:
                    lines = result.stdout.strip().split('\n')
                    for line in lines[-5:]:  # 显示最后5行
                        if line.strip():
                            print(f"   📋 {line.strip()}")
                
                self.results[tool_script] = {
                    'status': 'success',
                    'duration': duration,
                    'output': result.stdout
                }
                return True
            else:
                print(f"❌ {tool_name} 运行失败")
                if result.stderr:
                    print(f"错误信息: {result.stderr}")
                
                self.results[tool_script] = {
                    'status': 'failed',
                    'duration': duration,
                    'error': result.stderr
                }
                return False
                
        except Exception as e:
            print(f"❌ 运行 {tool_name} 时出错: {str(e)}")
            self.results[tool_script] = {
                'status': 'error',
                'error': str(e)
            }
            return False
    
    def generate_summary_report(self):
        """生成总结报告"""
        print(f"\n📋 生成综合分析总结报告...")
        
        total_duration = (datetime.now() - self.start_time).total_seconds()
        
        summary_content = f"""
# 🎉 综合数据分析完成报告

## 📊 执行概要

**执行时间**: {self.start_time.strftime('%Y年%m月%d日 %H:%M:%S')}
**总耗时**: {total_duration:.1f} 秒
**数据文件**: {EXCEL_FILE}

## 🔧 分析工具执行结果

"""
        
        for tool_script, tool_name in ANALYSIS_TOOLS:
            if tool_script in self.results:
                result = self.results[tool_script]
                status_icon = "✅" if result['status'] == 'success' else "❌"
                duration = result.get('duration', 0)
                
                summary_content += f"""
### {status_icon} {tool_name}

- **状态**: {result['status']}
- **耗时**: {duration:.1f} 秒
"""
                if result['status'] == 'success':
                    summary_content += f"- **结果**: 成功生成分析结果\n"
                else:
                    summary_content += f"- **错误**: {result.get('error', '未知错误')}\n"
        
        summary_content += f"""

## 📁 生成的文件

根据分析工具的执行情况，可能生成了以下文件和目录：

### 📈 深度分析结果
- `深度分析结果/` - 高级统计分析结果
- `深度分析结果/深度分析洞察_*.txt` - 深度洞察报告

### 🌐 交互式可视化
- `可视化网页/` - 交互式网页文件
- `可视化网页/金融数据交互分析仪表板.html` - 主要可视化网页

### 📄 专业PDF报告
- `PDF报告/` - PDF报告和图表
- `PDF报告/金融数据深度分析报告_*.pdf` - 专业分析报告
- `PDF报告/图表/` - 分析图表

## 🎯 后续操作建议

### 1. 查看深度分析结果
打开 `深度分析结果/` 目录查看详细的统计分析结果

### 2. 浏览交互式可视化
双击 `可视化网页/金融数据交互分析仪表板.html` 在浏览器中查看交互式图表

### 3. 阅读专业报告
打开 `PDF报告/` 目录中的PDF文件，查看完整的专业分析报告

## 📞 技术支持

如果在使用过程中遇到问题，请检查：
1. 所有依赖包是否正确安装
2. 数据文件是否存在且格式正确
3. 磁盘空间是否充足

---

**报告生成时间**: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}
"""
        
        # 保存总结报告
        summary_file = f"综合分析执行报告_{datetime.now().strftime('%Y%m%d_%H%M')}.md"
        with open(summary_file, 'w', encoding='utf-8') as f:
            f.write(summary_content)
        
        print(f"✅ 总结报告已保存: {summary_file}")
        return summary_file
    
    def open_results(self):
        """打开结果文件"""
        print(f"\n🌐 尝试打开生成的结果...")
        
        # 尝试打开交互式可视化网页
        html_file = "可视化网页/金融数据交互分析仪表板.html"
        if os.path.exists(html_file):
            try:
                if sys.platform.startswith('darwin'):  # macOS
                    subprocess.run(['open', html_file])
                elif sys.platform.startswith('win'):  # Windows
                    os.startfile(html_file)
                else:  # Linux
                    subprocess.run(['xdg-open', html_file])
                print(f"✅ 已打开交互式可视化网页")
            except:
                print(f"📋 请手动打开: {html_file}")
        
        # 显示PDF报告位置
        pdf_dir = "PDF报告"
        if os.path.exists(pdf_dir):
            pdf_files = [f for f in os.listdir(pdf_dir) if f.endswith('.pdf')]
            if pdf_files:
                print(f"📄 PDF报告位置: {pdf_dir}/{pdf_files[0]}")
    
    def run_comprehensive_analysis(self):
        """运行综合分析"""
        print("🚀 开始综合数据深度分析")
        print("=" * 60)
        print(f"时间: {self.start_time.strftime('%Y年%m月%d日 %H:%M:%S')}")
        print(f"数据源: {EXCEL_FILE}")
        print("=" * 60)
        
        # 1. 检查环境
        if not self.check_dependencies():
            return False
        
        # 2. 检查数据文件
        if not self.check_data_file():
            return False
        
        # 3. 依次运行分析工具
        success_count = 0
        for tool_script, tool_name in ANALYSIS_TOOLS:
            if self.run_analysis_tool(tool_script, tool_name):
                success_count += 1
        
        # 4. 生成总结报告
        summary_file = self.generate_summary_report()
        
        # 5. 显示最终结果
        print(f"\n🎉 综合分析完成!")
        print(f"📊 成功执行: {success_count}/{len(ANALYSIS_TOOLS)} 个分析工具")
        print(f"⏱️ 总耗时: {(datetime.now() - self.start_time).total_seconds():.1f} 秒")
        print(f"📋 总结报告: {summary_file}")
        
        # 6. 尝试打开结果
        self.open_results()
        
        return success_count == len(ANALYSIS_TOOLS)


def main():
    """主函数"""
    print("🔬 金融数据综合深度分析系统")
    print("=" * 40)
    print("本系统将自动执行以下分析:")
    print("1️⃣ 高级统计分析和深度洞察")
    print("2️⃣ 交互式可视化网页生成")
    print("3️⃣ 专业PDF分析报告生成")
    print("=" * 40)
    
    # 确认执行
    try:
        confirm = input("\n🤔 是否开始综合分析? (y/N): ").lower().strip()
        if confirm not in ['y', 'yes']:
            print("❌ 分析已取消")
            return
    except KeyboardInterrupt:
        print("\n❌ 分析已取消")
        return
    
    # 创建并运行分析器
    runner = ComprehensiveAnalysisRunner()
    success = runner.run_comprehensive_analysis()
    
    if success:
        print(f"\n🎊 恭喜！所有分析工具都执行成功!")
        print(f"📈 您现在可以:")
        print(f"   🌐 查看交互式可视化网页")
        print(f"   📄 阅读专业PDF分析报告")
        print(f"   📊 深入研究统计分析结果")
    else:
        print(f"\n⚠️ 部分分析工具执行失败，请查看错误信息")


if __name__ == "__main__":
    main() 