#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
一鍵生成所有 Excel 檔案
執行此腳本將自動建立所有專案管理 Excel 模板
"""

import subprocess
import sys
import os

def check_dependencies():
    """檢查必要套件"""
    print("🔍 檢查必要套件...")
    try:
        import openpyxl
        print("✅ openpyxl 已安裝")
        return True
    except ImportError:
        print("❌ openpyxl 未安裝")
        print("\n正在安裝 openpyxl...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
            print("✅ openpyxl 安裝成功")
            return True
        except Exception as e:
            print(f"❌ 安裝失敗: {e}")
            print("\n請手動執行: pip install openpyxl")
            return False

def run_script(script_name, description):
    """執行 Python 腳本"""
    print(f"\n{'='*60}")
    print(f"🚀 {description}")
    print(f"{'='*60}")
    try:
        subprocess.check_call([sys.executable, script_name])
        print(f"✅ {description} 完成")
        return True
    except Exception as e:
        print(f"❌ {description} 失敗: {e}")
        return False

def main():
    """主程式"""
    print("""
╔══════════════════════════════════════════════════════════╗
║                                                          ║
║     I&C Project Management - Excel 模板生成工具          ║
║                                                          ║
║     此工具將自動建立以下 Excel 檔案:                      ║
║     1. Dashboard.xlsx - 主儀表板                         ║
║     2. Task_Tracker.xlsx - 任務追蹤表                    ║
║     3. Weekly_Report_Template.xlsx - 週報模板            ║
║                                                          ║
╚══════════════════════════════════════════════════════════╝
    """)
    
    # 檢查依賴
    if not check_dependencies():
        print("\n❌ 無法繼續，請先安裝必要套件")
        return False
    
    # 執行腳本
    scripts = [
        ("generate_dashboard.py", "建立 Dashboard.xlsx"),
        ("generate_task_tracker.py", "建立 Task_Tracker.xlsx"),
        ("generate_weekly_report.py", "建立 Weekly_Report_Template.xlsx"),
    ]
    
    success_count = 0
    for script, description in scripts:
        if os.path.exists(script):
            if run_script(script, description):
                success_count += 1
        else:
            print(f"⚠️  找不到腳本: {script}")
    
    # 總結
    print(f"\n{'='*60}")
    print(f"📊 執行總結")
    print(f"{'='*60}")
    print(f"成功: {success_count}/{len(scripts)}")
    
    if success_count == len(scripts):
        print("\n✅ 所有 Excel 檔案已成功建立！")
        print("\n📁 生成的檔案:")
        print("   • Dashboard.xlsx")
        print("   • Task_Tracker.xlsx")
        print("   • Weekly_Report_Template.xlsx")
        print("\n📖 下一步:")
        print("   1. 開啟 Excel 檔案檢視內容")
        print("   2. 根據您的專案需求自訂資料")
        print("   3. 上傳至 OneDrive 或 SharePoint")
        print("   4. 整合至 Microsoft Teams")
        print("\n📚 參考文件:")
        print("   • EXCEL_DASHBOARD_GUIDE.md - Excel 建立指南")
        print("   • USER_GUIDE.md - 使用手冊")
        print("   • EXAMPLE_PROJECT.md - 範例專案說明")
        return True
    else:
        print("\n⚠️  部分檔案建立失敗，請檢查錯誤訊息")
        return False

if __name__ == '__main__':
    try:
        success = main()
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\n⚠️  使用者中斷執行")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ 發生錯誤: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
