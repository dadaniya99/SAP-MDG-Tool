#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SAP MDG 工具打包脚本
一键生成 Windows 可执行文件 (.exe)
"""

import subprocess
import sys
import os


def check_pyinstaller():
    """检查是否安装了 PyInstaller"""
    try:
        import PyInstaller
        return True
    except ImportError:
        return False


def install_pyinstaller():
    """安装 PyInstaller"""
    print("📦 正在安装 PyInstaller...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller", "-q"])
    print("✅ PyInstaller 安装完成\n")


def build_exe():
    """打包 exe"""
    
    # 检查/安装 PyInstaller
    if not check_pyinstaller():
        install_pyinstaller()
    
    print("🔨 开始打包...")
    print("=" * 50)
    
    # PyInstaller 参数
    cmd = [
        sys.executable, "-m", "PyInstaller",
        
        # 主程序文件
        "客商数据处理工具_GUI_v3.py",
        
        # 单文件模式（所有内容打包到一个exe）
        "--onefile",
        
        # 不显示控制台窗口（GUI程序）
        "--noconsole",
        
        # 程序名称
        "--name", "SAP_MDG_客商数据处理工具",
        
        # 图标（如果有的话）
        # "--icon", "icon.ico",
        
        # 清理之前的构建
        "--clean",
        
        # 添加数据文件（如果需要）
        # "--add-data", "templates;templates",
        
        # 隐藏导入（openpyxl 依赖）
        "--hidden-import", "openpyxl",
        "--hidden-import", "openpyxl.cell._writer",
        "--hidden-import", "openpyxl.styles",
        "--hidden-import", "openpyxl.utils",
        
        # 优化
        "--strip",
    ]
    
    try:
        subprocess.check_call(cmd)
        print("=" * 50)
        print("✅ 打包成功！")
        print()
        print("📁 输出文件位置:")
        print("   dist/SAP_MDG_客商数据处理工具.exe")
        print()
        print("💡 使用说明:")
        print("   1. 把 exe 文件复制到任意位置")
        print("   2. 双击即可运行（无需安装 Python）")
        print("   3. 可以发给其他人使用")
        print()
        
    except subprocess.CalledProcessError as e:
        print(f"❌ 打包失败: {e}")
        sys.exit(1)


def build_with_upx():
    """使用 UPX 压缩（可选，减小体积）"""
    print("📦 使用 UPX 压缩模式...")
    
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "客商数据处理工具_GUI_v3.py",
        "--onefile",
        "--noconsole",
        "--name", "SAP_MDG_客商数据处理工具",
        "--clean",
        "--hidden-import", "openpyxl",
        "--hidden-import", "openpyxl.cell._writer",
        "--upx-dir", "upx",  # 如果有 UPX 压缩工具
        "--strip",
    ]
    
    subprocess.check_call(cmd)


def main():
    """主函数"""
    print("=" * 50)
    print("SAP MDG 客商数据处理工具 - 打包工具")
    print("=" * 50)
    print()
    
    # 检查当前目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    
    # 检查源文件是否存在
    if not os.path.exists("客商数据处理工具_GUI_v3.py"):
        print("❌ 错误: 找不到 客商数据处理工具_GUI_v3.py")
        print(f"   请确保在目录: {script_dir}")
        sys.exit(1)
    
    print(f"📂 工作目录: {script_dir}")
    print()
    
    # 执行打包
    build_exe()


if __name__ == "__main__":
    main()
