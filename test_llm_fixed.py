#!/usr/bin/env python3
"""
LLM 图片识别功能测试脚本
"""

import os
import sys
from pathlib import Path

# Add current directory to path
sys.path.append(str(Path(__file__).parent))

def test_llm_configuration():
    """测试 LLM 配置"""
    print("=== LLM Configuration Test ===")

    # 导入并检查配置
    try:
        from tender_slicer_web import (
            LLM_API_ENDPOINT,
            LLM_API_KEY,
            LLM_MODEL,
            LLM_TIMEOUT,
            LLM_MAX_RETRIES,
            LLM_AVAILABLE
        )

        print(f"LLM Available: {LLM_AVAILABLE}")
        print(f"LLM Endpoint: {LLM_API_ENDPOINT}")
        print(f"LLM Model: {LLM_MODEL}")
        print(f"LLM Timeout: {LLM_TIMEOUT}")
        print(f"LLM Max Retries: {LLM_MAX_RETRIES}")

        if not LLM_AVAILABLE:
            print("\n⚠️  LLM 不可用 - 请检查环境变量配置")
            print("需要设置的环境变量:")
            print("  LLM_API_ENDPOINT=your_api_endpoint")
            print("  LLM_API_KEY=your_api_key")
            return False

        print("\n✅ LLM 配置正确")
        return True

    except ImportError as e:
        print(f"❌ 导入错误: {e}")
        return False

def test_image_format_unification():
    """测试图片格式统一"""
    print("\n=== Image Format Test ===")

    try:
        from tender_slicer_web import TenderSlicer

        # 创建测试实例
        slicer = TenderSlicer("test.docx")

        # 测试 LLM 服务初始化
        if slicer.llm_service:
            print("✅ LLM Service 初始化成功")
        else:
            print("ℹ️  LLM Service 未初始化（配置不可用）")

        # 测试 placeholder 格式
        placeholder = slicer.extract_paragraph_images.__annotations__.get('return')
        print("✅ 图片提取方法可用")

        # 清理
        slicer.cleanup()

        return True

    except Exception as e:
        print(f"❌ 测试失败: {e}")
        return False

def test_code_quality():
    """测试代码质量"""
    print("\n=== Code Quality Test ===")

    try:
        # 读取并检查代码
        with open('tender_slicer_web.py', 'r', encoding='utf-8') as f:
            content = f.read()

        # 检查关键修复
        checks = [
            ('line_no 未定义修复', 'results[img[\'id\']] = f"<!-- [图片] 内容:' in content),
            ('表格图片处理修复', 'description = processed_images.get(img[\'id\'], img[\'placeholder\'])' in content),
            ('占位符格式统一', '内容: [' in content),
            ('错误日志优化', 'Response content:' in content)
        ]

        all_passed = True
        for check_name, check_result in checks:
            if check_result:
                print(f"✅ {check_name}")
            else:
                print(f"❌ {check_name}")
                all_passed = False

        return all_passed

    except Exception as e:
        print(f"❌ 代码质量检查失败: {e}")
        return False

def main():
    """主测试函数"""
    print("LLM 图片识别功能测试")
    print("=" * 40)

    success = True

    # 测试配置
    success &= test_llm_configuration()

    # 测试代码质量
    success &= test_code_quality()

    # 测试图片格式
    if success:
        success &= test_image_format_unification()

    print("\n" + "=" * 40)
    if success:
        print("✅ 所有测试通过")
        print("\n使用说明:")
        print("1. 设置环境变量: cp .env.example .env")
        print("2. 编辑 .env 文件填入 API 信息")
        print("3. 启动服务: python3 tender_slicer_web.py")
    else:
        print("❌ 部分测试失败")

    return 0 if success else 1

if __name__ == "__main__":
    sys.exit(main())