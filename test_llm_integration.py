#!/usr/bin/env python3
"""
LLM 集成测试脚本
"""

import os
import sys
from pathlib import Path

# Add current directory to path
sys.path.append(str(Path(__file__).parent))

from tender_slicer_web import TenderSlicer, LLM_AVAILABLE, LLM_API_ENDPOINT

def test_llm_configuration():
    """测试 LLM 配置"""
    print("=== LLM Configuration Test ===")
    print(f"LLM Available: {LLM_AVAILABLE}")
    print(f"LLM Endpoint: {LLM_API_ENDPOINT}")

    if not LLM_AVAILABLE:
        print("\n⚠️  LLM 不可用 - 请检查环境变量配置")
        print("需要设置的环境变量:")
        print("  LLM_API_ENDPOINT=your_api_endpoint")
        print("  LLM_API_KEY=your_api_key")
        return False

    print("\n✅ LLM 配置正确")
    return True

def test_image_extraction():
    """测试图片提取功能"""
    print("\n=== Image Extraction Test ===")
    # 创建一个测试用的 TenderSlicer 实例
    # 注意：这里不加载实际文档，只测试初始化
    slicer = TenderSlicer("test.docx")  # 文件不存在，只测试初始化
    print(f"LLM service initialized: {slicer.llm_service is not None}")

    # 测试 cleanup
    slicer.cleanup()
    print("Cleanup completed")

    return True

def main():
    """主测试函数"""
    print("LLM Integration Test")
    print("=" * 30)

    success = True

    # 测试配置
    success &= test_llm_configuration()

    # 测试图片提取（如果配置可用）
    if LLM_AVAILABLE:
        success &= test_image_extraction()

    print("\n" + "=" * 30)
    if success:
        print("✅ 所有测试通过")
    else:
        print("❌ 部分测试失败")

    return 0 if success else 1

if __name__ == "__main__":
    sys.exit(main())