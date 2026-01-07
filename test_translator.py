"""
测试翻译功能
用于验证DeepSeek或OpenAI API配置是否正确
"""
import os
from dotenv import load_dotenv
from translator import Translator

load_dotenv()


def test_translator():
    """测试翻译功能"""
    print("=" * 50)
    print("翻译功能测试")
    print("=" * 50)
    
    # 检查API提供商
    api_provider = os.getenv('API_PROVIDER', 'deepseek').lower()
    print(f"\n使用API提供商: {api_provider}")
    
    # 检查API密钥
    if api_provider == "deepseek":
        api_key = os.getenv('DEEPSEEK_API_KEY')
        if not api_key:
            print("❌ 错误: 未设置 DEEPSEEK_API_KEY 环境变量")
            print("请在 .env 文件中设置: DEEPSEEK_API_KEY=your_key_here")
            return False
        print(f"✅ DeepSeek API Key 已配置 (长度: {len(api_key)})")
    else:
        api_key = os.getenv('OPENAI_API_KEY')
        if not api_key:
            print("❌ 错误: 未设置 OPENAI_API_KEY 环境变量")
            print("请在 .env 文件中设置: OPENAI_API_KEY=your_key_here")
            return False
        print(f"✅ OpenAI API Key 已配置 (长度: {len(api_key)})")
    
    # 测试翻译
    try:
        print("\n正在初始化翻译器...")
        translator = Translator(provider=api_provider)
        print("✅ 翻译器初始化成功")
        
        # 测试单个文本翻译
        test_text = "科学研究"
        print(f"\n测试翻译文本: {test_text}")
        print("正在调用API...")
        
        result = translator.translate_text(test_text)
        print(f"✅ 翻译结果: {result}")
        
        # 测试幻灯片翻译
        print("\n测试幻灯片级翻译...")
        slide_texts = ["科学研究", "农业", "畜牧业"]
        print(f"待翻译文本: {slide_texts}")
        
        translation_map = translator.translate_slide(slide_texts, slide_index=0)
        print("✅ 翻译结果:")
        for original, translated in translation_map.items():
            print(f"  {original} -> {translated}")
        
        print("\n" + "=" * 50)
        print("✅ 所有测试通过！翻译功能正常工作")
        print("=" * 50)
        return True
        
    except Exception as e:
        print(f"\n❌ 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == '__main__':
    test_translator()

