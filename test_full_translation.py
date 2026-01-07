"""
完整翻译流程测试
测试PPT文件的完整翻译功能
"""
import os
import sys
from dotenv import load_dotenv
from ppt_processor import PPTProcessor
from translator import Translator

load_dotenv()


def test_full_translation(input_file: str, output_file: str):
    """
    测试完整的PPT翻译流程
    
    Args:
        input_file: 输入的PPT文件路径
        output_file: 输出的PPT文件路径
    """
    if not os.path.exists(input_file):
        print(f"❌ 错误：文件 {input_file} 不存在")
        return False
    
    try:
        print("=" * 60)
        print("PPT完整翻译测试")
        print("=" * 60)
        
        # 1. 打开并解析PPT
        print(f"\n[1/4] 正在打开PPT文件: {input_file}")
        processor = PPTProcessor(input_file)
        
        print("[2/4] 正在提取文本...")
        slides_data = processor.extract_texts()
        
        total_texts = sum(len(slide['texts']) for slide in slides_data)
        print(f"✅ 找到 {len(slides_data)} 张包含可翻译文本的幻灯片，共 {total_texts} 个文本块")
        
        if not slides_data:
            print("⚠️  未找到需要翻译的文本")
            return False
        
        # 2. 初始化翻译器
        print("\n[3/4] 正在初始化翻译器...")
        api_provider = os.getenv('API_PROVIDER', 'deepseek').lower()
        translator = Translator(provider=api_provider)
        print(f"✅ 使用 {api_provider.upper()} API")
        
        # 3. 翻译每个幻灯片
        print("\n[4/4] 开始翻译...")
        translated_count = 0
        
        for slide_data in slides_data:
            slide_index = slide_data['slide_index']
            texts = [item['text'] for item in slide_data['texts']]
            
            print(f"\n  处理幻灯片 {slide_index + 1}/{len(slides_data)}...")
            print(f"    包含 {len(texts)} 个文本块")
            
            # 获取幻灯片所有文本用于上下文翻译
            slide_texts = processor.get_slide_texts(slide_index)
            
            if slide_texts:
                # 翻译整个幻灯片
                print(f"    正在调用API翻译...")
                text_map = translator.translate_slide(slide_texts, slide_index)
                
                # 更新PPT中的文本
                for item in slide_data['texts']:
                    original_text = item['text']
                    if original_text in text_map:
                        translated_text = text_map[original_text]
                        
                        if item['text_type'] == 'textbox':
                            processor.update_text(
                                slide_index=item['slide_index'],
                                shape_index=item['shape_index'],
                                original_text=original_text,
                                translated_text=translated_text,
                                paragraph_index=item.get('paragraph_index')
                            )
                        elif item['text_type'] == 'group_textbox':
                            processor.update_text(
                                slide_index=item['slide_index'],
                                shape_index=item['shape_index'],
                                original_text=original_text,
                                translated_text=translated_text,
                                paragraph_index=item.get('paragraph_index'),
                                sub_shape_index=item.get('sub_shape_index')
                            )
                        elif item['text_type'] == 'table':
                            processor.update_text(
                                slide_index=item['slide_index'],
                                shape_index=item['shape_index'],
                                original_text=original_text,
                                translated_text=translated_text,
                                row_index=item.get('row_index'),
                                col_index=item.get('col_index')
                            )
                        
                        translated_count += 1
                        print(f"    ✓ {original_text[:30]}... -> {translated_text[:30]}...")
        
        # 4. 保存翻译后的文件
        print(f"\n正在保存翻译后的文件: {output_file}")
        processor.save(output_file)
        
        print("\n" + "=" * 60)
        print(f"✅ 翻译完成！")
        print(f"   - 处理了 {len(slides_data)} 张幻灯片")
        print(f"   - 翻译了 {translated_count} 个文本块")
        print(f"   - 输出文件: {output_file}")
        print("=" * 60)
        
        return True
    
    except Exception as e:
        print(f"\n❌ 翻译失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("使用方法: python3 test_full_translation.py <input_ppt_file> [output_ppt_file]")
        print("示例: python3 test_full_translation.py test.pptx translated.pptx")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else input_file.replace('.pptx', '_translated.pptx').replace('.ppt', '_translated.pptx')
    
    test_full_translation(input_file, output_file)

