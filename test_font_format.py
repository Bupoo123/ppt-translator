"""
测试字体格式保留功能
验证翻译后是否保留了字体颜色、大小，并设置了Arial字体
"""
from ppt_processor import PPTProcessor
import sys
import os


def test_font_format(input_file: str, output_file: str):
    """
    测试字体格式保留
    """
    if not os.path.exists(input_file):
        print(f"❌ 错误：文件 {input_file} 不存在")
        return False
    
    try:
        print("=" * 70)
        print("字体格式保留测试")
        print("=" * 70)
        
        processor = PPTProcessor(input_file)
        slides_data = processor.extract_texts()
        
        print(f"找到 {len(slides_data)} 张包含可翻译文本的幻灯片")
        
        # 只处理前3张幻灯片作为测试
        test_count = 0
        for slide_data in slides_data[:3]:
            slide_index = slide_data['slide_index']
            print(f"\n处理幻灯片 {slide_index + 1}:")
            
            for item in slide_data['texts'][:5]:  # 每张幻灯片只处理前5个文本
                original_text = item['text']
                translated_text = f"[TEST] {original_text[:20]}..."
                
                print(f"  翻译: {original_text[:30]}... -> {translated_text[:30]}...")
                
                try:
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
                    
                    test_count += 1
                    print(f"    ✓ 成功")
                except Exception as e:
                    print(f"    ✗ 失败: {str(e)}")
        
        print(f"\n保存到: {output_file}")
        processor.save(output_file)
        
        print("\n" + "=" * 70)
        print(f"✅ 测试完成！")
        print(f"   - 处理了 {test_count} 个文本块")
        print(f"   - 输出文件: {output_file}")
        print(f"\n请打开输出文件检查：")
        print(f"  1. 字体颜色是否保留")
        print(f"  2. 字体大小是否保留")
        print(f"  3. 英文字体是否为Arial")
        print("=" * 70)
        
        return True
    
    except Exception as e:
        print(f"\n❌ 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("使用方法: python3 test_font_format.py <input_ppt_file> [output_ppt_file]")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else input_file.replace('.pptx', '_font_test.pptx').replace('.ppt', '_font_test.pptx')
    
    test_font_format(input_file, output_file)

