"""
测试文本替换功能 - 将所有文本替换为"test"以验证识别和操作
不调用API，直接替换文本
"""
import os
import sys
from ppt_processor import PPTProcessor


def test_text_replacement(input_file: str, output_file: str):
    """
    测试文本替换功能
    
    Args:
        input_file: 输入的PPT文件路径
        output_file: 输出的PPT文件路径
    """
    if not os.path.exists(input_file):
        print(f"❌ 错误：文件 {input_file} 不存在")
        return False
    
    try:
        print("=" * 70)
        print("文本替换测试 - 将所有文本替换为'test'")
        print("=" * 70)
        
        # 1. 打开并解析PPT
        print(f"\n[1/3] 正在打开PPT文件: {input_file}")
        processor = PPTProcessor(input_file)
        
        print("[2/3] 正在提取文本...")
        slides_data = processor.extract_texts()
        
        total_texts = sum(len(slide['texts']) for slide in slides_data)
        print(f"✅ 找到 {len(slides_data)} 张包含可翻译文本的幻灯片，共 {total_texts} 个文本块")
        
        if not slides_data:
            print("⚠️  未找到需要翻译的文本")
            return False
        
        # 2. 显示提取的文本并替换为"test"
        print("\n[3/3] 开始替换文本...")
        replaced_count = 0
        failed_count = 0
        
        for slide_data in slides_data:
            slide_index = slide_data['slide_index']
            texts = slide_data['texts']
            
            print(f"\n幻灯片 {slide_index + 1}:")
            print(f"  包含 {len(texts)} 个文本块")
            
            for item in texts:
                original_text = item['text']
                
                try:
                    # 替换为"test"
                    if item['text_type'] == 'textbox':
                        processor.update_text(
                            slide_index=item['slide_index'],
                            shape_index=item['shape_index'],
                            original_text=original_text,
                            translated_text="test",
                            paragraph_index=item.get('paragraph_index')
                        )
                    elif item['text_type'] == 'group_textbox':
                        processor.update_text(
                            slide_index=item['slide_index'],
                            shape_index=item['shape_index'],
                            original_text=original_text,
                            translated_text="test",
                            paragraph_index=item.get('paragraph_index'),
                            sub_shape_index=item.get('sub_shape_index')
                        )
                    elif item['text_type'] == 'table':
                        processor.update_text(
                            slide_index=item['slide_index'],
                            shape_index=item['shape_index'],
                            original_text=original_text,
                            translated_text="test",
                            row_index=item.get('row_index'),
                            col_index=item.get('col_index')
                        )
                    
                    replaced_count += 1
                    preview = original_text[:40].replace('\n', ' ')
                    print(f"  ✓ [{item['text_type']}] {preview}... -> test")
                    
                except Exception as e:
                    failed_count += 1
                    print(f"  ✗ 替换失败: {original_text[:30]}... 错误: {str(e)}")
        
        # 3. 保存文件
        print(f"\n正在保存到: {output_file}")
        processor.save(output_file)
        
        print("\n" + "=" * 70)
        print("✅ 测试完成！")
        print(f"   - 成功替换: {replaced_count} 个文本块")
        print(f"   - 替换失败: {failed_count} 个文本块")
        print(f"   - 输出文件: {output_file}")
        print("=" * 70)
        
        if failed_count > 0:
            print(f"\n⚠️  警告: 有 {failed_count} 个文本块替换失败，请检查代码")
            return False
        
        return True
    
    except Exception as e:
        print(f"\n❌ 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("使用方法: python3 test_text_replacement.py <input_ppt_file> [output_ppt_file]")
        print("示例: python3 test_text_replacement.py input.pptx output.pptx")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else input_file.replace('.pptx', '_test_replaced.pptx').replace('.ppt', '_test_replaced.pptx')
    
    test_text_replacement(input_file, output_file)

