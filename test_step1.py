"""
第一步测试：用python-pptx打开PPT，修改文本并保存
这是一个简单的测试脚本，用于验证PPT处理功能
"""
from ppt_processor import PPTProcessor
import sys
import os


def test_ppt_processing(input_file: str, output_file: str):
    """
    测试PPT处理功能
    
    Args:
        input_file: 输入的PPT文件路径
        output_file: 输出的PPT文件路径
    """
    if not os.path.exists(input_file):
        print(f"错误：文件 {input_file} 不存在")
        return False
    
    try:
        print(f"正在打开PPT文件: {input_file}")
        processor = PPTProcessor(input_file)
        
        print("正在提取文本...")
        slides_data = processor.extract_texts()
        
        print(f"找到 {len(slides_data)} 张包含可翻译文本的幻灯片")
        
        # 显示提取的文本
        for slide_data in slides_data:
            slide_index = slide_data['slide_index']
            texts = slide_data['texts']
            print(f"\n幻灯片 {slide_index + 1}:")
            for i, text_item in enumerate(texts):
                print(f"  [{i+1}] {text_item['text'][:50]}...")
        
        # 简单测试：修改第一个文本（添加测试标记）
        if slides_data and slides_data[0]['texts']:
            first_text = slides_data[0]['texts'][0]
            original = first_text['text']
            test_translation = f"[TEST] {original}"
            
            print(f"\n测试修改文本:")
            print(f"  原文: {original}")
            print(f"  修改为: {test_translation}")
            
            processor.update_text(
                slide_index=first_text['slide_index'],
                shape_index=first_text['shape_index'],
                original_text=original,
                translated_text=test_translation,
                paragraph_index=first_text.get('paragraph_index')
            )
        
        print(f"\n正在保存到: {output_file}")
        processor.save(output_file)
        
        print("✅ 测试成功！")
        print(f"请检查输出文件: {output_file}")
        return True
    
    except Exception as e:
        print(f"❌ 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("使用方法: python test_step1.py <input_ppt_file> [output_ppt_file]")
        print("示例: python test_step1.py test.pptx test_output.pptx")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else input_file.replace('.pptx', '_test.pptx').replace('.ppt', '_test.pptx')
    
    test_ppt_processing(input_file, output_file)

