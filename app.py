"""
Flask后端应用
"""
import os
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from ppt_processor import PPTProcessor
from translator import Translator
import uuid

app = Flask(__name__)
CORS(app)

# 创建必要的目录
os.makedirs('uploads', exist_ok=True)
os.makedirs('outputs', exist_ok=True)


@app.route('/health', methods=['GET'])
def health():
    """健康检查"""
    return jsonify({'status': 'ok'})


@app.route('/translate', methods=['POST'])
def translate_ppt():
    """
    翻译PPT文件
    
    请求：
    - file: PPT文件（multipart/form-data）
    
    返回：
    - output_file: 翻译后的PPT文件路径
    """
    try:
        # 检查文件
        if 'file' not in request.files:
            return jsonify({'error': '没有上传文件'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '文件名为空'}), 400
        
        if not file.filename.endswith(('.pptx', '.ppt')):
            return jsonify({'error': '只支持PPT/PPTX文件'}), 400
        
        # 保存上传的文件
        file_id = str(uuid.uuid4())
        input_path = f'uploads/{file_id}.pptx'
        file.save(input_path)
        
        # 处理PPT
        processor = PPTProcessor(input_path)
        slides_data = processor.extract_texts()
        
        # 翻译（使用DeepSeek API）
        translator = Translator()
        
        for slide_data in slides_data:
            slide_index = slide_data['slide_index']
            texts = [item['text'] for item in slide_data['texts']]
            
            # 获取幻灯片所有文本用于上下文翻译
            slide_texts = processor.get_slide_texts(slide_index)
            
            if slide_texts:
                # 翻译整个幻灯片（返回字典映射）
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
        
        # 保存翻译后的文件
        output_path = f'outputs/{file_id}_translated.pptx'
        processor.save(output_path)
        
        return jsonify({
            'success': True,
            'file_id': file_id,
            'output_file': output_path,
            'slides_processed': len(slides_data)
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/download/<file_id>', methods=['GET'])
def download_file(file_id):
    """下载翻译后的文件"""
    try:
        file_path = f'outputs/{file_id}_translated.pptx'
        if not os.path.exists(file_path):
            return jsonify({'error': '文件不存在'}), 404
        
        return send_file(file_path, as_attachment=True, 
                        download_name=f'translated_{file_id}.pptx')
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5014)

