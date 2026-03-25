import os
import threading
import time
import uuid
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from ppt_processor import PPTProcessor
from translator import Translator

app = Flask(__name__)
CORS(app)

os.makedirs('uploads', exist_ok=True)
os.makedirs('outputs', exist_ok=True)

jobs = {}
jobs_lock = threading.Lock()


def set_job(job_id, **fields):
    with jobs_lock:
        job = jobs.setdefault(job_id, {})
        job.update(fields)
        job['updated_at'] = int(time.time())
        return dict(job)


def get_job(job_id):
    with jobs_lock:
        job = jobs.get(job_id)
        return dict(job) if job else None


def process_translation_job(job_id: str, input_path: str):
    try:
        set_job(job_id, status='processing', stage='正在解析 PPT')
        processor = PPTProcessor(input_path)
        slides_data = processor.extract_texts()

        set_job(job_id, status='processing', stage='正在初始化翻译模型', slides_total=len(slides_data))
        api_provider = os.getenv('API_PROVIDER', 'deepseek').lower()
        translator = Translator(provider=api_provider)

        translated_slides = 0
        for slide_data in slides_data:
            slide_index = slide_data['slide_index']
            texts = [item['text'] for item in slide_data['texts']]
            slide_texts = processor.get_slide_texts(slide_index)

            if slide_texts:
                set_job(
                    job_id,
                    status='processing',
                    stage=f'正在翻译第 {translated_slides + 1}/{len(slides_data)} 页',
                    slides_done=translated_slides,
                    slides_total=len(slides_data),
                )
                text_map = translator.translate_slide(slide_texts, slide_index)
                for item in slide_data['texts']:
                    original_text = item['text']
                    if original_text not in text_map:
                        continue
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
            translated_slides += 1

        set_job(job_id, status='processing', stage='正在生成输出文件', slides_done=translated_slides)
        output_path = f'outputs/{job_id}_translated.pptx'
        processor.save(output_path)
        set_job(
            job_id,
            status='completed',
            stage='已完成',
            success=True,
            file_id=job_id,
            output_file=output_path,
            slides_processed=len(slides_data),
            slides_done=translated_slides,
        )
    except Exception as e:
        set_job(job_id, status='failed', stage='失败', error=str(e), success=False)


@app.route('/', methods=['GET'])
def index():
    return send_file('frontend/index.html')


@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})


@app.route('/translate', methods=['POST'])
def translate_ppt():
    try:
        if 'file' not in request.files:
            return jsonify({'error': '没有上传文件'}), 400
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '文件名为空'}), 400
        if not file.filename.endswith(('.pptx', '.ppt')):
            return jsonify({'error': '只支持PPT/PPTX文件'}), 400

        job_id = str(uuid.uuid4())
        input_path = f'uploads/{job_id}.pptx'
        file.save(input_path)
        set_job(job_id, status='queued', stage='文件已上传，等待处理', filename=file.filename, success=True)

        worker = threading.Thread(target=process_translation_job, args=(job_id, input_path), daemon=True)
        worker.start()

        return jsonify({'success': True, 'job_id': job_id, 'status': 'queued', 'stage': '文件已上传，等待处理'}), 202
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/status/<job_id>', methods=['GET'])
def job_status(job_id):
    job = get_job(job_id)
    if not job:
        return jsonify({'error': '任务不存在'}), 404
    return jsonify(job)


@app.route('/download/<file_id>', methods=['GET'])
def download_file(file_id):
    try:
        file_path = f'outputs/{file_id}_translated.pptx'
        if not os.path.exists(file_path):
            return jsonify({'error': '文件不存在'}), 404
        return send_file(file_path, as_attachment=True, download_name=f'translated_{file_id}.pptx')
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=False, use_reloader=False, host='127.0.0.1', port=5014)
