"""
翻译模块 - 负责AI翻译功能
使用DeepSeek API
"""
import os
from typing import List, Dict
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()


class Translator:
    """翻译器类 - 使用DeepSeek API"""
    
    def __init__(self):
        """
        初始化翻译器
        """
        # DeepSeek API配置
        api_key = os.getenv('DEEPSEEK_API_KEY')
        if not api_key:
            raise ValueError("请设置 DEEPSEEK_API_KEY 环境变量")
        
        # DeepSeek API endpoint - 使用最新V3.2版本
        # base_url 不带 /v1，因为 OpenAI SDK 会自动添加 /v1/chat/completions
        self.client = OpenAI(
            api_key=api_key,
            base_url="https://api.deepseek.com"
        )
        self.model = "deepseek-v3.2"  # 使用最新 V3.2 模型
    
    def translate_slide(self, texts: List[str], slide_index: int) -> Dict[str, str]:
        """
        翻译整个幻灯片的文本（上下文感知）
        
        Args:
            texts: 幻灯片中的文本列表
            slide_index: 幻灯片索引
            
        Returns:
            翻译映射字典 {原文: 译文}
        """
        if not texts:
            return {}
        
        # 构建提示词
        prompt = self._build_prompt(texts, slide_index)
        
        # 调用API
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {
                    "role": "system",
                    "content": "You are a professional scientific presentation translator. Translate Chinese text into natural, concise English used in PowerPoint slides."
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            temperature=0.3
        )
        
        # 解析响应
        translated_text = response.choices[0].message.content.strip()
        
        # 解析翻译结果
        translated_lines = self._parse_translation_result(translated_text, texts)
        
        # 创建映射字典
        translation_map = {}
        for i, original in enumerate(texts):
            if i < len(translated_lines):
                translation_map[original] = translated_lines[i]
            else:
                # 如果行数不匹配，使用最后一个翻译结果
                translation_map[original] = translated_lines[-1] if translated_lines else original
        
        return translation_map
    
    def _parse_translation_result(self, result: str, original_texts: List[str]) -> List[str]:
        """
        解析翻译结果
        
        Args:
            result: API返回的翻译结果
            original_texts: 原始文本列表
            
        Returns:
            翻译后的文本列表
        """
        lines = []
        
        # 移除可能的编号前缀（如 "1. ", "2. " 等）
        for line in result.split('\n'):
            line = line.strip()
            if not line:
                continue
            
            # 移除编号前缀
            if line and line[0].isdigit():
                # 匹配 "1. ", "2. " 等格式
                import re
                line = re.sub(r'^\d+\.\s*', '', line)
            
            lines.append(line)
        
        # 如果行数匹配，直接返回
        if len(lines) == len(original_texts):
            return lines
        
        # 如果只有一行，可能是所有文本合并翻译了
        if len(lines) == 1 and len(original_texts) > 1:
            # 尝试按句号分割
            sentences = [s.strip() for s in lines[0].split('.') if s.strip()]
            if len(sentences) == len(original_texts):
                return sentences
        
        # 如果还是不匹配，返回解析出的行（可能不完整）
        return lines if lines else original_texts
    
    def _build_prompt(self, texts: List[str], slide_index: int) -> str:
        """
        构建翻译提示词
        
        Args:
            texts: 文本列表
            slide_index: 幻灯片索引
            
        Returns:
            提示词字符串
        """
        texts_str = '\n'.join([f"{i+1}. {text}" for i, text in enumerate(texts)])
        
        prompt = f"""Translate the following Chinese text from slide {slide_index + 1} into natural, concise English used in PowerPoint slides.

Rules:
- Keep it short and presentation-style
- Do NOT add explanations
- Do NOT change numbers or symbols
- Preserve bullet structure
- Use consistent terminology within the same slide
- Return ONLY the translated text, one item per line, in the same order as the input
- Do NOT add line numbers or prefixes
- Each line should be a direct translation of the corresponding Chinese text

Chinese text:
{texts_str}

Translated English (one line per item, same order):"""
        
        return prompt
    
    def translate_text(self, text: str) -> str:
        """
        翻译单个文本（简单模式，用于测试）
        
        Args:
            text: 待翻译文本
            
        Returns:
            翻译后的文本
        """
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {
                    "role": "system",
                    "content": "You are a professional translator. Translate Chinese to English for PowerPoint presentations. Keep it concise and natural."
                },
                {
                    "role": "user",
                    "content": f"Translate this text: {text}\n\nReturn only the translation, no explanations."
                }
            ],
            temperature=0.3
        )
        
        return response.choices[0].message.content.strip()

