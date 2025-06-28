import openpyxl
import requests
import time
import logging
import re

logger = logging.getLogger(__name__)

class Translator:
    def __init__(self, api_key, api_provider="DeepSeek"):
        self.api_key = api_key
        self.api_provider = api_provider
        self.headers = {
            'Content-Type': 'application/json',
            'Authorization': f'Bearer {self.api_key}'
        }
        self.api_url = self._get_api_url()

    def _get_api_url(self):
        if self.api_provider == "DeepSeek":
            return 'https://api.deepseek.com/v1/chat/completions'
        # 未来可以在此添加其他API提供商
        raise ValueError(f"不支持的API提供商: {self.api_provider}")

    def _clean_translation(self, text):
        """清理翻译结果，移除开头的序号和点。"""
        return re.sub(r'^\d+\.?\s*', '', text.strip())

    def translate_batch(self, texts, prompt_template, source_language, target_language):
        """使用AI翻译一个文本批次。"""
        if not texts:
            return []

        # 使用模板构建最终的prompt
        prompt = prompt_template.format(
            source_language=source_language,
            target_language=target_language,
            text_to_translate='\n'.join(texts)
        )

        payload = {
            "model": "deepseek-chat",
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.1,
            "stream": False
        }

        try:
            response = requests.post(self.api_url, json=payload, headers=self.headers, timeout=60)
            response.raise_for_status()  # 如果请求失败则抛出HTTPError
            
            response_data = response.json()
            content = response_data['choices'][0]['message']['content']
            
            # 按行分割并清理结果
            translated_lines = content.split('\n')
            cleaned_lines = [self._clean_translation(line) for line in translated_lines]
            
            return cleaned_lines

        except requests.RequestException as e:
            logger.error(f"API请求失败: {e}")
            # 返回与输入相同数量的空字符串，以保持行对应
            return [''] * len(texts)
        except (KeyError, IndexError) as e:
            logger.error(f"解析API响应失败: {e}")
            return [''] * len(texts)

def get_excel_columns(file_path):
    """获取Excel文件中的所有列字母。"""
    try:
        workbook = openpyxl.load_workbook(file_path, read_only=True)
        sheet = workbook.active
        # 返回前26列（A-Z），无论是否有数据
        max_col = 26 
        return [openpyxl.utils.get_column_letter(i) for i in range(1, max_col + 1)]
    except Exception as e:
        logger.error(f"读取Excel列时出错: {e}")
        return []
