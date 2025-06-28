import requests
import json
import logging
import re
import time

logger = logging.getLogger(__name__)

# A unique separator that is unlikely to appear in the text.
LINE_SEPARATOR = "|||---|||"

def fetch_gemini_models(api_key: str, proxy_config: dict = None) -> list[str]:
    api_url = "https://generativelanguage.googleapis.com/v1beta/openai/models"
    headers = {"Authorization": f"Bearer {api_key}"}
    proxies = None

    if proxy_config:
        proxy_type = proxy_config.get("type", "http").lower()
        address = proxy_config.get("address")
        port = proxy_config.get("port")
        username = proxy_config.get("username")
        password = proxy_config.get("password")
        if address and port:
            auth = f"{username}:{password}@" if username and password else ""
            proxy_url = f"{proxy_type}://{auth}{address}:{port}"
            proxies = {"http": proxy_url, "https": proxy_url}
            logger.info(f"正在通过代理 {address}:{port} 获取Gemini模型列表...")

    try:
        response = requests.get(api_url, headers=headers, proxies=proxies, timeout=30)
        response.raise_for_status()
        data = response.json()
        model_ids = [model['id'] for model in data.get('data', []) if 'gemini' in model.get('id')]
        model_ids.sort(key=lambda x: ('pro' not in x, 'flash' in x, x), reverse=False)
        logger.info(f"成功获取到 {len(model_ids)} 个Gemini模型。")
        return model_ids
    except requests.exceptions.HTTPError as e:
        error_msg = f"HTTP {e.response.status_code}"
        try:
            error_details = e.response.json().get('error', {}).get('message', e.response.text)
            error_msg += f": {error_details}"
        except json.JSONDecodeError:
            error_msg += f": {e.response.text}"
        logger.error(f"获取Gemini模型列表失败: {error_msg}")
        raise Exception(f"获取Gemini模型列表失败: {error_msg}") from e
    except Exception as e:
        logger.error(f"获取Gemini模型列表时发生未知错误: {e}")
        raise Exception(f"获取Gemini模型列表时发生未知错误: {e}") from e

def fetch_deepseek_models(api_key: str, proxy_config: dict = None) -> list[str]:
    logger.info("成功获取到DeepSeek的静态模型列表。")
    return ['deepseek-chat', 'deepseek-reasoner']

class Translator:
    def __init__(self, api_key, model_id, api_provider="Custom", custom_api_url=None, proxy_config=None):
        if not api_key or not model_id:
            raise ValueError("API密钥和模型ID不能为空。")
        
        self.api_key = api_key
        self.model_id = model_id
        self.api_provider = api_provider
        self.session = requests.Session()

        if proxy_config:
            proxy_type = proxy_config.get("type", "http").lower()
            address = proxy_config.get("address")
            port = proxy_config.get("port")
            username = proxy_config.get("username")
            password = proxy_config.get("password")
            
            if not address or not port:
                logger.warning("代理配置不完整，已忽略。")
            else:
                auth = f"{username}:{password}@" if username and password else ""
                proxy_url = f"{proxy_type}://{auth}{address}:{port}"
                self.session.proxies = {"http": proxy_url, "https": proxy_url}
                logger.info(f"翻译器已配置代理: {proxy_type}://{address}:{port}")

        if self.api_provider == "Gemini":
            self.api_url = "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions"
            logger.info("翻译器已在 [Gemini - OpenAI兼容模式]下初始化。")
        elif self.api_provider == "DeepSeek":
            self.api_url = "https://api.deepseek.com/chat/completions"
            logger.info("翻译器已在 [DeepSeek模式]下初始化。")
        else: # Custom
            if not custom_api_url:
                raise ValueError("自定义模式下必须提供API URL。")
            self.api_url = custom_api_url
            logger.info(f"翻译器已在 [自定义模式]下初始化, URL: {self.api_url}")
        
        self.session.headers.update({
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}"
        })

    def _prepare_payload(self, prompt):
        return {
            "model": self.model_id,
            "messages": [{"role": "user", "content": prompt}],
            "stream": False,
            "temperature": 0.1,
            "top_p": 0.9
        }

    def _parse_response(self, response_data):
        try:
            if 'choices' in response_data and response_data['choices']:
                message = response_data['choices'][0].get('message', {})
                content = message.get('content', '')
                return content.strip()
            else:
                logger.error(f"API响应格式不正确: {response_data}")
                error_info = response_data.get('error', {})
                return f"[API响应格式错误: {error_info.get('message', '未知')}]"
        except (KeyError, IndexError, TypeError) as e:
            logger.error(f"解析JSON响应时出错: {e}")
            return "[解析响应时出错]"

    def translate_batch(self, sources: list, prompt_template: str, source_language: str, target_language: str) -> list:
        if not sources:
            return []

        # Use the unique separator to join the source texts
        text_to_translate = LINE_SEPARATOR.join(sources)
        final_prompt = prompt_template.format(
            source_language=source_language,
            target_language=target_language,
            text_to_translate=text_to_translate,
            line_separator=LINE_SEPARATOR
        )
        
        payload = self._prepare_payload(final_prompt)
        logger.info(f"--- [开始批量翻译] ({len(sources)} 行) ---")

        max_retries = 3
        for attempt in range(max_retries):
            try:
                response = self.session.post(self.api_url, json=payload, timeout=180)
                response.raise_for_status()
                response_data = response.json()
                
                raw_content = self._parse_response(response_data)
                
                if raw_content.startswith("[") and raw_content.endswith("]"):
                    logger.error(f"API返回解析错误: {raw_content}")
                    return [raw_content] * len(sources)

                # Split the response using the unique separator
                translations = raw_content.split(LINE_SEPARATOR)
                
                if len(translations) == len(sources):
                    logger.info(f"--- [批量翻译成功] ({len(translations)} 行) ---")
                    return [t.strip() for t in translations]
                else:
                    error_msg = f"[翻译结果行数校验失败] 预期 {len(sources)} 行, 收到 {len(translations)} 行。"
                    logger.error(error_msg)
                    logger.debug(f"原始返回内容: {raw_content}")
                    return [error_msg] * len(sources)

            except requests.exceptions.HTTPError as e:
                if e.response.status_code == 429 and attempt < max_retries - 1:
                    retry_delay = 60
                    try:
                        error_json = e.response.json()
                        if isinstance(error_json, dict):
                            details = error_json.get('error', {}).get('details', [])
                            for detail in details:
                                if detail.get('@type') == 'type.googleapis.com/google.rpc.RetryInfo':
                                    delay_str = detail.get('retryDelay', '60s')
                                    if 's' in delay_str:
                                        retry_delay = int(delay_str.replace('s', ''))
                                    break
                    except Exception:
                        pass
                    
                    logger.warning(f"触发API速率限制。将在 {retry_delay} 秒后重试 (尝试 {attempt + 2}/{max_retries})...")
                    time.sleep(retry_delay)
                else:
                    error_message = f"[HTTP错误 {e.response.status_code}]"
                    try:
                        error_json = e.response.json()
                        error_details = error_json.get('error', {}).get('message', str(error_json))
                    except json.JSONDecodeError:
                        error_details = e.response.text
                    error_message += f": {error_details}"
                    logger.critical(f"批量翻译失败: {error_message}")
                    return [error_message] * len(sources)
            
            except requests.exceptions.RequestException as e:
                logger.critical(f"网络层请求API失败: {e}")
                return [f"[网络错误: {e}]"] * len(sources)
            
            except Exception as e:
                logger.critical(f"处理API时发生未知错误: {e}", exc_info=True)
                return [f"[未知错误: {e}]"] * len(sources)
        
        return ["[批量翻译失败: 已达到最大重试次数]"] * len(sources)