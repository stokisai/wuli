"""
ComfyUI API Client for Image-to-Image Workflow
支持动态GPU端口配置，与Excel任务系统集成
"""
import os
import json
import time
import uuid
import random
import logging
import requests
from pathlib import Path
from urllib.parse import urljoin, urlparse

# 抑制 HTTPS 自签名证书的 InsecureRequestWarning
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

logger = logging.getLogger(__name__)


class ComfyUIClient:
    """ComfyUI API客户端，支持图生图工作流"""
    
    @classmethod
    def from_url(cls, url: str, timeout: int = 300):
        """
        从完整URL创建客户端
        
        Args:
            url: 完整的ComfyUI URL，如 'https://wp08.unicorn.org.cn:37062/?__theme=dark'
            timeout: 处理超时时间
            
        Returns:
            ComfyUIClient实例
        """
        parsed = urlparse(url)
        host = parsed.hostname or "127.0.0.1"
        port = parsed.port or (443 if parsed.scheme == 'https' else 8188)
        scheme = parsed.scheme or 'http'
        
        client = cls(port=port, host=host, timeout=timeout, scheme=scheme)
        return client
    
    def __init__(self, port: int = 8188, host: str = "127.0.0.1", timeout: int = 300, scheme: str = "http"):
        """
        初始化ComfyUI客户端
        
        Args:
            port: ComfyUI服务端口（支持动态配置）
            host: ComfyUI服务主机
            timeout: 处理超时时间（秒）
            scheme: 协议 (http/https)
        """
        self.host = host
        self.port = port
        self.timeout = timeout
        self.scheme = scheme
        self.base_url = f"{scheme}://{host}:{port}"
        self.client_id = str(uuid.uuid4())
        
        # 加载工作流模板
        self.workflow_path = os.path.join(os.path.dirname(__file__), "workflow_i2i.json")
        self.workflow_template = None
        
        logger.info(f"ComfyUI客户端初始化: {self.base_url}")
    
    def check_connection(self) -> bool:
        """检查与ComfyUI服务器的连接，HTTPS失败时自动回退HTTP"""
        try:
            response = requests.get(f"{self.base_url}/system_stats", timeout=5, verify=False)
            return response.status_code == 200
        except requests.exceptions.SSLError:
            # HTTPS 握手失败，尝试回退到 HTTP
            if self.scheme == "https":
                fallback_url = f"http://{self.host}:{self.port}"
                logger.info(f"HTTPS连接失败，尝试HTTP回退: {fallback_url}")
                try:
                    response = requests.get(f"{fallback_url}/system_stats", timeout=5)
                    if response.status_code == 200:
                        self.scheme = "http"
                        self.base_url = fallback_url
                        logger.info(f"HTTP回退成功，已切换到: {self.base_url}")
                        return True
                except Exception as e2:
                    logger.error(f"HTTP回退也失败: {e2}")
            return False
        except Exception as e:
            logger.error(f"ComfyUI连接检查失败: {e}")
            return False
    
    def load_workflow(self, workflow_path: str = None) -> dict:
        """加载工作流JSON"""
        path = workflow_path or self.workflow_path
        try:
            with open(path, 'r', encoding='utf-8') as f:
                self.workflow_template = json.load(f)
            logger.info(f"工作流加载成功: {path}")
            return self.workflow_template
        except Exception as e:
            logger.error(f"工作流加载失败: {e}")
            return None
    
    def upload_image(self, image_path: str, subfolder: str = "", overwrite: bool = True, target_type: str = "input") -> str:
        """
        上传图片到ComfyUI服务器

        Args:
            image_path: 本地图片路径
            subfolder: 服务器子文件夹
            overwrite: 是否覆盖同名文件
            target_type: 上传目标目录 ("input", "output", "temp")

        Returns:
            服务器端文件名，失败返回None
        """
        if not os.path.exists(image_path):
            logger.error(f"图片不存在: {image_path}")
            return None

        try:
            filename = os.path.basename(image_path)

            with open(image_path, 'rb') as f:
                files = {
                    'image': (filename, f, 'image/png')
                }
                data = {
                    'subfolder': subfolder,
                    'overwrite': str(overwrite).lower(),
                    'type': target_type,
                }

                response = requests.post(
                    f"{self.base_url}/upload/image",
                    files=files,
                    data=data,
                    timeout=30,
                    verify=False
                )

                if response.status_code == 200:
                    result = response.json()
                    server_filename = result.get('name', filename)
                    logger.info(f"图片上传成功({target_type}): {filename} -> {server_filename}")
                    return server_filename
                else:
                    logger.error(f"图片上传失败: {response.status_code} - {response.text}")
                    return None

        except Exception as e:
            logger.error(f"图片上传异常: {e}")
            return None
    
    def _is_api_format(self, workflow: dict) -> bool:
        """检测工作流是否已经是API格式（而非Web/原始格式）"""
        if "nodes" in workflow and isinstance(workflow["nodes"], list):
            return False
        for key, value in workflow.items():
            if isinstance(value, dict) and "class_type" in value:
                return True
        return False

    def _prepare_api_format_workflow(self, workflow: dict, source_image_name: str, prompt_text: str = None) -> dict:
        """处理已经是API格式的工作流，设置源图片和提示词"""
        api_workflow = json.loads(json.dumps(workflow))

        for node_id, node in api_workflow.items():
            class_type = node.get("class_type", "")
            inputs = node.get("inputs", {})

            if class_type == "LoadImage":
                inputs["image"] = source_image_name
                logger.info(f"设置源图片节点{node_id}(LoadImage): {source_image_name}")
            elif class_type == "LoadImageOutput":
                # 将 LoadImageOutput 转换为 LoadImage，从 input 目录加载上传的图片
                node["class_type"] = "LoadImage"
                inputs["image"] = source_image_name
                # 移除 LoadImageOutput 特有的字段
                for key in ("refresh", "upload_to_output"):
                    inputs.pop(key, None)
                logger.info(f"节点{node_id}: LoadImageOutput -> LoadImage, 图片: {source_image_name}")

            # KSampler / KSamplerAdvanced: 每次生成使用随机种子
            # ComfyUI 网页端 control_after_generate 默认 randomize，但该设置不保存到 API JSON
            if class_type in ("KSampler", "KSamplerAdvanced") and "seed" in inputs:
                new_seed = random.randint(0, 2**53 - 1)
                logger.info(f"节点{node_id}({class_type}): seed {inputs['seed']} -> {new_seed}")
                inputs["seed"] = new_seed
            elif class_type == "RandomNoise" and "noise_seed" in inputs:
                new_seed = random.randint(0, 2**53 - 1)
                logger.info(f"节点{node_id}(RandomNoise): noise_seed {inputs['noise_seed']} -> {new_seed}")
                inputs["noise_seed"] = new_seed

            if class_type == "DeepTranslatorTextNode" and prompt_text:
                inputs["text"] = prompt_text
                logger.info(f"覆盖提示词节点{node_id}: {prompt_text[:50]}...")

            node.pop("_meta", None)
            for inp_key in list(inputs.keys()):
                if inp_key == "speak_and_recognation":
                    del inputs[inp_key]

        logger.info(f"API格式工作流准备完成，包含 {len(api_workflow)} 个节点")
        return api_workflow

    def prepare_workflow(self, source_image_name: str, prompt_text: str = None) -> dict:
        """
        准备工作流配置，设置输入图片
        将原始工作流格式转换为ComfyUI API格式
        
        Args:
            source_image_name: 服务器端的源图片文件名
            prompt_text: 可选的提示词覆盖
            
        Returns:
            准备好的工作流API格式
        """
        if not self.workflow_template:
            self.load_workflow()
        
        if not self.workflow_template:
            logger.error("无法加载工作流模板")
            return None
        
        # 深拷贝工作流
        workflow = json.loads(json.dumps(self.workflow_template))

        # 检测是否已经是API格式
        if self._is_api_format(workflow):
            logger.info("检测到API格式工作流，跳过转换")
            return self._prepare_api_format_workflow(workflow, source_image_name, prompt_text)

        # 构建链接映射: link_id -> (source_node_id, output_slot)
        links_map = {}
        for link in workflow.get("links", []):
            # link格式: [link_id, source_node, source_slot, target_node, target_slot, type]
            link_id = link[0]
            source_node = str(link[1])
            source_slot = link[2]
            links_map[link_id] = (source_node, source_slot)
        
        # 节点类型到widget参数名的映射
        widget_mappings = {
            "VAELoader": ["vae_name"],
            "UNETLoader": ["unet_name", "weight_dtype"],
            "DualCLIPLoader": ["clip_name1", "clip_name2", "type", "device"],
            "LoadImage": ["image", "upload"],  # 新版工作流使用LoadImage
            "LoadImageOutput": ["image", "upload_to_output", "refresh", "upload"],  # 保留旧版兼容
            "DeepTranslatorTextNode": ["from_translate", "to_translate", "add_proxies", "proxies", 
                                       "auth_data", "service", "text"],
            "KSamplerSelect": ["sampler_name"],
            "RandomNoise": ["noise_seed", "control_after_generate"],
            "BasicScheduler": ["scheduler", "steps", "denoise"],
            "EmptySD3LatentImage": ["width", "height", "batch_size"],
            "FluxGuidance": ["guidance"],
            "CLIPTextEncode": ["text"],
            "SaveImage": ["filename_prefix"],
            "easy hiresFix": ["model_name", "rescale_after_model", "rescale_method", "rescale", 
                             "percent", "width", "height", "longer_side", "crop", "image_output", 
                             "link_id", "save_prefix"],  # 4x放大节点
        }
        
        # 转换为API格式
        api_workflow = {}
        
        # 首先收集所有禁用节点的ID
        disabled_nodes = set()
        for node in workflow.get("nodes", []):
            if node.get("mode") == 4:
                disabled_nodes.add(str(node["id"]))
        logger.debug(f"禁用的节点: {disabled_nodes}")
        
        for node in workflow.get("nodes", []):
            node_id = str(node["id"])
            node_type = node["type"]
            
            # 跳过禁用的节点 (mode == 4)
            if node.get("mode") == 4:
                continue
            
            # 构建API格式的节点
            api_node = {
                "class_type": node_type,
                "inputs": {}
            }
            
            # 1. 处理输入连接 - 跳过来自禁用节点的连接
            for inp in node.get("inputs", []):
                inp_name = inp.get("name")
                link_id = inp.get("link")
                
                if link_id and link_id in links_map:
                    source_node, source_slot = links_map[link_id]
                    # 检查源节点是否被禁用
                    if source_node in disabled_nodes:
                        logger.debug(f"跳过来自禁用节点{source_node}的连接: {inp_name}")
                        continue
                    api_node["inputs"][inp_name] = [source_node, source_slot]
            
            # 2. 处理widgets_values
            widgets_values = node.get("widgets_values", [])
            param_names = widget_mappings.get(node_type, [])
            
            for i, value in enumerate(widgets_values):
                if i < len(param_names):
                    param_name = param_names[i]
                    # 跳过已经通过连接设置的参数
                    if param_name not in api_node["inputs"]:
                        api_node["inputs"][param_name] = value
            
            # 3. 特殊处理: LoadImage节点 (node 8) - 设置源图片 (新版工作流)
            if node_id == "8" and node_type == "LoadImage":
                api_node["inputs"]["image"] = source_image_name
                logger.info(f"设置源图片节点8: {source_image_name}")
            
            # 兼容旧版: LoadImageOutput节点 (node 142) -> 转换为 LoadImage
            if node_id == "142" and node_type == "LoadImageOutput":
                api_node["class_type"] = "LoadImage"
                api_node["inputs"]["image"] = source_image_name
                # 移除 LoadImageOutput 特有的字段
                for key in ("refresh", "upload_to_output"):
                    api_node["inputs"].pop(key, None)
                logger.info(f"节点142: LoadImageOutput -> LoadImage, 图片: {source_image_name}")
            
            # 4. 特殊处理: DeepTranslatorTextNode节点 (node 20 新版 / node 190 旧版) - 可选覆盖提示词
            if node_type == "DeepTranslatorTextNode" and node_id in ("20", "190"):
                if prompt_text:
                    api_node["inputs"]["text"] = prompt_text
                else:
                    # 使用工作流中的默认提示词
                    if len(widgets_values) > 6:
                        api_node["inputs"]["text"] = widgets_values[6]
                
                # 设置必需的参数
                api_node["inputs"]["from_translate"] = widgets_values[0] if widgets_values else "auto"
                api_node["inputs"]["to_translate"] = widgets_values[1] if len(widgets_values) > 1 else "english"
                api_node["inputs"]["add_proxies"] = widgets_values[2] if len(widgets_values) > 2 else False
                api_node["inputs"]["proxies"] = widgets_values[3] if len(widgets_values) > 3 else ""
                api_node["inputs"]["auth_data"] = widgets_values[4] if len(widgets_values) > 4 else ""
                api_node["inputs"]["service"] = widgets_values[5] if len(widgets_values) > 5 else "GoogleTranslator"
            
            api_workflow[node_id] = api_node
        
        logger.info(f"工作流转换完成，包含 {len(api_workflow)} 个节点")
        return api_workflow
    
    def queue_prompt(self, workflow: dict) -> str:
        """
        提交工作流到队列执行
        
        Args:
            workflow: API格式的工作流配置
            
        Returns:
            prompt_id，失败返回None
        """
        try:
            payload = {
                "prompt": workflow,
                "client_id": self.client_id
            }
            
            response = requests.post(
                f"{self.base_url}/prompt",
                json=payload,
                timeout=30,
                verify=False
            )
            
            if response.status_code == 200:
                result = response.json()
                prompt_id = result.get("prompt_id")
                logger.info(f"工作流已提交: {prompt_id}")
                return prompt_id
            else:
                logger.error(f"工作流提交失败: {response.status_code} - {response.text}")
                return None
                
        except Exception as e:
            logger.error(f"工作流提交异常: {e}")
            return None
    
    def get_history(self, prompt_id: str) -> dict:
        """获取工作流执行历史"""
        try:
            response = requests.get(
                f"{self.base_url}/history/{prompt_id}",
                timeout=10,
                verify=False
            )
            if response.status_code == 200:
                return response.json()
            return {}
        except Exception as e:
            logger.debug(f"获取历史失败: {e}")
            return {}
    
    def wait_for_completion(self, prompt_id: str) -> dict:
        """
        等待工作流完成
        
        Args:
            prompt_id: 工作流ID
            
        Returns:
            输出信息字典，包含生成的图片信息
        """
        start_time = time.time()
        
        while time.time() - start_time < self.timeout:
            history = self.get_history(prompt_id)
            
            if prompt_id in history:
                outputs = history[prompt_id].get("outputs", {})
                status = history[prompt_id].get("status", {})
                
                if status.get("completed", False) or outputs:
                    logger.info(f"工作流完成: {prompt_id}")
                    return outputs
                
                if status.get("status_str") == "error":
                    logger.error(f"工作流执行错误: {status}")
                    return None
            
            time.sleep(1)
            elapsed = int(time.time() - start_time)
            if elapsed % 10 == 0:
                logger.info(f"等待工作流完成... {elapsed}s")
        
        logger.error(f"工作流超时: {self.timeout}s")
        return None
    
    def download_image(self, filename: str, subfolder: str, output_path: str, img_type: str = "output") -> bool:
        """
        从ComfyUI下载生成的图片
        
        Args:
            filename: 服务器端文件名
            subfolder: 子文件夹
            output_path: 本地保存路径
            img_type: 图片类型 (output/input/temp)
            
        Returns:
            是否成功
        """
        try:
            params = {
                "filename": filename,
                "subfolder": subfolder,
                "type": img_type
            }
            
            response = requests.get(
                f"{self.base_url}/view",
                params=params,
                timeout=60,
                verify=False
            )
            
            if response.status_code == 200:
                # 确保输出目录存在
                os.makedirs(os.path.dirname(output_path), exist_ok=True)
                
                with open(output_path, 'wb') as f:
                    f.write(response.content)

                logger.info(f"图片下载成功: {filename} -> {output_path}")

                return True
            else:
                logger.error(f"图片下载失败: {response.status_code}")
                return False
                
        except Exception as e:
            logger.error(f"图片下载异常: {e}")
            return False
    

    def process_image(self, source_path: str, output_path: str, prompt_text: str = None, max_retries: int = 3) -> bool:
        """
        完整的图生图处理流程，带自动重试
        
        Args:
            source_path: 源图片本地路径
            output_path: 输出图片保存路径
            prompt_text: 可选的提示词
            max_retries: 最大重试次数
            
        Returns:
            是否成功
        """
        for attempt in range(max_retries):
            if attempt > 0:
                wait_time = 5 * attempt  # 递增等待时间
                logger.info(f"第 {attempt + 1}/{max_retries} 次重试，等待 {wait_time} 秒...")
                time.sleep(wait_time)
            
            result = self._process_image_once(source_path, output_path, prompt_text)
            if result:
                return True
            
            if attempt < max_retries - 1:
                logger.warning(f"处理失败，将进行重试...")
        
        logger.error(f"处理失败，已达到最大重试次数 ({max_retries})")
        return False
    
    def _process_image_once(self, source_path: str, output_path: str, prompt_text: str = None) -> bool:
        """单次图生图处理尝试"""
        logger.info(f"开始图生图处理: {source_path}")

        # 1. 检查连接
        if not self.check_connection():
            logger.error(f"无法连接到ComfyUI服务器: {self.base_url}")
            return False

        # 2. 上传图片到 input 目录（LoadImageOutput 会在 prepare 阶段被转换为 LoadImage）
        server_filename = self.upload_image(source_path)
        if not server_filename:
            return False
        
        # 3. 准备工作流
        workflow = self.prepare_workflow(server_filename, prompt_text)
        if not workflow:
            return False

        # 4. 提交执行
        prompt_id = self.queue_prompt(workflow)
        if not prompt_id:
            return False
        
        # 5. 等待完成
        outputs = self.wait_for_completion(prompt_id)
        if not outputs:
            return False
        
        # 6. 下载结果图片
        # 查找输出图片 - 通常在PreviewImage或SaveImage节点
        logger.info(f"工作流输出节点: {list(outputs.keys())}")
        
        node_class_by_id = {}
        if isinstance(workflow, dict):
            for wf_node_id, wf_node in workflow.items():
                if isinstance(wf_node, dict):
                    node_class_by_id[str(wf_node_id)] = wf_node.get("class_type", "")

        candidates = []
        type_priority = {"output": 0, "temp": 1, "input": 2}

        # 同类型图片时，优先真实处理节点，避免优先取 PreviewImage 包装节点导致结果不稳定
        def node_rank(node_class: str) -> int:
            if node_class == "SaveImage":
                return 0
            if node_class == "PreviewImage":
                return 3
            if node_class:
                return 1
            return 2

        for node_id, node_output in outputs.items():
            logger.debug(f"节点 {node_id} 输出: {node_output}")
            images = node_output.get("images", [])
            for img_info in images:
                filename = img_info.get("filename")
                subfolder = img_info.get("subfolder", "")
                img_type = img_info.get("type", "output")
                n_class = node_class_by_id.get(str(node_id), "")

                logger.info(f"找到图片: {filename}, 类型: {img_type}, 节点: {node_id}/{n_class}, 子目录: {subfolder}")
                if not filename:
                    continue

                try:
                    node_order = -int(node_id)
                except (TypeError, ValueError):
                    node_order = 0

                candidates.append({
                    "filename": filename,
                    "subfolder": subfolder,
                    "img_type": img_type,
                    "node_id": str(node_id),
                    "node_class": n_class,
                    "priority": type_priority.get(img_type, 9),
                    "class_rank": node_rank(n_class),
                    "node_order": node_order,
                })

        if not candidates:
            logger.error("未找到可下载的输出图片")
            return False

        candidates.sort(key=lambda item: (item["priority"], item["class_rank"], item["node_order"]))

        for candidate in candidates:
            logger.info(
                "尝试下载候选图片: node=%s, type=%s, file=%s",
                candidate["node_id"],
                candidate["img_type"],
                candidate["filename"],
            )
            success = self.download_image(
                candidate["filename"],
                candidate["subfolder"],
                output_path,
                candidate["img_type"],
            )
            if success:
                logger.info(f"图生图处理完成: {output_path}")
                return True

        logger.error("候选图片下载失败")
        return False


# 便捷函数
def process_with_comfyui(source_path: str, output_path: str, port: int = 8188, 
                         host: str = "127.0.0.1", prompt_text: str = None) -> bool:
    """
    便捷的图生图处理函数
    
    Args:
        source_path: 源图片路径
        output_path: 输出图片路径
        port: ComfyUI端口
        host: ComfyUI主机
        prompt_text: 可选提示词
        
    Returns:
        是否成功
    """
    client = ComfyUIClient(port=port, host=host)
    return client.process_image(source_path, output_path, prompt_text)


if __name__ == "__main__":
    # 测试代码
    logging.basicConfig(level=logging.INFO)
    
    client = ComfyUIClient(port=8188)
    
    if client.check_connection():
        print("ComfyUI连接成功!")
    else:
        print("ComfyUI连接失败，请检查服务是否运行")

