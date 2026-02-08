"""
阿里云 OSS 上传模块
用于替换 Google Drive 上传功能
"""

import oss2
import os
import logging
import configparser

logger = logging.getLogger(__name__)


class OSSUploader:
    """阿里云 OSS 上传器"""
    
    def __init__(self, config_path="config.ini"):
        self.config = configparser.ConfigParser()
        self.config.read(config_path, encoding='utf-8')
        
        # 读取OSS配置
        self.endpoint = self.config.get("OSS", "Endpoint", fallback="")
        self.bucket_name = self.config.get("OSS", "Bucket", fallback="")
        self.access_key_id = self.config.get("OSS", "AccessKeyId", fallback="")
        self.access_key_secret = self.config.get("OSS", "AccessKeySecret", fallback="")
        self.prefix = self.config.get("OSS", "Prefix", fallback="").strip("/")
        
        self.bucket = None
        self.auth = None
    
    def authenticate(self):
        """验证并连接到 OSS"""
        if not all([self.endpoint, self.bucket_name, self.access_key_id, self.access_key_secret]):
            logger.error("OSS配置不完整，请检查 config.ini 中的 [OSS] 配置节")
            return False
        
        try:
            # 确保endpoint格式正确
            endpoint = self.endpoint
            if not endpoint.startswith("http"):
                endpoint = f"https://{endpoint}"
            
            # 创建认证对象
            self.auth = oss2.Auth(self.access_key_id, self.access_key_secret)
            
            # 创建 Bucket 对象
            self.bucket = oss2.Bucket(self.auth, endpoint, self.bucket_name)
            
            # 测试连接 - 尝试获取bucket信息
            self.bucket.get_bucket_info()
            logger.info(f"OSS 连接成功: {self.bucket_name} @ {endpoint}")
            return True
            
        except oss2.exceptions.NoSuchBucket:
            logger.error(f"OSS Bucket 不存在: {self.bucket_name}")
            return False
        except oss2.exceptions.AccessDenied:
            logger.error("OSS 访问被拒绝，请检查 AccessKey 权限")
            return False
        except Exception as e:
            logger.error(f"OSS 连接失败: {e}")
            return False
    
    def create_folder(self, folder_name, parent_id=None):
        """
        创建文件夹（在OSS中实际上是返回一个前缀路径）
        为了与 DriveUploader 接口兼容，返回文件夹的 "ID"（实际上是路径前缀）
        """
        if not self.bucket:
            logger.error("OSS 未初始化，请先调用 authenticate()")
            return None
        
        # 清理文件夹名称
        folder_name = folder_name.replace("\\", "_").replace("/", "_")
        
        # 构建完整的前缀路径
        if self.prefix:
            folder_prefix = f"{self.prefix}/{folder_name}"
        else:
            folder_prefix = folder_name
        
        # OSS 不需要真正创建文件夹，只需返回前缀
        logger.info(f"OSS 文件夹前缀: {folder_prefix}")
        return folder_prefix
    
    def upload_file(self, file_path, folder_prefix):
        """
        上传文件到 OSS
        
        Args:
            file_path: 本地文件路径
            folder_prefix: 文件夹前缀（由 create_folder 返回）
            
        Returns:
            成功时返回包含文件信息的字典，失败返回 None
        """
        if not self.bucket:
            logger.error("OSS 未初始化，请先调用 authenticate()")
            return None
        
        if not os.path.exists(file_path):
            logger.error(f"文件不存在: {file_path}")
            return None
        
        file_name = os.path.basename(file_path)
        
        # 构建 OSS 对象键（路径）
        if folder_prefix:
            object_key = f"{folder_prefix}/{file_name}"
        else:
            object_key = file_name
        
        # 重试逻辑
        max_retries = 3
        for attempt in range(max_retries):
            try:
                if attempt > 0:
                    logger.info(f"重试上传 ({attempt + 1}/{max_retries})...")
                    import time
                    time.sleep(2 * attempt)
                
                # 上传文件
                result = self.bucket.put_object_from_file(object_key, file_path)
                
                if result.status == 200:
                    logger.info(f"OSS 上传成功: {object_key}")
                    
                    # 返回与 DriveUploader 兼容的格式
                    return {
                        'id': object_key,
                        'name': file_name,
                        'object_key': object_key
                    }
                else:
                    logger.warning(f"上传返回状态码: {result.status}")
                    
            except Exception as e:
                logger.warning(f"上传尝试 {attempt + 1} 失败: {e}")
                if attempt == max_retries - 1:
                    logger.error(f"上传最终失败: {file_path}")
                    return None
        
        return None
    
    def get_direct_link(self, object_key):
        """
        获取文件的公开访问链接
        
        注意：需要确保 Bucket 设置了公共读权限，或者文件设置了公共读权限
        
        Args:
            object_key: OSS 对象键（由 upload_file 返回的 id）
            
        Returns:
            公开访问的 URL
        """
        # 构建公开访问 URL
        # 格式: https://{bucket}.{endpoint}/{object_key}
        endpoint = self.endpoint
        if endpoint.startswith("https://"):
            endpoint = endpoint[8:]
        elif endpoint.startswith("http://"):
            endpoint = endpoint[7:]
        
        url = f"https://{self.bucket_name}.{endpoint}/{object_key}"
        return url
    
    def set_object_acl_public(self, object_key):
        """
        设置单个对象为公共读权限
        
        Args:
            object_key: OSS 对象键
            
        Returns:
            成功返回 True，失败返回 False
        """
        if not self.bucket:
            return False
        
        try:
            self.bucket.put_object_acl(object_key, oss2.OBJECT_ACL_PUBLIC_READ)
            logger.info(f"已设置对象为公共读: {object_key}")
            return True
        except Exception as e:
            logger.error(f"设置对象权限失败: {e}")
            return False
