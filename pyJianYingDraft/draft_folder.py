"""草稿文件夹管理器"""

import os
import shutil
import json
import time

from typing import List

from .script_file import Script_file

class Draft_folder:
    """管理一个文件夹及其内的一系列草稿"""

    folder_path: str
    """根路径"""

    def __init__(self, folder_path: str):
        """初始化草稿文件夹管理器

        Args:
            folder_path (`str`): 包含若干草稿的文件夹, 一般取剪映保存草稿的位置即可

        Raises:
            `FileNotFoundError`: 路径不存在
        """
        self.folder_path = folder_path

        if not os.path.exists(self.folder_path):
            raise FileNotFoundError(f"根文件夹 {self.folder_path} 不存在")

    def list_drafts(self) -> List[str]:
        """列出文件夹中所有草稿的名称

        注意: 本函数只是如实地列出子文件夹的名称, 并不检查它们是否符合草稿的格式
        """
        return [f for f in os.listdir(self.folder_path) if os.path.isdir(os.path.join(self.folder_path, f))]

    def remove(self, draft_name: str) -> None:
        """删除指定名称的草稿

        Args:
            draft_name (`str`): 草稿名称, 即相应文件夹名称

        Raises:
            `FileNotFoundError`: 对应的草稿不存在
            `Exception`: 更新草稿列表失败
        """
        draft_path = os.path.join(self.folder_path, draft_name)
        if not os.path.exists(draft_path):
            raise FileNotFoundError(f"草稿文件夹 {draft_name} 不存在")

        try:
            # 先从root_meta_info.json中移除草稿节点
            drafts_folder = self.get_drafts_folder()
            if drafts_folder:
                root_meta_file = os.path.join(drafts_folder, "root_meta_info.json")
                if os.path.exists(root_meta_file):
                    # 读取配置文件
                    with open(root_meta_file, 'r', encoding='utf-8') as f:
                        root_meta = json.load(f)

                    # 查找并移除目标草稿
                    if "all_draft_store" in root_meta:
                        draft_list = root_meta["all_draft_store"]
                        draft_list[:] = [draft for draft in draft_list if draft.get("draft_name") != draft_name]

                        # 写入临时文件
                        temp_file = root_meta_file + ".tmp"
                        with open(temp_file, 'w', encoding='utf-8') as f:
                            json.dump(root_meta, f, ensure_ascii=False, indent=2)

                        # 备份原文件
                        backup_file = root_meta_file + ".bak"
                        if os.path.exists(root_meta_file):
                            shutil.copy2(root_meta_file, backup_file)

                        # 替换原文件
                        os.replace(temp_file, root_meta_file)

            # 删除草稿文件夹
            shutil.rmtree(draft_path)
        except PermissionError:
            raise
        except Exception as e:
            raise Exception(f"删除草稿失败: {str(e)}")

    def inspect_material(self, draft_name: str) -> None:
        """输出指定名称草稿中的贴纸素材元数据

        Args:
            draft_name (`str`): 草稿名称, 即相应文件夹名称

        Raises:
            `FileNotFoundError`: 对应的草稿不存在
        """
        draft_path = os.path.join(self.folder_path, draft_name)
        if not os.path.exists(draft_path):
            raise FileNotFoundError(f"草稿文件夹 {draft_name} 不存在")

        script_file = self.load_template(draft_name)
        script_file.inspect_material()

    def load_template(self, draft_name: str) -> Script_file:
        """在文件夹中打开一个草稿作为模板, 并在其上进行编辑

        Args:
            draft_name (`str`): 草稿名称, 即相应文件夹名称

        Returns:
            `Script_file`: 以模板模式打开的草稿对象

        Raises:
            `FileNotFoundError`: 对应的草稿不存在
        """
        draft_path = os.path.join(self.folder_path, draft_name)
        if not os.path.exists(draft_path):
            raise FileNotFoundError(f"草稿文件夹 {draft_name} 不存在")
        print(os.path.join(draft_path, "draft_content.json"))
        return Script_file.load_template(os.path.join(draft_path, "draft_content.json"))

    def duplicate_as_template(self, template_name: str, new_draft_name: str, allow_replace: bool = False) -> Script_file:
        """复制一份给定的草稿, 并在复制出的新草稿上进行编辑

        Args:
            template_name (`str`): 原草稿名称
            new_draft_name (`str`): 新草稿名称
            allow_replace (`bool`, optional): 是否允许覆盖与`new_draft_name`重名的草稿. 默认为否.

        Returns:
            `Script_file`: 以模板模式打开的**复制后的**草稿对象

        Raises:
            `FileNotFoundError`: 原始草稿不存在
            `FileExistsError`: 已存在与`new_draft_name`重名的草稿, 但不允许覆盖.
        """
        template_path = os.path.join(self.folder_path, template_name)
        new_draft_path = os.path.join(self.folder_path, new_draft_name)
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"模板草稿 {template_name} 不存在")
        if os.path.exists(new_draft_path) and not allow_replace:
            raise FileExistsError(f"新草稿 {new_draft_name} 已存在且不允许覆盖")

        # 复制草稿文件夹
        shutil.copytree(template_path, new_draft_path, dirs_exist_ok=allow_replace)

        try:
            # 生成新的时间戳和草稿ID
            current_time = int(time.time() * 1000000)  # 生成16位时间戳
            new_draft_id = f"draft_{int(time.time() * 1000)}"

            # 更新草稿元数据文件
            draft_meta_file = os.path.join(new_draft_path, "draft_meta_info.json")
            if os.path.exists(draft_meta_file):
                with open(draft_meta_file, 'r', encoding='utf-8') as f:
                    draft_meta = json.load(f)

                draft_meta["draft_name"] = new_draft_name
                draft_meta["draft_id"] = new_draft_id
                draft_meta["tm_draft_modified"] = current_time
                draft_meta["tm_draft_create"] = current_time

                with open(draft_meta_file, 'w', encoding='utf-8') as f:
                    json.dump(draft_meta, f, ensure_ascii=False, indent=2)
            drafts_folder = self.get_drafts_folder()
            if drafts_folder:
                # 更新根目录的草稿列表文件
                root_meta_file = os.path.join(drafts_folder, "root_meta_info.json")
                if os.path.exists(root_meta_file):
                    with open(root_meta_file, 'r', encoding='utf-8') as f:
                        root_meta = json.load(f)

                    # 创建新草稿的元数据
                    new_draft_meta = {
                        "draft_name": new_draft_name,
                        "draft_id": new_draft_id,
                        "tm_draft_modified": current_time,
                        "tm_draft_create": current_time
                    }

                    # 将新草稿添加到列表开头
                    if "all_draft_store" in root_meta:
                        root_meta["all_draft_store"].insert(0, new_draft_meta)

                        # 写入临时文件
                        temp_file = root_meta_file + ".tmp"
                        with open(temp_file, 'w', encoding='utf-8') as f:
                            json.dump(root_meta, f, ensure_ascii=False, indent=2)

                        # 备份原文件
                        backup_file = root_meta_file + ".bak"
                        if os.path.exists(root_meta_file):
                            shutil.copy2(root_meta_file, backup_file)

                        # 替换原文件
                        os.replace(temp_file, root_meta_file)

        except Exception as e:
            # 如果更新元数据失败，清理新建的文件夹
            if os.path.exists(new_draft_path):
                shutil.rmtree(new_draft_path)
            raise Exception(f"更新草稿元数据失败: {str(e)}")

        # 打开草稿
        return self.load_template(new_draft_name)
    
    def get_drafts_folder(self):
        """获取剪映草稿文件夹路径"""
        appdata = os.getenv('APPDATA')
        if not appdata:
            return None
            
        # 修正路径拼接方式
        local_appdata = os.path.dirname(appdata)  # 获取上级目录
        local_appdata = os.path.join(local_appdata, "Local")  # 进入Local目录
        drafts_path = os.path.join(local_appdata, "JianyingPro", "User Data", "Projects", "com.lveditor.draft")
        
        print(f"草稿文件夹路径: {drafts_path}")
        print(f"文件夹是否存在: {os.path.exists(drafts_path)}")
        
        if not os.path.exists(drafts_path):
            # 尝试其他可能的路径
            alternative_paths = [
                os.path.join(appdata, "JianyingPro", "User Data", "Projects", "com.lveditor.draft"),
                os.path.join(appdata, "JianyingPro", "Projects", "com.lveditor.draft"),
                os.path.join(os.getenv('LOCALAPPDATA', ''), "JianyingPro", "User Data", "Projects", "com.lveditor.draft")
            ]
            
            for path in alternative_paths:
                print(f"尝试替代路径: {path}")
                if os.path.exists(path):
                    print(f"找到有效的草稿文件夹路径: {path}")
                    return path
                    
            return None
            
        return drafts_path