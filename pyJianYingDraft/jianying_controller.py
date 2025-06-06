"""剪映自动化控制，主要与自动导出有关"""

import time
import shutil
import uiautomation as uia
import os

from enum import Enum
from typing import Optional, Literal, Callable

from . import exceptions
from .exceptions import AutomationError

class Export_resolution(Enum):
    """导出分辨率"""
    DEFAULT = "默认"
    RES_8K = "8K"
    RES_4K = "4K"
    RES_2K = "2K"
    RES_1080P = "1080P"
    RES_720P = "720P"
    RES_480P = "480P"

    @classmethod
    def from_value(cls, value: str):
        """根据值获取对应的枚举
        Args:
            value: 分辨率值，如 "1080P"
        Returns:
            Export_resolution: 对应的枚举值，如果未找到则返回默认值 RES_1080P
        """
        try:
            return next(item for item in cls if item.value == value)
        except StopIteration:
            return cls.RES_1080P

class Export_framerate(Enum):
    """导出帧率"""
    DEFAULT = "默认"
    FR_24 = "24fps"
    FR_25 = "25fps"
    FR_30 = "30fps"
    FR_50 = "50fps"
    FR_60 = "60fps"

    @classmethod
    def from_value(cls, value: str):
        """根据值获取对应的枚举
        Args:
            value: 帧率值，如 "30fps"
        Returns:
            Export_framerate: 对应的枚举值，如果未找到则返回默认值 FR_30
        """
        try:
            return next(item for item in cls if item.value == value)
        except StopIteration:
            return cls.FR_30

class ControlFinder:
    """控件查找器，封装部分与控件查找相关的逻辑"""

    @staticmethod
    def desc_matcher(target_desc: str, depth: int = 2, exact: bool = False) -> Callable[[uia.Control, int], bool]:
        """根据full_description查找控件的匹配器"""
        target_desc = target_desc.lower()
        def matcher(control: uia.Control, _depth: int) -> bool:
            if _depth != depth:
                return False
            full_desc: str = control.GetPropertyValue(30159).lower()
            return (target_desc == full_desc) if exact else (target_desc in full_desc)
        return matcher

    @staticmethod
    def class_name_matcher(class_name: str, depth: int = 1, exact: bool = False) -> Callable[[uia.Control, int], bool]:
        """根据ClassName查找控件的匹配器"""
        class_name = class_name.lower()
        def matcher(control: uia.Control, _depth: int) -> bool:
            if _depth != depth:
                return False
            curr_class_name: str = control.ClassName.lower()
            return (class_name == curr_class_name) if exact else (class_name in curr_class_name)
        return matcher

class Jianying_controller:
    """剪映控制器"""

    app: uia.WindowControl
    """剪映窗口"""
    app_status: Literal["home", "edit", "pre_export"]

    def __init__(self,set_top=True):
        """初始化剪映控制器, 此时剪映应该处于目录页"""
        self.get_window(set_top)

    def export_draft(self, draft_name: str, output_path: Optional[str] = None, *,
                     resolution: Optional[Export_resolution] = None,
                     framerate: Optional[Export_framerate] = None,
                     timeout: float = 1200,juming= None) -> str | None:
        """导出指定的剪映草稿, **目前仅支持剪映6及以下版本**

        **注意: 需要确认有导出草稿的权限(不使用VIP功能或已开通VIP), 否则可能陷入死循环**

        Args:
            draft_name (`str`): 要导出的剪映草稿名称
            output_path (`str`, optional): 导出路径, 支持指向文件夹或直接指向文件, 不指定则使用剪映默认路径.
            resolution (`Export_resolution`, optional): 导出分辨率, 默认不改变剪映导出窗口中的设置.
            framerate (`Export_framerate`, optional): 导出帧率, 默认不改变剪映导出窗口中的设置.
            timeout (`float`, optional): 导出超时时间(秒), 默认为20分钟.

        Raises:
            `DraftNotFound`: 未找到指定名称的剪映草稿
            `AutomationError`: 剪映操作失败
        """
        # logger.info(f"开始导出 {draft_name} 至 {output_path}")
        self.get_window()
        self.switch_to_home()

        # 点击对应草稿
        draft_name_text = self.app.TextControl(
            searchDepth=2,
            Compare=ControlFinder.desc_matcher(f"HomePageDraftTitle:{draft_name}", exact=True)
        )
        if not draft_name_text.Exists(0):
            raise exceptions.DraftNotFound(f"未找到名为{draft_name}的剪映草稿")
        draft_btn = draft_name_text.GetParentControl()
        assert draft_btn is not None
        draft_btn.Click(simulateMove=False)
        time.sleep(3)
        # self.close_relink_window()
        self.get_window()

        start_time = time.time()
        while True:
            self.send_keys('{Ctrl}e',1)
            self.get_window()
            # 获取原始导出路径（带后缀名）
            export_path_sib = self.app.TextControl(searchDepth=2, Compare=ControlFinder.desc_matcher("ExportPath"))
            if export_path_sib.Exists(0):
                break
            if time.time() - start_time > 20:
                raise AutomationError(f"未找到导出路径，超时时间：{20}秒")

        export_path_text = export_path_sib.GetSiblingControl(lambda ctrl: True)
        assert export_path_text is not None
        export_path = export_path_text.GetPropertyValue(30159)

        # 设置分辨率
        if resolution is not None:
            setting_group = self.app.GroupControl(searchDepth=1,
                                                  Compare=ControlFinder.class_name_matcher("PanelSettingsGroup_QMLTYPE"))
            if not setting_group.Exists(0):
                raise AutomationError("未找到导出设置组")
            resolution_btn = setting_group.TextControl(searchDepth=2, Compare=ControlFinder.desc_matcher("ExportSharpnessInput"))
            if not resolution_btn.Exists(0.5):
                raise AutomationError("未找到导出分辨率下拉框")
            resolution_btn.Click(simulateMove=False)
            time.sleep(0.5)
            resolution_item = self.app.TextControl(
                searchDepth=2, Compare=ControlFinder.desc_matcher(resolution.value)
            )
            if not resolution_item.Exists(0.5):
                raise AutomationError(f"未找到{resolution.value}分辨率选项")
            resolution_item.Click(simulateMove=False)
            time.sleep(0.5)

        # 设置帧率
        if framerate is not None:
            setting_group = self.app.GroupControl(searchDepth=1,
                                                  Compare=ControlFinder.class_name_matcher("PanelSettingsGroup_QMLTYPE"))
            if not setting_group.Exists(0):
                raise AutomationError("未找到导出设置组")
            framerate_btn = setting_group.TextControl(searchDepth=2, Compare=ControlFinder.desc_matcher("FrameRateInput"))
            if not framerate_btn.Exists(0.5):
                raise AutomationError("未找到导出帧率下拉框")
            framerate_btn.Click(simulateMove=False)
            time.sleep(0.5)
            framerate_item = self.app.TextControl(
                searchDepth=2, Compare=ControlFinder.desc_matcher(framerate.value)
            )
            if not framerate_item.Exists(0.5):
                raise AutomationError(f"未找到{framerate.value}帧率选项")
            framerate_item.Click(simulateMove=False)
            time.sleep(0.5)


        # 点击导出
        export_btn = self.app.TextControl(searchDepth=2, Compare=ControlFinder.desc_matcher("ExportOkBtn", exact=True))
        if not export_btn.Exists(0):
            raise AutomationError("未在导出窗口中找到导出按钮")
        # export_btn.Click(simulateMove=False)
        start_time = time.time()
        while True:
            try:
                export_btn.Click(simulateMove=False)
                export_btn = self.app.TextControl(searchDepth=2, Compare=ControlFinder.desc_matcher("ExportOkBtn", exact=True))
                if not export_btn.Exists(0):
                    break
            except:
                pass
            if time.time() - start_time > 5:  # 10秒超时
                raise AutomationError("未在导出窗口中找到导出按钮")
            pass
            time.sleep(0.5)  # 添加短暂延迟，避免过于频繁的尝试
        # 等待导出完成
        st = time.time()
        while True:
            # self.get_window()
            if self.app_status != "pre_export": continue
            has_close = False
            succeed_close_btn = self.app.TextControl(searchDepth=2, Compare=ControlFinder.desc_matcher("ExportSucceedCloseBtn"))
            if succeed_close_btn.Exists(0):
                start_time = time.time()
                while True:
                    try:
                        self.get_window()
                        succeed_close_btn.Click(simulateMove=False)
                        self.switch_to_home()
                        has_close = True
                        break
                    except:
                        if time.time() - start_time > 10:  # 10秒超时
                            raise AutomationError("关闭导出窗口超时")
                        pass
                    time.sleep(0.5)  # 添加短暂延迟，避免过于频繁的尝试

            if time.time() - st > timeout:
                raise AutomationError("导出超时, 时限为%d秒" % timeout)
            if has_close:
                break
            time.sleep(1)
        # 移动文件到目标目录
        if output_path is not None:
            
            # 获取导出文件的文件名
            export_filename = os.path.basename(export_path)
            # logger.info(os.path.isdir(export_path))
            # 如果output_path是目录，则拼接完整路径
            if os.path.isdir(export_path):
                # 在目录下查找视频文件
                video_files = [f for f in os.listdir(export_path) if f.endswith(('.mp4', '.mov', '.avi', '.wmv'))]
                if video_files:
                    video_name = video_files[0]  # 获取第一个视频文件的完整名称（包含后缀）
                    final_path = os.path.join(output_path, export_filename, video_name)
                    
                    # 检查草稿目录下是否有video_cover.jpg，如果有则一起复制
                    cover_file = os.path.join(export_path, "video_cover.jpg")
                    if os.path.exists(cover_file):
                        target_dir = os.path.join(output_path, export_filename)
                        if not os.path.exists(target_dir):
                            os.makedirs(target_dir)
                        # 使用视频文件名来命名封面图
                        video_name_without_ext = os.path.splitext(video_name)[0]
                        target_cover = os.path.join(target_dir, f"{video_name_without_ext}_cover.jpg")
                        shutil.copy2(cover_file, target_cover)
            else:
                # logger.info(f'output_path {output_path}')
                output_path = os.path.join(output_path, juming)
                final_path = os.path.join(output_path, export_filename)
                if not os.path.exists(output_path):
                    os.makedirs(output_path)
                
                # 检查草稿目录下是否有video_cover.jpg，如果有则一起复制
                draft_dir = os.path.dirname(export_path)
                cover_file = os.path.join(draft_dir, "video_cover.jpg")
                if os.path.exists(cover_file):
                    # 使用导出文件名来命名封面图
                    export_name_without_ext = os.path.splitext(export_filename)[0]
                    target_cover = os.path.join(output_path, f"{export_name_without_ext}_cover.jpg")
                    shutil.copy2(cover_file, target_cover)
            
            # 移动文件
            shutil.move(export_path, output_path)
            return final_path
        return None

    def close_relink_window(self):
        windows = uia.GetRootControl().GetChildren()
        def search_relink_window(control):
            """递归搜索链接媒体窗口"""
            if (control.Name == "链接媒体" and
                    "RelinkMediaView" in control.ClassName and
                    control.ControlTypeName == "WindowControl" and
                    control.IsEnabled and not control.IsOffscreen):
                return control

            for child in control.GetChildren():
                result = search_relink_window(child)
                if result:
                    return result
            return None
        # 首先在主窗口中查找
        main_window = None
        for window in windows:
            if window.Name == "剪映专业版" and "MainWindow" in window.ClassName:
                main_window = window
                break
        if main_window:
            relink_window = search_relink_window(main_window)
            if relink_window:
                # 重试几次确保窗口关闭
                retry_count = 3
                while retry_count > 0:
                    self.get_window()
                    self.send_keys('{Esc}', 1)
                    # 重新检查窗口是否还存在
                    if not search_relink_window(main_window):
                        raise AutomationError("有素材未找到，结束导出")
                    retry_count -= 1

    def switch_to_home(self) -> None:
        """切换到剪映主页"""
        if self.app_status == "home":
            return
        # if self.app_status != "edit":
        self.send_keys('{Esc}', 3)
        # close_btn = self.app.GroupControl(searchDepth=1, ClassName="TitleBarButton", foundIndex=3)
        # close_btn.Click(simulateMove=False)
        self.send_keys('{Ctrl}{Alt}q',3)
        # time.sleep(1.5)
        self.get_window()

    def send_keys(self,key,count):
        for _ in range(count):
            uia.SendKeys(key)
            time.sleep(0.5)
    def get_window(self,set_top=True) -> None:
        """寻找剪映窗口并置顶"""
        if hasattr(self, "app") and self.app.Exists(0):
            self.app.SetTopmost(False)

        self.app = uia.WindowControl(searchDepth=1, Compare=self.__jianying_window_cmp)
        if not self.app.Exists(0):
            raise AutomationError("剪映窗口未找到")

        # 寻找可能存在的导出窗口
        export_window = self.app.WindowControl(searchDepth=1, Name="导出")
        if export_window.Exists(0):
            self.app = export_window
            self.app_status = "pre_export"
        if set_top:
            self.app.SetActive()
            # self.app.SetTopmost()

    def __jianying_window_cmp(self, control: uia.WindowControl, depth: int) -> bool:
        if control.Name != "剪映专业版":
            return False
        if "HomePage".lower() in control.ClassName.lower():
            self.app_status = "home"
            return True
        if "MainWindow".lower() in control.ClassName.lower():
            self.app_status = "edit"
            return True
        return False

    def click_draft(self, draft_name):
        draft_name_text = self.app.TextControl(searchDepth=2,
                                               Compare=ControlFinder.desc_matcher(f"HomePageDraftTitle:{draft_name}", exact=True))
        if not draft_name_text.Exists(0):
            raise exceptions.DraftNotFound(f"未找到名为{draft_name}的剪映草稿")
        draft_btn = draft_name_text.GetParentControl()
        assert draft_btn is not None
        draft_btn.Click(simulateMove=False)
        time.sleep(3)

    def have_draft(self, draft_name):
        # 点击对应草稿
        draft_name_text = self.app.TextControl(searchDepth=2,
                                               Compare=ControlFinder.desc_matcher(f"HomePageDraftTitle:{draft_name}", exact=True))
        if not draft_name_text.Exists(0):
            return False
        return True

    def find_button_with_timeout(self, searchDepth: int, matcher: Callable[[uia.Control, int], bool], timeout: float) -> uia.Control:
        """在指定时间内找到控件
        
        Args:
            control: 要搜索的控件
            matcher: 匹配器函数
            timeout: 超时时间（秒）
            searchDepth: 搜索深度，默认为2
            
        Returns:
            找到的控件
            
        Raises:
            AutomationError: 超时未找到控件
        """
        st = time.time()
        while True:
            btn = self.app.TextControl(searchDepth=searchDepth, Compare=matcher)
            if btn.Exists(0):
                return btn
            
            if time.time() - st > timeout:
                raise AutomationError("控件查找超时, 时限为%d秒" % timeout)
            
            time.sleep(0.1)
            self.get_window()  # 刷新窗口状态
    def find_button_par_with_timeout(self, searchDepth: int, matcher: Callable[[uia.Control, int], bool], timeout: float) -> uia.Control:
        """在指定时间内找到控件

        Args:
            control: 要搜索的控件
            matcher: 匹配器函数
            timeout: 超时时间（秒）
            searchDepth: 搜索深度，默认为2

        Returns:
            找到的控件

        Raises:
            AutomationError: 超时未找到控件
        """
        st = time.time()
        while True:
            setting_group = self.app.GroupControl(searchDepth=1, foundIndex=4)
            if not setting_group.Exists(0):
                raise AutomationError("未找到导出设置组")
            btn = setting_group.TextControl(searchDepth=searchDepth, Compare=matcher)
            if btn.Exists(0):
                return btn

            if time.time() - st > timeout:
                raise AutomationError("控件查找超时, 时限为%d秒" % timeout)

            time.sleep(0.1)
            self.get_window()  # 刷新窗口状态