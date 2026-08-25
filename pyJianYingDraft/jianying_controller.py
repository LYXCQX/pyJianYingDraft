"""剪映自动化控制，主要与自动导出有关"""

import time
import shutil
import threading
import uiautomation as uia
import os
import subprocess
import psutil
from loguru import logger

from enum import Enum
from typing import Optional, Literal, Callable

from . import exceptions
from .exceptions import AutomationError

class ExportResolution(Enum):
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

class ExportFramerate(Enum):
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

class JianyingController:
    """剪映控制器"""

    app: uia.WindowControl
    """剪映窗口"""
    app_status: Literal["home", "edit", "pre_export"]
    jianying_exe_path: Optional[str] = None
    """剪映可执行文件路径"""

    def __init__(self, set_top=True, jianying_exe_path: Optional[str] = None):
        """初始化剪映控制器, 此时剪映应该处于目录页
        
        Args:
            set_top: 是否置顶窗口
            jianying_exe_path: 剪映可执行文件路径，用于重启剪映
        """
        self.jianying_exe_path = jianying_exe_path
        self.get_window(set_top)
        # 跨线程取消信号：外部线程调用 cancel_export() 置位，
        # export_draft 的 while 循环检测到后抛 AutomationError 快速退出，避免长时间阻塞
        self._cancel_event = threading.Event()

    def cancel_export(self):
        """请求取消正在进行的导出操作（线程安全，可从其他线程调用）"""
        self._cancel_event.set()
        logger.info("[Jianying_controller] 收到取消导出请求")

    def _check_cancelled(self):
        """检查取消信号，如已被取消则抛 AutomationError 终止 export_draft"""
        if self._cancel_event.is_set():
            raise AutomationError("导出已被用户取消")

    def _move_exported_file(self, export_path: str, output_path: Optional[str],
                            draft_name: str, juming: Optional[str]) -> Optional[str]:
        """移动导出文件到目标目录并返回最终路径。

        Args:
            export_path: 剪映实际导出的文件或目录路径
            output_path: 用户指定的目标根目录（如果为None则不移动）
            draft_name: 草稿名
            juming: 剧名子目录名

        Returns:
            移动后的最终文件路径；若 output_path 为 None 则返回 None
        """
        if output_path is None:
            return None

        export_filename = os.path.basename(export_path)

        if os.path.isdir(export_path):
            video_files = [f for f in os.listdir(export_path) if f.endswith(('.mp4', '.mov', '.avi', '.wmv'))]
            if video_files:
                video_name = video_files[0]
                final_path = os.path.join(output_path, export_filename, video_name)

                cover_file = os.path.join(export_path, "video_cover.jpg")
                if os.path.exists(cover_file):
                    target_dir = os.path.join(output_path, export_filename)
                    if not os.path.exists(target_dir):
                        os.makedirs(target_dir)
                    video_name_without_ext = os.path.splitext(video_name)[0]
                    target_cover = os.path.join(target_dir, f"{video_name_without_ext}_cover.jpg")
                    shutil.copy2(cover_file, target_cover)
        else:
            output_path = os.path.join(output_path, juming)
            final_path = os.path.join(output_path, export_filename)
            if not os.path.exists(output_path):
                os.makedirs(output_path)

            draft_dir = os.path.dirname(export_path)
            cover_file = os.path.join(draft_dir, "video_cover.jpg")
            if os.path.exists(cover_file):
                export_name_without_ext = os.path.splitext(export_filename)[0]
                target_cover = os.path.join(output_path, f"{export_name_without_ext}_cover.jpg")
                shutil.copy2(cover_file, target_cover)

        shutil.move(export_path, output_path)
        return final_path

    def _check_file_exists_stable(self, path: str, stable_checks: int = 3, interval: float = 2.0) -> bool:
        """检查文件/目录是否存在且大小稳定（导出仍在写入时大小会持续变化）。

        Args:
            path: 待检查路径
            stable_checks: 连续多少次大小一致才认为稳定（>=1）
            interval: 每次检查间隔秒数

        Returns:
            文件存在且大小稳定返回 True，否则 False
        """
        if not path or not os.path.exists(path):
            return False
        if stable_checks <= 1:
            return True

        prev_size = None
        consistent = 0
        for _ in range(stable_checks):
            try:
                if os.path.isdir(path):
                    total = 0
                    for root, _, files in os.walk(path):
                        for f in files:
                            fp = os.path.join(root, f)
                            if os.path.exists(fp):
                                total += os.path.getsize(fp)
                    cur_size = total
                else:
                    cur_size = os.path.getsize(path)
            except OSError:
                return False

            if prev_size is not None and cur_size == prev_size:
                consistent += 1
            else:
                consistent = 0
            if consistent >= stable_checks - 1:
                return True
            prev_size = cur_size
            time.sleep(interval)
        return False

    def export_draft(self, draft_name: str, output_path: Optional[str] = None, *,
                     resolution: Optional[ExportResolution] = None,
                     framerate: Optional[ExportFramerate] = None,
                     timeout: float = 1200,juming= None) -> None:
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
        # self.close_relink_window()
        # time.sleep(6)
        self.get_window()

        start_time = time.time()
        while True:
            self._check_cancelled()
            if time.time() - start_time > 20:
                raise AutomationError(f"未找到导出路径，超时时间：{20}秒")
            export_btn = self.app.TextControl(searchDepth=2, Compare=ControlFinder.desc_matcher("MainWindowTitleBarExportBtn"))
            if not export_btn.Exists(0):
                if self._cancel_event.wait(2):
                    raise AutomationError("导出已被用户取消")
                self.get_window()
                assert draft_btn is not None
                draft_btn.Click(simulateMove=False)
                continue
            self.send_keys('{Ctrl}e',1)
            self.get_window()
            # 获取原始导出路径（带后缀名）
            export_path_sib = self.app.TextControl(searchDepth=2, Compare=ControlFinder.desc_matcher("ExportPath"))
            if export_path_sib.Exists(0):
                break
        export_path_text = export_path_sib.GetSiblingControl(lambda ctrl: True)
        assert export_path_text is not None
        export_path = export_path_text.GetPropertyValue(30159)

        # 保存导出路径，供异常回退判断使用
        export_path_for_fallback = export_path
        export_filename_for_fallback = os.path.basename(export_path) if export_path else None

        # ✅ 导出前再次校验：导出文件名必须以草稿名开头
        # 防止 UI 点击错位 / 剪映首页草稿重排 / 弹窗拦截后点进了错误草稿，
        # 结果把别人家的内容导出来（最严重会把错误的成品发到线上）。
        # 校验失败：关闭当前导出窗口回到首页，然后抛 DraftNotFound（与草稿不存在一致的错误语义）。
        export_filename = os.path.splitext(os.path.basename(export_path))[0]  if export_path else ""
        match = bool(draft_name) and bool(export_filename) and draft_name.startswith(export_filename)
        logger.warning(
            f"[export_draft] 草稿名前缀校验: draft_name={draft_name!r}, export_path={export_path!r}, "
            f"export_filename={export_filename!r}, match={match}  {bool(draft_name)}  { bool(export_filename)}  {draft_name.startswith(export_filename)}"
        )
        if not match:
            logger.error(
                f"[export_draft] ❌ 导出文件名不以草稿名开头，疑似点错了草稿！"
                f"draft_name={draft_name!r} 但 export_filename={export_filename!r}，将回首页并抛出 DraftNotFound。"
            )
            # 尝试关闭当前导出弹层（按 ESC），然后切回首页
            try:
                self.send_keys('{ESC}', 3)
            except Exception:
                pass
            try:
                self.switch_to_home()
            except Exception:
                pass
            raise exceptions.DraftNotFound(
                f"草稿导出名校验失败：draft_name={draft_name!r}，实际导出文件={export_filename!r}，"
                f"疑似点击了错误的草稿，已回首页。请重新尝试或手动确认首页草稿顺序。"
            )

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
            self._check_cancelled()
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
            if self._cancel_event.wait(0.5):
                raise AutomationError("导出已被用户取消")

        # 等待导出完成
        export_completed_normally = False
        st = time.time()
        export_error: Optional[Exception] = None
        try:
            while True:
                self._check_cancelled()
                # self.get_window()
                if self.app_status != "pre_export": continue
                has_close = False
                succeed_close_btn = self.app.TextControl(searchDepth=2, Compare=ControlFinder.desc_matcher("ExportSucceedCloseBtn"))
                if succeed_close_btn.Exists(0):
                    start_time = time.time()
                    while True:
                        self._check_cancelled()
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
                        if self._cancel_event.wait(0.5):
                            raise AutomationError("导出已被用户取消")

                if time.time() - st > timeout:
                    raise AutomationError("导出超时, 时限为%d秒" % timeout)
                if has_close:
                    break
                if self._cancel_event.wait(1):
                    raise AutomationError("导出已被用户取消")
            export_completed_normally = True
        except Exception as e:
            export_error = e
            logger.warning(
                f"[export_draft] 等待导出成功弹窗失败或超时：{e}。"
                f"将回退检查 export_path={export_path_for_fallback!r} 是否已经存在导出文件。"
            )

        # 失败兜底：如果 UI 层面没等到成功提示，但导出路径文件已经存在且大小稳定，
        # 则认为实际上已经导出完成（常见于剪映成功提示未弹 / 弹了但识别不到 / 卡死）。
        if not export_completed_normally:
            fallback_path = export_path_for_fallback
            fallback_filename = export_filename_for_fallback
            fallback_ok = False
            if fallback_path:
                # 额外校验：兜底时也检查文件名以草稿名开头，避免把其它文件当成成功导出
                fallback_match = bool(draft_name) and bool(fallback_filename) and fallback_filename.startswith(draft_name)
                if not fallback_match:
                    logger.error(
                        f"[export_draft] 兜底路径文件名不匹配草稿名前缀，跳过兜底判断："
                        f"draft_name={draft_name!r}, fallback_filename={fallback_filename!r}"
                    )
                elif self._check_file_exists_stable(fallback_path, stable_checks=3, interval=2.0):
                    logger.warning(
                        f"[export_draft] ✅ 兜底成功：虽然未检测到剪映成功弹窗，"
                        f"但导出文件已存在且大小稳定，视为导出成功。path={fallback_path!r}"
                    )
                    fallback_ok = True
                    export_path = fallback_path
                    # 尽量恢复到主页状态，便于后续草稿导出
                    try:
                        self.switch_to_home()
                    except Exception:
                        pass

            if not fallback_ok:
                logger.error(
                    f"[export_draft] ❌ 兜底失败：导出文件不存在或仍在写入中，"
                    f"将抛出原异常。path={fallback_path!r}, err={export_error}"
                )
                raise export_error

        # 移动文件到目标目录
        return self._move_exported_file(export_path, output_path, draft_name, juming)

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
    
    def is_jianying_process_running(self) -> bool:
        """检查剪映进程是否在运行"""
        for proc in psutil.process_iter(['name']):
            try:
                if proc.info['name'] and proc.info['name'].lower() == 'jianyingpro.exe':
                    return True
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        return False
    
    def kill_jianying_process(self) -> None:
        """强制杀掉剪映进程"""
        killed = False
        for proc in psutil.process_iter(['name', 'pid']):
            try:
                if proc.info['name'] and proc.info['name'].lower() == 'jianyingpro.exe':
                    proc.kill()
                    killed = True
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        if killed:
            time.sleep(2)  # 等待进程完全结束
    
    def start_jianying(self) -> None:
        """启动剪映程序"""
        if self.jianying_exe_path and os.path.exists(self.jianying_exe_path):
            subprocess.Popen([self.jianying_exe_path])
            # 等待剪映启动，最多等待30秒
            max_wait = 30
            wait_interval = 1
            elapsed = 0
            while elapsed < max_wait:
                time.sleep(wait_interval)
                elapsed += wait_interval
                # 检查进程是否已启动
                if self.is_jianying_process_running():
                    # 进程已启动，再等待几秒让界面完全加载
                    time.sleep(3)
                    return
            raise AutomationError(f"剪映启动超时（{max_wait}秒内未检测到进程）")
        else:
            raise AutomationError("未配置剪映可执行文件路径或路径不存在")
    
    def get_window(self, set_top=True) -> None:
        """寻找剪映窗口并置顶，如果找不到则检查进程并重启"""
        if hasattr(self, "app") and self.app.Exists(0):
            self.app.SetTopmost(False)

        self.app = uia.WindowControl(searchDepth=1, Compare=self.__jianying_window_cmp)
        if not self.app.Exists(0):
            # 窗口未找到，检查进程
            if self.is_jianying_process_running():
                # 进程存在但窗口找不到，强制杀掉进程
                self.kill_jianying_process()
                # 重新启动剪映
                self.start_jianying()
                # 再次尝试查找窗口，增加重试次数
                max_retries = 5
                for i in range(max_retries):
                    time.sleep(2)  # 每次重试前等待2秒
                    self.app = uia.WindowControl(searchDepth=1, Compare=self.__jianying_window_cmp)
                    if self.app.Exists(0):
                        break
                else:
                    raise AutomationError("剪映窗口未找到（重启后仍未找到）")
            else:
                # 进程不存在，直接启动
                self.start_jianying()
                # 尝试查找窗口，增加重试次数
                max_retries = 5
                for i in range(max_retries):
                    time.sleep(2)  # 每次重试前等待2秒
                    self.app = uia.WindowControl(searchDepth=1, Compare=self.__jianying_window_cmp)
                    if self.app.Exists(0):
                        break
                else:
                    raise AutomationError("剪映窗口未找到（启动后仍未找到）")

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
