import os
import uuid
import shutil
import subprocess
import pymediainfo
import logging

logger = logging.getLogger(__name__)


def get_media_duration_seconds(video_path: str, stream_type: str = "video") -> float:
    """三级方案获取媒体真实时长（秒）

    1. ffprobe stream 级（优先，直接读取流时长）
    2. ffprobe format 级（回退，读取容器层时长）
    3. MediaInfo（最后手段）

    Args:
        video_path: 媒体文件路径
        stream_type: "video" 或 "audio"，决定 ffprobe 选择哪个流
    """
    import os as _os
    _basename = _os.path.basename(video_path)
    _stream_selector = "v:0" if stream_type == "video" else "a:0"

    # 方案1: ffprobe stream 级
    try:
        _ffprobe = shutil.which("ffprobe")
        if _ffprobe:
            logger.info(f"[时长获取] ffprobe路径: {_ffprobe}, 文件: {_basename}")
            result = subprocess.run(
                [_ffprobe, "-v", "error", "-select_streams", _stream_selector,
                 "-show_entries", "stream=duration",
                 "-of", "default=noprint_wrappers=1:nokey=1", video_path],
                capture_output=True, text=True, timeout=10
            )
            logger.info(f"[时长获取] ffprobe stream级 returncode={result.returncode}, stdout={result.stdout.strip()!r}, stderr={result.stderr.strip()!r}")
            if result.returncode == 0 and result.stdout.strip():
                duration_seconds = float(result.stdout.strip().split('\n')[0])
                logger.info(f"[时长获取] ffprobe stream级成功: {_basename} -> {duration_seconds:.2f}秒")
                return duration_seconds
        else:
            logger.info(f"[时长获取] 未找到ffprobe，跳过")
    except Exception as ff_err:
        logger.warning(f"[时长获取] ffprobe stream级失败: {ff_err}")

    # 方案2: ffprobe format 级
    try:
        if _ffprobe:
            logger.info(f"[时长获取] ffprobe stream级未取到，尝试format级")
            result = subprocess.run(
                [_ffprobe, "-v", "error", "-show_entries", "format=duration",
                 "-of", "default=noprint_wrappers=1:nokey=1", video_path],
                capture_output=True, text=True, timeout=10
            )
            logger.info(f"[时长获取] ffprobe format级 returncode={result.returncode}, stdout={result.stdout.strip()!r}, stderr={result.stderr.strip()!r}")
            if result.returncode == 0 and result.stdout.strip():
                duration_seconds = float(result.stdout.strip())
                logger.info(f"[时长获取] ffprobe format级成功: {_basename} -> {duration_seconds:.2f}秒")
                return duration_seconds
    except Exception as ff_err:
        logger.warning(f"[时长获取] ffprobe format级失败: {ff_err}")

    # 方案3: MediaInfo 后备
    try:
        logger.info(f"[时长获取] ffprobe未取到时长，使用MediaInfo后备: {_basename}")
        info = pymediainfo.MediaInfo.parse(video_path, mediainfo_options={"File_TestContinuousFileNames": "0"})
        tracks = info.video_tracks if stream_type == "video" else info.audio_tracks
        if tracks:
            duration_value = tracks[0].duration
            if isinstance(duration_value, (list, tuple)):
                duration_value = duration_value[0] if duration_value else 0
            if isinstance(duration_value, str):
                try:
                    duration_value = float(duration_value)
                except (ValueError, TypeError):
                    duration_value = 0
            duration_seconds = float(duration_value) / 1000.0
            logger.info(f"[时长获取] MediaInfo成功: {_basename} -> {duration_seconds:.2f}秒")
            return duration_seconds
        else:
            logger.warning(f"[时长获取] MediaInfo无{stream_type}轨道: {_basename}")
    except Exception as mi_err:
        logger.warning(f"[时长获取] MediaInfo失败: {mi_err}")

    logger.warning(f"[时长获取] 所有方案均失败，返回0: {_basename}")
    return 0.0

from typing import Optional, Literal
from typing import Dict, Any

class CropSettings:
    """素材的裁剪设置, 各属性均在0-1之间, 注意素材的坐标原点在左上角"""

    upper_left_x: float
    upper_left_y: float
    upper_right_x: float
    upper_right_y: float
    lower_left_x: float
    lower_left_y: float
    lower_right_x: float
    lower_right_y: float

    def __init__(self, *, upper_left_x: float = 0.0, upper_left_y: float = 0.0,
                 upper_right_x: float = 1.0, upper_right_y: float = 0.0,
                 lower_left_x: float = 0.0, lower_left_y: float = 1.0,
                 lower_right_x: float = 1.0, lower_right_y: float = 1.0):
        """初始化裁剪设置, 默认参数表示不裁剪"""
        self.upper_left_x = upper_left_x
        self.upper_left_y = upper_left_y
        self.upper_right_x = upper_right_x
        self.upper_right_y = upper_right_y
        self.lower_left_x = lower_left_x
        self.lower_left_y = lower_left_y
        self.lower_right_x = lower_right_x
        self.lower_right_y = lower_right_y

    def export_json(self) -> Dict[str, Any]:
        return {
            "upper_left_x": self.upper_left_x,
            "upper_left_y": self.upper_left_y,
            "upper_right_x": self.upper_right_x,
            "upper_right_y": self.upper_right_y,
            "lower_left_x": self.lower_left_x,
            "lower_left_y": self.lower_left_y,
            "lower_right_x": self.lower_right_x,
            "lower_right_y": self.lower_right_y
        }

class VideoMaterial:
    """本地视频素材（视频或图片）, 一份素材可以在多个片段中使用"""

    material_id: str
    """素材全局id, 自动生成"""
    local_material_id: str
    """素材本地id, 意义暂不明确"""
    material_name: str
    """素材名称"""
    path: str
    """素材文件路径"""
    duration: int
    """素材时长, 单位为微秒"""
    height: int
    """素材高度"""
    width: int
    """素材宽度"""
    crop_settings: CropSettings
    """素材裁剪设置"""
    material_type: Literal["video", "photo"]
    """素材类型: 视频或图片"""

    def __init__(self, path: str, material_name: Optional[str] = None, crop_settings: CropSettings = CropSettings()):
        """从指定位置加载视频（或图片）素材

        Args:
            path (`str`): 素材文件路径, 支持mp4, mov, avi等常见视频文件及jpg, jpeg, png等图片文件.
            material_name (`str`, optional): 素材名称, 如果不指定, 默认使用文件名作为素材名称.
            crop_settings (`CropSettings`, optional): 素材裁剪设置, 默认不裁剪.

        Raises:
            `FileNotFoundError`: 素材文件不存在.
            `ValueError`: 不支持的素材文件类型.
        """
        path = os.path.abspath(path)
        postfix = os.path.splitext(path)[1]
        if not os.path.exists(path):
            raise FileNotFoundError(f"找不到 {path}")

        self.material_name = material_name if material_name else os.path.basename(path)
        self.material_id = uuid.uuid4().hex
        self.path = path
        self.crop_settings = crop_settings
        self.local_material_id = ""

        if not pymediainfo.MediaInfo.can_parse():
            raise ValueError(f"不支持的视频素材类型 '{postfix}'")

        info: pymediainfo.MediaInfo = \
            pymediainfo.MediaInfo.parse(path, mediainfo_options={"File_TestContinuousFileNames": "0"})  # type: ignore
        # 有视频轨道的视为视频素材
        if len(info.video_tracks):
            self.material_type = "video"
            duration_seconds = get_media_duration_seconds(path)
            self.duration = int(duration_seconds * 1e6)  # 秒→微秒
            # Handle width and height which might also be strings or lists/tuples
            width_value = info.video_tracks[0].width
            height_value = info.video_tracks[0].height
            
            if isinstance(width_value, (list, tuple)):
                width_value = width_value[0] if width_value else 0
            if isinstance(width_value, str):
                try:
                    width_value = int(width_value)
                except (ValueError, TypeError):
                    width_value = 0
            
            if isinstance(height_value, (list, tuple)):
                height_value = height_value[0] if height_value else 0
            if isinstance(height_value, str):
                try:
                    height_value = int(height_value)
                except (ValueError, TypeError):
                    height_value = 0
            
            self.width, self.height = int(width_value), int(height_value)
        # gif文件使用imageio库获取长度
        elif postfix.lower() == ".gif":
            import imageio
            gif = imageio.get_reader(path)

            self.material_type = "gif"
            try:
                # 尝试从元数据获取duration
                meta_data = gif.get_meta_data()
                duration_per_frame = meta_data.get('duration', 0.1)  # 默认100ms每帧
                self.duration = int(round(duration_per_frame * gif.get_length() * 1e3))
                
            except (KeyError, AttributeError):
                # 如果获取失败，使用默认值：假设每帧100ms
                self.duration = int(round(0.1 * gif.get_length() * 1e3))
            self.width, self.height = info.image_tracks[0].width, info.image_tracks[0].height  # type: ignore
            gif.close()
        elif len(info.image_tracks):
            self.material_type = "photo"
            self.duration = 10800000000  # 相当于3h
            self.width, self.height = info.image_tracks[0].width, info.image_tracks[0].height  # type: ignore
        else:
            raise ValueError(f"输入的素材文件 {path} 没有视频轨道或图片轨道")

    def export_json(self) -> Dict[str, Any]:
        video_material_json = {
            "audio_fade": None,
            "category_id": "",
            "category_name": "local",
            "check_flag": 63487,
            "crop": self.crop_settings.export_json(),
            "crop_ratio": "free",
            "crop_scale": 1.0,
            "duration": self.duration,
            "height": self.height,
            "id": self.material_id,
            "local_material_id": self.local_material_id,
            "material_id": self.material_id,
            "material_name": self.material_name,
            "media_path": "",
            "path": self.path,
            "type": self.material_type,
            "width": self.width
        }
        return video_material_json

class AudioMaterial:
    """本地音频素材"""

    material_id: str
    """素材全局id, 自动生成"""
    material_name: str
    """素材名称"""
    path: str
    """素材文件路径"""

    duration: int
    """素材时长, 单位为微秒"""

    def __init__(self, path: str, material_name: Optional[str] = None):
        """从指定位置加载音频素材, 注意视频文件不应该作为音频素材使用

        Args:
            path (`str`): 素材文件路径, 支持mp3, wav等常见音频文件.
            material_name (`str`, optional): 素材名称, 如果不指定, 默认使用文件名作为素材名称.

        Raises:
            `FileNotFoundError`: 素材文件不存在.
            `ValueError`: 不支持的素材文件类型.
        """
        path = os.path.abspath(path)
        if not os.path.exists(path):
            raise FileNotFoundError(f"找不到 {path}")

        self.material_name = material_name if material_name else os.path.basename(path)
        self.material_id = uuid.uuid4().hex
        self.path = path

        if not pymediainfo.MediaInfo.can_parse():
            raise ValueError("不支持的音频素材类型 %s" % os.path.splitext(path)[1])
        info: pymediainfo.MediaInfo = pymediainfo.MediaInfo.parse(path)  # type: ignore
        if len(info.video_tracks):
            raise ValueError("音频素材不应包含视频轨道")
        if not len(info.audio_tracks):
            raise ValueError(f"给定的素材文件 {path} 没有音频轨道")
        duration_seconds = get_media_duration_seconds(path, stream_type="audio")
        self.duration = int(duration_seconds * 1e6)  # 秒→微秒

    def export_json(self) -> Dict[str, Any]:
        return {
            "app_id": 0,
            "category_id": "",
            "category_name": "local",
            "check_flag": 3,
            "copyright_limit_type": "none",
            "duration": self.duration,
            "effect_id": "",
            "formula_id": "",
            "id": self.material_id,
            "local_material_id": self.material_id,
            "music_id": self.material_id,
            "name": self.material_name,
            "path": self.path,
            "source_platform": 0,
            "type": "extract_music",
            "wave_points": []
        }
