import frida
import base64
import json
import logging
from copy import deepcopy
from frida.core import Device
from fastapi import APIRouter
from fastapi.responses import JSONResponse
from internal.frida.device import get_device

from internal.response.model import ApiBaseResponse, FAILED_PRECONDITION, OK

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/frida/app")

is_first = True


def safe_str(val):
    """安全地将值转为 UTF-8 字符串，处理非 UTF-8 编码的应用名"""
    if isinstance(val, bytes):
        return val.decode("utf-8", errors="replace")
    if isinstance(val, str):
        return val.encode("utf-8", errors="replace").decode("utf-8")
    return str(val) if val is not None else "-"


@router.get("", response_model=None)
async def getApps():
    res = deepcopy(OK)
    packages = {}
    global is_first
    if is_first == True:
        is_first = False
        return JSONResponse(
            content={"code": 400, "status": "FAILED_PRECONDITION", "description": "需要初始化", "detail": [], "count": 0},
            status_code=200
        )
    try:
        try:
            logger.info("[getApps] 尝试 frida.get_usb_device(timeout=5)...")
            device = frida.get_usb_device(timeout=5)
            logger.info("[getApps] get_usb_device 成功, device=%s", device)
        except Exception as e:
            logger.warning("[getApps] get_usb_device 失败: %s, 尝试 get_remote_device...", e)
            device = frida.get_remote_device()
            logger.info("[getApps] get_remote_device 成功, device=%s", device)
        logger.info("[getApps] 开始 enumerate_applications(scope='full')...")
        apps = device.enumerate_applications(scope="full")
        logger.info("[getApps] enumerate_applications 完成, 应用数=%d", len(apps))
    except Exception as e:
        return JSONResponse(
            content={"code": 400, "status": "FAILED_PRECONDITION", "description": "未找到设备: " + str(e), "detail": [], "count": 0},
            status_code=200
        )
    for i in apps:
        try:
            if i.identifier in packages:
                continue
            packages[i.identifier] = 2
            if len(i.parameters["icons"]) != 0:
                raw_image = i.parameters["icons"][0]["image"]
                icon_format = str(i.parameters["icons"][0].get("format", "png"))
                icon_b64 = base64.b64encode(raw_image).decode("ascii")
                res.detail.append(
                    {
                        "name": safe_str(i.name),
                        "package": safe_str(i.identifier),
                        "version": safe_str(i.parameters.get("version", "-")),
                        "icon": "data:image/" + icon_format + ";base64," + icon_b64,
                    }
                )
            else:
                res.detail.append(
                    {
                        "name": safe_str(i.name),
                        "package": safe_str(i.identifier),
                        "version": safe_str(i.parameters.get("version", "-")),
                        "icon": "",
                    }
                )
        except Exception as e:
            logger.warning("[getApps] 跳过应用 %s, 原因: %s", safe_str(i.name), e)
            continue
    res.count = len(apps)
    logger.info("[getApps] 返回应用数=%d", len(res.detail))
    # 手动序列化，确认每个 icon 都是字符串而非 bytes
    result_dict = res.model_dump()
    for item in result_dict.get("detail", []):
        if isinstance(item.get("icon"), bytes):
            logger.error("[getApps] 发现 icon 仍是 bytes! name=%s", item.get("name"))
            item["icon"] = ""
    try:
        json_str = json.dumps(result_dict, ensure_ascii=False, default=str)
        return JSONResponse(content=json.loads(json_str), status_code=200)
    except Exception as e:
        logger.error("[getApps] JSON序列化失败: %s, 前200字符: %s", e, str(result_dict)[:200])
        return {"code": 500, "status": "ERROR", "description": "JSON序列化失败: " + str(e)[:200], "detail": [], "count": 0}
