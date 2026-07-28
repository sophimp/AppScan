import os
import time
import platform
import subprocess
import logging
from copy import deepcopy
from fastapi import APIRouter

from internal.response.model import (
    ApiBaseResponse,
    FAILED_PRECONDITION,
    OK,
    USB_CLOSED,
    USB_UNAUTHORIZED,
    ROOT_CLOSED,
)

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/adb/init")

adb_path = (
    os.path.abspath(os.path.join(os.path.dirname(__file__), "../.."))
    + os.sep
    + "static"
    + os.sep
    + "windows"
    + os.sep
    + "adb.exe"
)  # default windows
if platform.system().lower() == "darwin":
    # mac 环境
    adb_path = (
        os.path.abspath(os.path.join(os.path.dirname(__file__), "../.."))
        + os.sep
        + "static"
        + os.sep
        + "darwin"
        + os.sep
        + "adb"
    )  # 默认mac环境
frida_server_arm = "hluda-server-arm64"
frida_server_x86 = "hluda-server-x86"
frida_server_x86_64 = "hluda-server-x86_64"
# 根据手机架构选择 frida-server, arm和x86
# 兼容模拟器
detecting_phone_architecture_cmd = [adb_path, "shell", "su -c 'getprop ro.product.cpu.abi'"]
frida_server = ""
frida_path = (
    os.path.abspath(os.path.join(os.path.dirname(__file__), "../.."))
    + os.sep
    + "static"
    + os.sep
    + frida_server
)
colse_SELinux_cmd = [adb_path, "shell", "su -c 'setenforce 0'"]
kill_cmd = [adb_path, "shell", "su -c 'pkill -9 hluda'"]
clean_cmd = [adb_path, "shell", "su -c 'rm -rf /data/local/tmp/*'"]
push_cmd = [adb_path, "push", frida_path, "/storage/emulated/0/{}".format(frida_server)]
mv_cmd = [
    adb_path,
    "shell",
    "su -c 'mv /storage/emulated/0/{} /data/local/tmp/'".format(frida_server),
]
chmod_cmd = [
    adb_path,
    "shell",
    "su -c 'chmod 777 /data/local/tmp/{}'".format(frida_server),
]
run_cmd = [adb_path, "shell", "su -c 'nohup /data/local/tmp/{} &'".format(frida_server)]
devices_cmd = [adb_path, "devices"]
root_cmd = [adb_path, "shell", "su -c 'exit'"]
stop_adb_cmd = [adb_path, "kill-server"]
start_adb_cmd = [adb_path, "start-server"]
# https://github.com/frida/frida/issues/1788
close_usap_cmd = [
    adb_path,
    "shell",
    "su -c 'setprop persist.device_config.runtime_native.usap_pool_enabled false'",
]

def generation_cmd():
    # 重新生成cmd
    global adb_path
    global frida_server
    global frida_path
    global colse_SELinux_cmd
    global kill_cmd
    global clean_cmd
    global push_cmd
    global mv_cmd
    global chmod_cmd
    global run_cmd
    global devices_cmd
    global root_cmd
    global stop_adb_cmd
    global start_adb_cmd
    global close_usap_cmd
    global detecting_phone_architecture_cmd
    global frida_server_arm
    global frida_server_x86
    global frida_server_x86_64
    adb_path = (
        os.path.abspath(os.path.join(os.path.dirname(__file__), "../.."))
        + os.sep
        + "static"
        + os.sep
        + "windows"
        + os.sep
        + "adb.exe"
    )  # default windows
    if platform.system().lower() == "darwin":
        # mac 环境
        adb_path = (
            os.path.abspath(os.path.join(os.path.dirname(__file__), "../.."))
            + os.sep
            + "static"
            + os.sep
            + "darwin"
            + os.sep
            + "adb"
        )  # 默认mac环境
    frida_server_arm = "hluda-server-arm64"
    frida_server_x86 = "hluda-server-x86"
    frida_server_x86_64 = "hluda-server-x86_64"
    # 根据手机架构选择 frida-server, arm和x86
    # 兼容模拟器
    detecting_phone_architecture_cmd = [adb_path, "shell", "su -c 'getprop ro.product.cpu.abi'"]
    frida_path = (
        os.path.abspath(os.path.join(os.path.dirname(__file__), "../.."))
        + os.sep
        + "static"
        + os.sep
        + frida_server
    )
    colse_SELinux_cmd = [adb_path, "shell", "su -c 'setenforce 0'"]
    kill_cmd = [adb_path, "shell", "su -c 'pkill -9 hluda'"]
    clean_cmd = [adb_path, "shell", "su -c 'rm -rf /data/local/tmp/*'"]
    push_cmd = [adb_path, "push", frida_path, "/storage/emulated/0/{}".format(frida_server)]
    mv_cmd = [
        adb_path,
        "shell",
        "su -c 'mv /storage/emulated/0/{} /data/local/tmp/'".format(frida_server),
    ]
    chmod_cmd = [
        adb_path,
        "shell",
        "su -c 'chmod 777 /data/local/tmp/{}'".format(frida_server),
    ]
    run_cmd = [adb_path, "shell", "su -c 'nohup /data/local/tmp/{} &'".format(frida_server)]
    devices_cmd = [adb_path, "devices"]
    root_cmd = [adb_path, "shell", "su -c 'exit'"]
    stop_adb_cmd = [adb_path, "kill-server"]
    start_adb_cmd = [adb_path, "start-server"]
    # https://github.com/frida/frida/issues/1788
    close_usap_cmd = [
        adb_path,
        "shell",
        "su -c 'setprop persist.device_config.runtime_native.usap_pool_enabled false'",
    ]

def detecting_phone_architecture():
    # 检测手机架构
    global frida_server
    result = subprocess.Popen(detecting_phone_architecture_cmd, stdout=subprocess.PIPE).communicate()
    outdata = result[0].decode("utf-8")
    if "arm" in outdata:
        frida_server = frida_server_arm
    elif "x86_64" in outdata:
        frida_server = frida_server_x86_64
    elif "x86" in outdata:
        frida_server = frida_server_x86
    else:
        raise Exception("手机架构不支持", outdata)
    return frida_server

@router.post("", response_model=ApiBaseResponse, response_model_exclude_unset=False)
async def init():
    res = deepcopy(OK)
    try:
        # https://github.com/zhengjim/camille/pull/32/commits/1be9236d7b0d8d4369ba0e0e84df5c660dc35c87
        # 重启一下 adb, 防止 adb 偶尔抽风
        # 停止 adb
        logger.info("[init] 步骤1: 停止 ADB...")
        subprocess.call(stop_adb_cmd)
        # 启动adb
        logger.info("[init] 步骤2: 启动 ADB...")
        subprocess.call(start_adb_cmd)
        time.sleep(5)
        # https://github.com/zhengjim/camille/pull/32/commits/e3084d92ba0db4206409246d5e8145c9b5820640
        logger.info("[init] 步骤3: 列出设备...")
        subprocess.call(devices_cmd)
        # 关闭SELinux
        logger.info("[init] 步骤4: 关闭 SELinux...")
        subprocess.call(colse_SELinux_cmd)
        # https://github.com/frida/frida/issues/1788 适配ROM
        logger.info("[init] 步骤5: 关闭 USAP...")
        subprocess.call(close_usap_cmd)
        # kill 可能残留的进程
        logger.info("[init] 步骤6: 杀掉残留 hluda 进程...")
        subprocess.call(kill_cmd)
        time.sleep(2)
        # 获取手机架构
        logger.info("[init] 步骤7: 检测手机架构...")
        detecting_phone_architecture()
        logger.info("[init] 手机架构检测结果: frida_server=%s", frida_server)
        generation_cmd()
        logger.info("[init] frida_path=%s", frida_path)
        # 清理数据
        logger.info("[init] 步骤8: 清理 /data/local/tmp/...")
        subprocess.call(clean_cmd)
        # 推送 frida-server 到设备
        logger.info("[init] 步骤9: 推送 frida-server 到设备...")
        result = subprocess.call(push_cmd)
        logger.info("[init] push 结果: %d", result)
        time.sleep(3)
        # 移动文件
        logger.info("[init] 步骤10: 移动文件到 /data/local/tmp/...")
        subprocess.call(mv_cmd)
        # 设置权限
        logger.info("[init] 步骤11: 设置权限 chmod 777...")
        subprocess.call(chmod_cmd)
        # 启动
        logger.info("[init] 步骤12: 启动 frida-server...")
        pid = subprocess.Popen(run_cmd)
        time.sleep(5)
        logger.info("[init] 步骤13: 关闭启动子进程...")
        pid.kill()
        logger.info("[init] 初始化完成!")
    except Exception as e:
        res = deepcopy(FAILED_PRECONDITION)
        res.description = "初始化错误: " + str(e)
        return res
    return res


@router.get(
    "/verify", response_model=ApiBaseResponse, response_model_exclude_unset=False
)
async def verify():
    res = deepcopy(OK)
    try:
        # 确认是否打开了 usb 调试
        result = subprocess.Popen(devices_cmd, stdout=subprocess.PIPE).communicate()
        if (
            result[0].decode("utf-8").split("\n")[1] == ""
            or result[0].decode("utf-8").split("\n")[1] == "\r"
        ):  # 兼容win
            # 未打开 USB 调试
            # 设备未连接
            res = deepcopy(USB_CLOSED)
            return res
        if result[0].decode("utf-8").split("\n")[1].split()[1] == "unauthorized":
            # 未打开 USB 调试
            # 设备未连接
            res = deepcopy(USB_UNAUTHORIZED)
            return res
        if result[0].decode("utf-8").split("\n")[1].split()[1] == "device":
            root_check = subprocess.call(root_cmd)
            if root_check != 0:
                res = deepcopy(ROOT_CLOSED)
                return res
            return res
    except Exception as e:
        res = deepcopy(FAILED_PRECONDITION)
        res.description = "api错误: " + str(e)
        return res
    return res
