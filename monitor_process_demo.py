# $language = "python3"
# $interface = "1.0"

# monitor_process_demo.py
#
# 描述:
#   演示如何以设定的时间间隔重复运行一个命令，并将输出
#   记录到一个本地文件中。
#
#   此脚本将无限期运行，直到会话断开或用户
#   手动停止脚本（通过菜单 "Script" -> "Cancel"）。
#
#   此版本将日志记录和一些通用功能封装在类和函数中，
#   使得主脚本更简洁，专注于核心逻辑。

import os
import time
import re


class SecureCRTLogger:
    """
    一个可插拔的日志记录类，用于在 SecureCRT 脚本中记录信息。
    它将日志输出到一个本地文件中。
    """

    def __init__(self, crt_object, log_file_path=None, clear_log_file_on_start=True):
        """
        初始化日志记录器。

        :param crt_object: SecureCRT 全局的 'crt' 对象。
        :param log_file_path: 本地日志文件的完整路径。如果为 None，将自动生成一个。
        :param clear_log_file_on_start: 是否在脚本开始时清空旧的日志文件。
        """
        self.crt = crt_object
        if log_file_path is None:
            self.log_file_path = os.path.join(os.path.expanduser('~'), 'Documents', 'securecrt_script_log.txt')
        else:
            self.log_file_path = log_file_path

        self.log_tab = None
        self.is_ready = False

        if clear_log_file_on_start:
            self._prepare_log_file()

        self.is_ready = True

    def _prepare_log_file(self):
        """如果文件存在，则删除旧的日志文件，为新的日志做准备。"""
        try:
            if os.path.exists(self.log_file_path):
                os.remove(self.log_file_path)
        except OSError as e:
            self.crt.Dialog.MessageBox(f"警告：无法删除旧的日志文件。\n\n路径: {self.log_file_path}\n原因: {str(e)}")

    def log(self, message):
        """将带时间戳的消息写入日志文件。"""
        if not self.is_ready:
            return

        timestamped_message = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {message}"

        try:
            log_dir = os.path.dirname(self.log_file_path)
            if not os.path.exists(log_dir):
                os.makedirs(log_dir)

            with open(self.log_file_path, 'a', encoding='utf-8') as f:
                f.write(timestamped_message + "\n")
        except Exception as e:
            if not hasattr(self, '_file_error_shown'):
                self.crt.Dialog.MessageBox(f"错误：无法写入日志文件！\n\n路径: {self.log_file_path}\n原因: {str(e)}")
                self._file_error_shown = True


def clean_ansi_codes(text):
    """
    使用正则表达式移除字符串中的 ANSI 转义序列（例如颜色代码、粘贴模式标记等）。
    """
    ansi_escape_pattern = re.compile(r'\x1B(?:\[[0-?]*[ -/]*[@-~]|\].*?(?:\x07|\x1B\\))')
    return ansi_escape_pattern.sub('', text)


def get_prompt(crt_object, tab, logger):
    """
    通过发送换行符并读取当前行来确定命令提示符。

    :param crt_object: SecureCRT 全局的 'crt' 对象。
    :param tab: 脚本运行的标签页对象。
    :param logger: 日志记录器实例。
    :return: 检测到的提示符字符串，如果失败则返回 None。
    """
    logger.log("正在获取设备提示符...")
    tab.Screen.Send("\n")
    tab.Screen.WaitForCursor(1)
    crt_object.Sleep(200)

    current_row = tab.Screen.CurrentRow
    prompt = tab.Screen.Get(current_row, 1, current_row, tab.Screen.CurrentColumn - 1).strip()

    if not prompt:
        logger.log("错误: 无法获取设备提示符。脚本将无法继续。")
        crt_object.Dialog.MessageBox("错误: 无法获取设备提示符。请确保已连接并登录到设备。")
        return None

    logger.log(f"检测到提示符: '{prompt}'")
    return prompt

# --- 配置区 ---
# 在这里修改您想要重复运行的命令。
# Linux 示例: "ps -ef | head -n 15"
# Cisco 示例: "show processes cpu sorted | i ^[0-9]"
COMMAND_TO_RUN = "ps -ef"

# 在这里修改每次执行命令的间隔时间（单位：秒）。
LOOP_INTERVAL_SECONDS = 5

# 日志文件的保存路径。脚本每次运行时会覆盖旧的日志文件。
LOG_FILE_PATH = os.path.join(os.path.expanduser('~'), 'Documents', 'securecrt_monitor_log.txt')
# --------------

def main():
    """运行监控脚本的主函数。"""
    # crt 对象由 SecureCRT 环境提供，是脚本的入口点
    script_tab = crt.GetScriptTab()

    # 确保脚本在已连接的会话中运行。
    if not script_tab.Session.Connected:
        crt.Dialog.MessageBox(
            "错误: 未连接会话。\n\n"
            "请先连接到一个会话再运行此脚本。")
        return

    # 1. 初始化日志记录器
    # 使用 securecrt_utils 模块中的 SecureCRTLogger 类
    logger = SecureCRTLogger(crt, LOG_FILE_PATH)
    if not logger.is_ready:
        # 初始化失败时，构造函数已经显示了错误信息
        return

    # 2. 激活脚本运行的原始标签页并设置同步模式
    script_tab.Activate()
    script_tab.Screen.Synchronous = True

    # 3. 提示用户如何使用
    crt.Dialog.MessageBox(
        f"脚本即将开始监控。\n\n"
        f"日志将写入文件: {LOG_FILE_PATH}\n\n"
        "要停止脚本，请从菜单选择 'Script' -> 'Cancel'。")

    logger.log("脚本开始执行。")

    # 4. 获取命令提示符，以便知道命令何时执行完毕。
    # 使用 securecrt_utils 模块中的 get_prompt 函数
    prompt = get_prompt(crt, script_tab, logger)
    if not prompt:
        logger.log("无法获取提示符，脚本终止。")
        return

    try:
        while True:
            # 在运行命令之前，检查我们是否仍然连接。
            if not script_tab.Session.Connected:
                logger.log("会话已断开，脚本停止。")
                break

            logger.log(f"正在执行命令: {COMMAND_TO_RUN}")

            # 发送命令。
            script_tab.Screen.Send(COMMAND_TO_RUN + "\n")

            # 等待命令本身被回显。
            # 注意: 对于非常短的命令，这可能立即返回。
            # 对于复杂的网络环境，可能需要增加超时时间。
            script_tab.Screen.WaitForString(COMMAND_TO_RUN.strip(), 5)

            # 读取所有输出，直到下一个提示符出现。
            # ReadString 的超时对于捕获长时间运行的命令的输出至关重要。
            output = script_tab.Screen.ReadString(prompt, timeout=10)

            if output:
                # 使用 securecrt_utils 中的函数清理输出中的 ANSI 转义码
                cleaned_output = clean_ansi_codes(output)
                # 记录捕获到的输出
                logger.log(f"命令输出:\n---\n{cleaned_output.strip()}\n---")
            else:
                logger.log("警告: 未能捕获到命令输出或等待提示符超时。")

            # 在下一次循环之前等待指定的时间。
            logger.log(f"等待 {LOOP_INTERVAL_SECONDS} 秒...")
            crt.Sleep(LOOP_INTERVAL_SECONDS * 1000)

    except Exception as e:
        # 捕获循环期间的任何意外错误。
        # SecureCRT 的异常通常是 VBScript 错误，GetLastErrorMessage() 更有用。
        error_msg = crt.GetLastErrorMessage()
        logger.log(f"脚本在执行过程中发生意外错误: {str(e)}")
        if error_msg:
            logger.log(f"SecureCRT Last Error: {error_msg}")
        crt.Dialog.MessageBox(f"脚本发生错误，请检查日志获取详细信息。\n\n错误: {str(e)}")

    logger.log("脚本执行结束。")

# 运行主函数
main()