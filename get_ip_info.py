# $language = "python3"
# $interface = "1.0"

import pandas as pd
import time

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


# Regex patterns are compiled once at module level for efficiency.
# This pattern matches the specific line-overwrite sequence used by some devices
# during pagination (e.g., <whitespace><cursor-back><whitespace><cursor-back>).
_LINE_OVERWRITE_PATTERN = re.compile(r'\s*\x1b\[\d+D\s+\x1b\[\d+D')

# This is a general-purpose pattern to find and remove most ANSI escape codes
# (like color codes, etc.).
_ANSI_ESCAPE_PATTERN = re.compile(r'\x1B(?:\[[0-?]*[ -/]*[@-~]|\].*?(?:\x07|\x1B\\))')


def clean_terminal_output(text):
    """
    Cleans raw terminal output by first handling special pagination sequences
    and then removing all other generic ANSI escape codes.

    This two-step process is crucial because pagination overwrite sequences
    (e.g., cursor-back codes mixed with spaces) must be handled differently
    (replaced with a newline) than standard ANSI codes (simply removed).

    :param text: The raw string captured from the terminal.
    :return: A cleaned string with special sequences and ANSI codes removed.
    """
    # Step 1: Replace the specific pagination overwrite sequence with a newline.
    # This fixes the formatting issues caused by clearing the "---- More ----" prompt.
    cleaned_text = _LINE_OVERWRITE_PATTERN.sub('\n', text)

    # Step 2: Remove all other remaining ANSI escape codes.
    cleaned_text = _ANSI_ESCAPE_PATTERN.sub('', cleaned_text)

    return cleaned_text


def get_prompt(crt_object, tab, logger):
    """
    通过发送换行符并读取当前行来确定命令提示符。

    :param crt_object: SecureCRT 全局的 'crt' 对象。
    :param tab: 脚本运行的标签页对象。
    :param logger: 日志记录器实例。
    :return: 检测到的提示符字符串，如果失败则返回 None。
    """
    logger.log("正在获取设备提示符...")
    # tab.Screen.Send(chr(26))
    # tab.Screen.WaitForCursor(1)
    # crt_object.Sleep(200)

    current_row = tab.Screen.CurrentRow
    prompt = tab.Screen.Get(current_row, 1, current_row, tab.Screen.CurrentColumn - 1).strip()

    if not prompt:
        logger.log("错误: 无法获取设备提示符。脚本将无法继续。")
        crt_object.Dialog.MessageBox("错误: 无法获取设备提示符。请确保已连接并登录到设备。")
        return None

    logger.log(f"检测到提示符: '{prompt}'")
    return prompt

def send_command_and_read_output(tab, command, prompt, logger):
    """
    发送命令并读取所有输出，能自动处理分页。
    此版本基于官方文档，使用 ReadString 和 MatchIndex 属性来高效处理分页。

    :param tab: 脚本运行的标签页对象。
    :param command: 要发送的命令字符串 (需要包含换行符)。
    :param prompt: 命令提示符。
    :param logger: 日志记录器实例。
    :return: 清理过的完整命令输出。
    """
    logger.log(f"发送命令: {command.strip()}")

    more_prompt = "---- More ----"
    # 增加超时以应对网络延迟或设备响应慢的情况
    timeout = 10

    tab.Screen.Send(command)

    output_accumulator = ""
    while True:
        # ReadString会捕获数据，直到匹配到列表中的一个字符串或超时
        captured_data = tab.Screen.ReadString([more_prompt, prompt], timeout)
        logger.log(captured_data)
        output_accumulator += captured_data

        # Screen.MatchIndex 会告诉我们是哪个字符串被匹配了 (1-based)
        match_index = tab.Screen.MatchIndex

        if match_index == 1:  # Matched more_prompt
            logger.log("检测到分页符，发送空格并继续读取。")
            tab.Screen.Send(" ")
        elif match_index == 2:  # Matched prompt
            logger.log("检测到最终提示符，输出读取完毕。")
            break
        elif match_index == 0:  # Timeout
            logger.log(f"警告: 等待 '{prompt}' 或 '{more_prompt}' 超时。输出可能不完整。")
            break
        else:  # Should not happen
            logger.log(f"警告: ReadString 返回了未知的 MatchIndex: {match_index}")
            break

    # --- 后续清理逻辑 ---

    # 1. 清理终端输出
    #    调用一个综合性的清理函数，该函数首先处理分页导致的行覆盖序列，
    #    然后移除所有其他通用的ANSI转义码。
    cleaned_output = clean_terminal_output(output_accumulator)

    # 2. 移除命令回显
    command_sent = command.strip()
    if cleaned_output.lstrip().startswith(command_sent):
        # 找到命令第一次出现的位置并切片
        command_start_index = cleaned_output.find(command_sent)
        if command_start_index != -1:
            cleaned_output = cleaned_output[command_start_index + len(command_sent):].lstrip()

    # 3. 移除最后的命令提示符
    if cleaned_output.rstrip().endswith(prompt):
        cleaned_output = cleaned_output.rstrip()[:-len(prompt)].rstrip()

    # 4. 移除整个输出块前后可能存在的空白行和空格
    cleaned_output = cleaned_output.strip()

    logger.log(f"完整命令输出: {cleaned_output}")
    return cleaned_output

def parse_routing_info(output, logger):
    """
    解析 'dis ip rou' 命令的输出，提取路由信息。
    注意：此函数基于对华为设备输出格式的通用假设，您可能需要根据实际情况微调。

    :param output: 命令的完整输出。
    :param logger: 日志记录器实例。
    :return: 包含路由信息的字典，如 {'protocol': 'IBGP', 'nexthop': 'x.x.x.x', 'interface': '...'}，或 None。
    """
    lines = output.splitlines()
    # 寻找路由信息行。通常在 "Destination/Mask" 表头之后。
    header_found = False
    for line in lines:
        line = line.strip()
        if "Destination/Mask" in line and "NextHop" in line:
            header_found = True
            continue

        if header_found and line:
            parts = line.split()
            if len(parts) >= 7:
                try:
                    route_info = {
                        'destination': parts[0],
                        'protocol': parts[1],
                        'flag': parts[4],
                        'nexthop': parts[5],
                        'interface': parts[6]
                    }
                    logger.log(f"解析到路由信息: {route_info}")
                    return route_info
                except IndexError:
                    logger.log(f"警告: 无法完整解析路由行: '{line}'")
                    continue

    logger.log("错误: 未能在输出中找到或解析路由信息。")
    return None

def get_interface_description(tab, interface, prompt, logger, send_cmd_helper):
    """
    获取指定接口的描述信息，并根据原始逻辑提取特定部分。
    """
    command = f"dis cur int {interface}\n"
    output = send_cmd_helper(tab, command, prompt, logger)

    for line in output.splitlines():
        line = line.strip()
        if line.startswith("description"):
            logger.log(f"找到描述行: '{line}'")
            # 原始逻辑: 从 'description AH-HF-xxxx.yyyy' 中提取 'xxxx'
            match = re.search(r'AH-HF-([^.]+)', line)
            if match:
                result = match.group(1)
                logger.log(f"提取到描述片段: '{result}'")
                return result
            else:
                # 如果不匹配，返回 'description' 后的整个字符串作为备用
                return line.split("description", 1)[1].strip()

    logger.log(f"错误: 未找到接口 {interface} 的 'description' 行。")
    return None

# 日志文件的保存路径。脚本每次运行时会覆盖旧的日志文件。
LOG_FILE_PATH = os.path.join(os.path.expanduser('~'), 'Documents', 'securecrt_monitor_log.txt')

def main():
    # crt 对象由 SecureCRT 环境提供，是脚本的入口点
    script_tab = crt.GetScriptTab()

    # 确保脚本在已连接的会话中运行。
    if not script_tab.Session.Connected:
        crt.Dialog.MessageBox(
            "错误: 未连接会话。\n\n"
            "请先连接到一个会话再运行此脚本。")
        return

    # 1. 初始化日志记录器
    logger = SecureCRTLogger(crt, LOG_FILE_PATH)
    if not logger.is_ready:
        # 初始化失败时，构造函数已经显示了错误信息
        return

    # 2. 激活脚本运行的原始标签页并设置同步模式
    script_tab.Activate()
    script_tab.Screen.Synchronous = True

    # 3. 提示用户如何使用
    crt.Dialog.MessageBox(
        f"脚本即将开始执行。\n\n"
        f"日志将写入文件: {LOG_FILE_PATH}\n\n"
        "要停止脚本，请从菜单选择 'Script' -> 'Cancel'。")

    logger.log("脚本开始执行。")

    # 4. 获取命令提示符，以便知道命令何时执行完毕。
    prompt = get_prompt(crt, script_tab, logger)
    if not prompt:
        logger.log("无法获取提示符，脚本终止。")
        return

    # 5. 让用户选择包含IP地址的Excel文件
    file_path = crt.Dialog.FileOpenDialog(title="选择专线号excel表")
    if not file_path:
        crt.Dialog.MessageBox("未选择文件，脚本终止。")
        return

    logger.log(f"正在读取文件: {file_path}")
    host_info = pd.read_excel(file_path, sheet_name=0, usecols='A')
    ips = host_info.iloc[:, 0].to_list()
    results = []

    # 6. 遍历IP列表，执行命令并处理结果
    for ip in ips:
        logger.log(f"--- 开始处理 IP: {ip} ---")
        final_result = ""  # 默认结果

        try:
            # 步骤 6.1: 获取路由信息
            command1 = f"dis ip rou {ip}\n"
            output1 = send_command_and_read_output(script_tab, command1, prompt, logger)

            if not output1:
                logger.log(f"错误: 命令 '{command1}' 没有返回任何有效输出。")
                final_result = "命令无输出"
            else:
                route_info = parse_routing_info(output1, logger)

                if route_info:
                    nexthop = route_info.get('nexthop')
                    flag = route_info.get('flag')
                    interface = route_info.get('interface')
                    protocol = route_info.get('protocol')

                    # 规则 1: 普通BAS设备 (protocol为IBGP，一般nexthop前三位固定为61.133.137)
                    if protocol == 'IBGP':
                        final_result = nexthop
                    # 规则 2: SR设备 (Flags为'D')
                    elif flag == 'D' and interface:
                        logger.log(f"检测到SR设备 (Flag='D'), 查询接口 {interface} 描述...")
                        desc_command = f"dis cur int {interface}\n"
                        desc_output = send_command_and_read_output(script_tab, desc_command, prompt, logger)

                        desc_found = False
                        for line in desc_output.splitlines():
                            line = line.strip()
                            if line.startswith("description dT:"):
                                logger.log(f"找到SR设备描述行: '{line}'")
                                # 提取 "AH-HF-" 和第一个 "." 之间的内容
                                match = re.search(r'AH-HF-([^.]+)', line)
                                if match:
                                    final_result = match.group(1)
                                    logger.log(f"提取到描述片段: '{final_result}'")
                                else:
                                    final_result = "SR描述未匹配提取规则"
                                desc_found = True
                                break
                        if not desc_found:
                            final_result = f"未找到接口 {interface} 的'description dT:'"
                    # 规则 3: 新城设备 (NextHop以124开头)
                    elif nexthop and nexthop.startswith('124'):
                        final_result = "新城"
                    # 规则 4: 出省地址 (NextHop以202开头)
                    elif nexthop and nexthop.startswith('202'):
                        final_result = "出省地址"
                    else:
                        # 如果没有任何规则匹配，可以提供一个默认值
                        final_result = "未匹配任何已知规则"
                else:
                    final_result = "未找到路由信息"
        except Exception as e:
            logger.log(f"处理IP {ip} 时发生未知错误: {str(e)}")
            final_result = "处理时发生错误"

        logger.log(f"IP {ip} 的处理结果: {final_result}")
        results.append(final_result)

    # 7. 将结果写入新的 Excel 文件
    logger.log("所有IP处理完毕，正在生成结果文件...")
    host_info['CR'] = results

    # 生成更友好的输出文件名
    base_name = os.path.basename(file_path)
    dir_name = os.path.dirname(file_path)
    output_filename = os.path.splitext(base_name)[0] + f"_result_{str(int(time.time()))}.xlsx"
    output_path = os.path.join(dir_name, output_filename)

    try:
        host_info.to_excel(output_path, index=False)
        logger.log(f"结果已成功保存到: {output_path}")
        crt.Dialog.MessageBox(f"处理完成！\n\n结果已保存到文件:\n{output_path}")
    except Exception as e:
        logger.log(f"错误: 保存Excel文件失败: {str(e)}")
        crt.Dialog.MessageBox(f"错误: 保存结果文件失败。\n\n路径: {output_path}\n原因: {str(e)}")


main()