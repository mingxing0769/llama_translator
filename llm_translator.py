import datetime
import subprocess
import sys
import threading
import time
from queue import Queue

import pytesseract
import win32gui
from PIL import Image
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QColor
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QSizeGrip
from PyQt5.QtWidgets import QGraphicsDropShadowEffect
from llama_cpp import Llama
from mss import mss


class OCRApp:
    def __init__(self):
        self.last_text = ''
        self.text_data = Queue()
        self.previous_text = ''

    def get_window_pos_size(self):
        try:
            hwnd = win32gui.FindWindow(None, '实时辅助字幕')  # C:\\Windows\\System32\\livecaptions.exe窗口名称
            rect = win32gui.GetWindowRect(hwnd)
            left = rect[0]
            top = rect[1]
            width = rect[2] - left
            height = rect[3] - top
            return left, top, width, height
        except Exception as e:
            print(f"win32gui.FindWindow 发生错误: {e}")
            return None

    def screenshot(self):
        try:
            with mss() as sct:
                left, top, width, height = self.get_window_pos_size()
                monitor = {"top": top + 5, "left": left + 5, "width": width - 98, "height": height - 10}
                screenshot = sct.grab(monitor)
                return screenshot
        except Exception as e:
            print(f"screenshot 发生错误: {e}")
            print(f'left:{left},top:{top},width:{width},height:{height}')
            return None

    def Image_from_bytes(self):
        try:
            screenshot = self.screenshot()
            screenshot = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")
            return screenshot
        except Exception as e:
            print(f"Image_from_bytes 发生错误: {e}")
            return None

    def cut_sentences(self, text):
        """
        标点不在句尾时, 倒序查标点,找到的第一个标点切割输入.
        返回两个片段,第1个是标点前文本,第二个是标点后的文本.
        """

        # 定义句子结束的标点符号
        sentence_endings = ',.!?'

        # 从文本末尾开始向前搜索标点符号
        for i in range(len(text) - 1, -1, -1):
            if text[i] in sentence_endings:
                # 返回标点符号及其之前文本(含标点) 和 标点后的内容
                return text[:i + 1].strip(), text[i + 1:].strip()

    def merge_texts(self, text1, text2):
        """
        逐字去重合并两段文本
        """
        # 将文本分割成单词列表
        words1 = text1.split()
        words2 = text2.split()

        # 记录在text2中匹配的起始位置
        match_start = -1
        match_length = 0

        # 逐个单词进行比较
        for i in range(len(words1)):
            for j in range(len(words2)):
                # 如果找到匹配的第一个单词
                if words1[i] == words2[j]:
                    k = 1
                    # 继续比较接下来的单词
                    while i + k < len(words1) and j + k < len(words2) and words1[i + k] == words2[j + k]:
                        k += 1
                    # 如果匹配长度大于之前的记录，则更新匹配起始位置和长度
                    if k > match_length:
                        match_start = j
                        match_length = k

        # 如果找到了匹配部分，进行合并
        if match_start != -1:
            merged_text = ' '.join(words2[:match_start] + words1 + words2[match_start + match_length:])
        else:
            # 如果没有匹配部分，直接拼接
            merged_text = ' '.join(words1 + words2)

        return merged_text

    def process_text(self, current_input):
        """
        处理逻辑:
        检查输入中是否有标点
            如果有标点:
                是不是在结尾:
                    是在结尾:
                        返回标点前的一个完整句子,处理流程应该是:合并上次保存的文本和此次输入,后找到标点前的完整句子返回并保存
                    不是在结尾:
                        标点前是否还有另一个标点:
                            有:仅返回标点后的文本
                            无:上次保存的文本和此次输入合并去重,找到上一个完整句子和标点后的文本 返回并保存
            如果无标点:
                将上次保存的文本和此次输入合并去重,找到标点后,返回标点后的文本 并保存
        :param current_input: 待处理文本 字符串
        :return: 处理后的文本
        """

        # 定义标点
        sentence_endings = (',', '.', '!', '?')

        # 如果当前输入中没有标点,将上次保存和此次输入合并 并倒序找到标点后返回并保存标点后的内容
        if not any(char in current_input for char in sentence_endings):
            new_text = self.merge_texts(self.previous_text, current_input)
            if any(char in new_text for char in sentence_endings):
                _, end_text = self.cut_sentences(new_text)
                self.previous_text = end_text

                return self.previous_text
            else:
                self.previous_text = new_text
                return new_text

        else:
            # 如果当前输入中有标点(但不是结尾),
            # 检查标点前是否还有标点,
            #   如果没有其它标点  说明标点前的文本结构可能并不完整, 将当前输入文本和上次保存合并 倒序找到第二个标点后面的内容返回并保存标点后的内容
            if not current_input.endswith(('.', ',', '?', '!')):
                up_text, end_text = self.cut_sentences(current_input)
                # print('up_text1:', up_text)

                if not any(char in up_text[:len(up_text) - 1] for char in sentence_endings):  # 去掉up_text最后的标点后检查是否还有标点
                    new_text = self.merge_texts(self.previous_text, current_input)  # 将上次保存和当前输入后并，再倒序找第二个有标点。
                    # print("new_text:", new_text)

                    # 从文本开始向后搜索标点符号
                    for i in range(len(new_text)):
                        if new_text[i] in sentence_endings:
                            # 返回标点符号及其之后的所有内容
                            self.previous_text = new_text[i + 1:].strip()
                            return self.previous_text
                        else:  # 如果没有第二标点 则返回整个合并字符串
                            self.previous_text = new_text
                            return self.previous_text

                    # for i in range(len(new_text) - 1, -1, -1):
                    #     if new_text[i] in sentence_endings:
                    #         # 找到第一个标点符号后，继续向前搜索第二个标点符号
                    #         for j in range(i - 1, -1, -1):
                    #             if new_text[j] in sentence_endings:
                    #                 end_text = new_text[j + 1:].strip()
                    #
                    #                 self.previous_text = end_text
                    #                 return end_text
                    #             else:  # 如果没有第二标点 则返回整个合并字符串
                    #                 self.previous_text = new_text
                    #                 return self.previous_text

                else:  # 接上 如果有其它标点，说明第1个标点前的句子结构完整，直接返回第一个标点后的内容 并保存
                    self.previous_text = end_text
                    return end_text
            else:  # 如果是标点结尾,将上次保存和此次输入合并去重 返回标点前的完整句子 并保存
                new_text = self.merge_texts(self.previous_text, current_input)  # 合并上次保存的文本和此次输入的文本
                for i in range(len(new_text) - 1, -1, -1):
                    if new_text[i] in sentence_endings:
                        # 找到第一个标点符号后，继续向前搜索第二个标点符号
                        for j in range(i - 1, -1, -1):
                            if new_text[j] in sentence_endings:
                                end_text = new_text[j + 1:].strip()

                                # 检查 end_text字符长度，如果够长 则返回end_text 否则返回合并的文本
                                if len(end_text) > 50:
                                    self.previous_text = end_text
                                    return end_text
                                else:
                                    self.previous_text = current_input
                                    return new_text
                            else:
                                self.previous_text = current_input
                                return new_text
                    else:
                        self.previous_text = new_text
                        return self.previous_text

    def model_translates(self):
        messages = [{
            "role": "system",
            "content":
                r"""您是高级翻译助手，不要将输入的英文片段直译中文，请在之前的输入的基础上，正确处理重复词句后，了解到正确意思后意译为中文，过程中不要评论、解释。"""
        }]
        # 识别时错误字符串，及替换的字符串
        replacements = {
            '\n': ' '           
        }

        while True:
            start_dt = datetime.datetime.now()
            screenshot = self.Image_from_bytes()
            text = pytesseract.image_to_string(screenshot, lang='eng')

            for old, new in replacements.items():
                text = text.replace(old, new)

            if text and text != self.last_text:
                print('原文：', text)

                self.last_text = text

                # 对输入和前面的内容合并去重 输出最新内容 得重新处理 要启用去掉下面注释
                # text = self.process_text(text)
                # print('去重后的文本：', text)
                # print('存入:', self.previous_text)

                messages.append({"role": "user", "content": text})

                # 最大历史记录长度
                max_msg_length = 10
                # 如果messages长度超过最大值，则移除旧消息
                if len(messages) > max_msg_length:
                    # 保留系统提示信息
                    sys_msg = next((msg for msg in messages if msg["role"] == "system"),
                                   None)
                    # 保留最新的消息
                    messages = messages[-max_msg_length:]
                    if sys_msg:
                        messages.insert(0, sys_msg)

                completion = model.create_chat_completion(
                    messages,
                    max_tokens=128,
                    # temperature=0.5,  # 调高温度，增加生成内容的多样性
                    # top_p=0.9,  # 控制生成内容的概率分布
                    # frequency_penalty=0.2,  # 增加重复惩罚
                    # presence_penalty=0.1,  # 增加出现惩罚
                    # repeat_penalty=1.15,
                    stream=True
                )

                new_message = {"role": "assistant", "content": ""}

                for chunk in completion:
                    if 'content' in chunk['choices'][0]['delta']:
                        generated_content = chunk['choices'][0]['delta']['content']
                        print(generated_content, end="", flush=True)
                        new_message["content"] += generated_content
                        self.text_data.put(new_message["content"])

                messages.append(new_message)
                print('\n')

                end_dt = datetime.datetime.now()
                run_time = (end_dt - start_dt).total_seconds()
                print(f"运行时间：{run_time:.2f}秒 \n")
                if run_time < 1.5:
                    time.sleep(1.5 - run_time)
            else:
                time.sleep(2)

    def create_subtitle_window(self):
        app = QApplication([])
        window = SubtitleWindow(self.text_data)
        window.show()
        app.exec_()

    def main(self):
        try:
            model_translates_thread = threading.Thread(target=self.model_translates)
            model_translates_thread.daemon = True
            model_translates_thread.start()

        except SystemExit:
            print("主线程退出")
            sys.exit()

        self.create_subtitle_window()


class SubtitleWindow(QWidget):
    def __init__(self, text_data):
        super().__init__()

        self.oldPos = None
        self.text_data = text_data

        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setGeometry(360, 150, 1200, 150)  # 字幕位置 长宽

        layout = QVBoxLayout()
        self.setLayout(layout)

        self.chinese_label = QLabel()

        # 调整字体大小
        font_size = 13

        chinese_font = QFont("黑体", font_size + 7)  # 中文字体及大小

        # # 设置字体颜色和背景颜色
        font_color = "white"
        self.setStyleSheet(f"QLabel {{ color: {font_color}}}")

        # 创建阴影效果
        shadow_effect = QGraphicsDropShadowEffect()
        shadow_effect.setOffset(0, 1)
        shadow_effect.setColor(QColor('black'))
        shadow_effect.setBlurRadius(8)

        # 将阴影效果应用到标签上

        self.chinese_label.setGraphicsEffect(shadow_effect)

        self.chinese_label.setFont(chinese_font)

        self.chinese_label.setWordWrap(True)
        self.chinese_label.setMaximumWidth(1200)  # 字幕长度

        layout.addWidget(self.chinese_label)

        self.size_grip = QSizeGrip(self)
        layout.addWidget(self.size_grip)

        self.timer = QTimer()
        self.timer.timeout.connect(self.get_text)
        self.timer.start(0)

    def mousePressEvent(self, event):
        self.oldPos = event.globalPos()

    def mouseMoveEvent(self, event):
        delta = event.globalPos() - self.oldPos
        self.move(self.x() + delta.x(), self.y() + delta.y())
        self.oldPos = event.globalPos()

    def get_text(self):
        if not self.text_data.empty():
            local_C_text = self.text_data.get()
            self.chinese_label.setText(local_C_text)


if __name__ == "__main__":

    # 模型地址
    model_qwen2 = "E:/oobabooga_windows/models/Qwen2-7B_gguf/qwen2-7b-instruct-q8_0.gguf"
    model_glm = "E:/download/glm-4-9b-chat.Q6_K.gguf"
    model_internlm2_5 = "E:/download/internlm2_5-7b-chat-q8_0.gguf"
    model_phi3 = "E:/oobabooga_windows/models/Phi-3-medium-128k-instruct-Q6_K/Phi-3-medium-128k-instruct-Q6_K.gguf"
    model_qwen1_5 = "E:/download/qwen1_5-14b-chat-q5_k_m.gguf"
    model_deepseek = "E:/oobabooga_windows/models/Deepseek-coder_gguf/DeepSeek-Coder-V2-Lite-Instruct-IQ4_XS.gguf"

    # 加载模型
    model = Llama(model_path=model_qwen2,
                  verbose=False,
                  n_gpu_layers=-1,
                  n_ctx=1024 * 2,
                  flash_attn=True,
                  # chat_format='chatglm3'
                  )
    try:
        # 打开实时字幕软件
        subprocess.Popen(['C:\\Windows\\System32\\livecaptions.exe'])
        time.sleep(0.5)
        app = OCRApp()
        app.main()
    except Exception as e:
        print(f"程序运行出错: {e}")
        sys.exit(1)
