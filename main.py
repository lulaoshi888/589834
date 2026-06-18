"""
奇偶分析工具 - Kivy 版 (可打包 Android APK)
支持: .txt 及所有 Excel 格式 (.xlsx/.xls/.xlsm/.xlsb)
自动使用数据的最后8行
"""

import os
import sys

from kivy.app import App
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.lang import Builder
from kivy.metrics import dp
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView
from kivy.uix.textinput import TextInput
from kivy.utils import platform

# ───────────────────────── 分析逻辑（与原版完全一致） ─────────────────────────

def read_numbers_from_file(file_path):
    """读取文件中的全部数字（返回完整列表），由调用方取最后8行"""
    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".txt":
        with open(file_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
        numbers = []
        for line in lines:
            text = line.strip()
            if text:
                try:
                    numbers.append(int(text))
                except ValueError:
                    pass
    else:
        numbers = []
        try:
            import openpyxl
            wb = openpyxl.load_workbook(file_path, data_only=True)
            ws = wb.active
            for row in ws.iter_rows(min_col=2, max_col=2, values_only=True):
                val = row[0]
                if val is not None:
                    try:
                        numbers.append(int(val))
                    except (ValueError, TypeError):
                        pass
            wb.close()
        except ImportError:
            pass
        except Exception:
            try:
                import xlrd
                wb = xlrd.open_workbook(file_path)
                ws = wb.sheet_by_index(0)
                for i in range(ws.nrows):
                    val = ws.cell_value(i, 1)
                    if val != "":
                        try:
                            numbers.append(int(val))
                        except (ValueError, TypeError):
                            pass
            except ImportError:
                pass
            except Exception:
                return None

    if not numbers or len(numbers) < 8:
        return None
    return numbers


def analyze_numbers(numbers):
    """分析数字列表（8个数字），返回结果字典"""
    result = {}
    result["per_line"] = [(n, "奇数" if n % 2 == 1 else "偶数") for n in numbers]

    result["window_5"] = []
    for i in range(len(numbers) - 5 + 1):
        window = numbers[i:i + 5]
        odd_count = sum(1 for n in window if n % 2 == 1)
        result["window_5"].append((i + 1, i + 5, window, odd_count))

    result["window_4"] = []
    for i in range(len(numbers) - 4 + 1):
        window = numbers[i:i + 4]
        odd_count = sum(1 for n in window if n % 2 == 1)
        result["window_4"].append((i + 1, i + 4, window, odd_count))

    odd_3_6 = sum(1 for n in numbers[2:6] if n % 2 == 1)
    odd_3_7 = sum(1 for n in numbers[2:7] if n % 2 == 1)
    odd_4_7 = sum(1 for n in numbers[3:7] if n % 2 == 1)
    odd_4_8 = sum(1 for n in numbers[3:8] if n % 2 == 1)

    condition_met = (odd_3_6 == odd_3_7 and odd_4_7 == odd_4_8)
    result["condition"] = {
        "odd_3_6": odd_3_6,
        "odd_3_7": odd_3_7,
        "odd_4_7": odd_4_7,
        "odd_4_8": odd_4_8,
        "met": condition_met,
        "output": "奇数" if condition_met else "偶数"
    }
    return result


# ───────────────────────── Kivy UI ─────────────────────────

Builder.load_string("""
<FileDialogContent>:
    BoxLayout:
        orientation: 'vertical'
        spacing: dp(10)
        padding: dp(20)
        size_hint_y: None
        height: dp(400)
        Label:
            text: '选择数据文件'
            font_size: dp(16)
            bold: True
            size_hint_y: None
            height: dp(40)
        FileChooserListView:
            id: filechooser
            filters: ['*.txt', '*.xlsx', '*.xls', '*.xlsm', '*.xlsb']
            size_hint_y: None
            height: dp(280)
        BoxLayout:
            size_hint_y: None
            height: dp(50)
            spacing: dp(10)
            Button:
                text: '取消'
                on_release: root.dispatch('on_cancel')
            Button:
                text: '选择'
                on_release: root.dispatch('on_select')
                background_color: 0.2, 0.6, 1, 1

<ManualInputContent>:
    BoxLayout:
        orientation: 'vertical'
        spacing: dp(15)
        padding: dp(20)
        size_hint: (0.8, 0.4)
        Label:
            text: '请输入最新一行数据的数字：'
            font_size: dp(15)
            halign: 'center'
        TextInput:
            id: num_input
            multiline: False
            font_size: dp(24)
            halign: 'center'
            input_filter: 'int'
            input_type: 'number'
        BoxLayout:
            size_hint_y: None
            height: dp(50)
            spacing: dp(10)
            Button:
                text: '取消'
                on_release: root.dispatch('on_cancel')
            Button:
                text: '确认追加'
                on_release: root.dispatch('on_confirm')
                background_color: 0.15, 0.68, 0.38, 1
""")


class FileDialogContent(FloatLayout):
    """文件选择弹窗内容"""

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.register_event_type('on_select')
        self.register_event_type('on_cancel')

    def on_select(self):
        pass

    def on_cancel(self):
        pass


class ManualInputContent(FloatLayout):
    """手动输入弹窗内容"""

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.register_event_type('on_confirm')
        self.register_event_type('on_cancel')

    def on_confirm(self):
        pass

    def on_cancel(self):
        pass


class JudgeParityApp(App):
    """主应用"""

    def build(self):
        self.all_numbers = []
        self.title = "奇偶分析工具"
        Window.clearcolor = (0.94, 0.94, 0.94, 1)

        # 根布局
        root = BoxLayout(orientation='vertical', spacing=0)

        # ── 标题栏 ──
        title_bar = BoxLayout(
            orientation='vertical',
            size_hint_y=None,
            height=dp(50),
        )
        with title_bar.canvas.before:
            from kivy.graphics import Color, Rectangle
            Color(0.173, 0.243, 0.314, 1)
            self._title_bg = Rectangle(size=title_bar.size, pos=title_bar.pos)
            title_bar.bind(
                pos=lambda w, v: setattr(self._title_bg, 'pos', w.pos),
                size=lambda w, v: setattr(self._title_bg, 'size', w.size),
            )

        title_label = Label(
            text='奇偶分析工具',
            font_size=dp(18),
            bold=True,
            color=(1, 1, 1, 1),
            size_hint_y=1,
        )
        title_bar.add_widget(title_label)
        root.add_widget(title_bar)

        # ── 按钮区域 ──
        btn_row = BoxLayout(
            orientation='horizontal',
            size_hint_y=None,
            height=dp(50),
            padding=[dp(30), dp(8), dp(30), dp(8)],
            spacing=dp(15),
        )

        self.select_btn = Button(
            text='📁 选择数据文件',
            font_size=dp(13),
            background_color=(0.2, 0.6, 1, 1),
            color=(1, 1, 1, 1),
            size_hint_x=0.5,
        )
        self.select_btn.bind(on_release=self.open_file_dialog)

        self.manual_btn = Button(
            text='✏️ 手动输入',
            font_size=dp(13),
            background_color=(0.6, 0.35, 0.72, 1),
            color=(1, 1, 1, 1),
            size_hint_x=0.5,
        )
        self.manual_btn.bind(on_release=self.open_manual_input)

        btn_row.add_widget(self.select_btn)
        btn_row.add_widget(self.manual_btn)
        root.add_widget(btn_row)

        # ── 状态栏 ──
        self.status_label = Label(
            text='就绪 | 点击上方按钮选择文件',
            font_size=dp(11),
            color=(0.3, 0.3, 0.3, 1),
            size_hint_y=None,
            height=dp(28),
            halign='left',
            valign='middle',
            text_size=(Window.width - dp(20), dp(28)),
        )

        with self.status_label.canvas.before:
            Color(0.925, 0.945, 0.95, 1)
            self._status_bg = Rectangle(
                size=self.status_label.size, pos=self.status_label.pos
            )
            self.status_label.bind(
                pos=lambda w, v: setattr(self._status_bg, 'pos', w.pos),
                size=lambda w, v: setattr(self._status_bg, 'size', w.size),
            )

        root.add_widget(self.status_label)

        # ── 结果显示区域 ──
        scroll_view = ScrollView(
            size_hint=(1, 1),
            do_scroll_x=False,
            do_scroll_y=True,
            bar_width=dp(8),
        )

        self.result_label = Label(
            text='',
            font_size=dp(14),
            bold=True,
            color=(0.83, 0.83, 0.83, 1),
            size_hint_y=None,
            padding=[dp(16), dp(12)],
            markup=True,
            halign='left',
            valign='top',
        )
        self.result_label.bind(
            texture_size=lambda w, v: setattr(
                w, 'height', max(v[1], scroll_view.height)
            )
        )

        with self.result_label.canvas.before:
            Color(0.118, 0.118, 0.118, 1)
            self._result_bg = Rectangle(
                size=self.result_label.size, pos=self.result_label.pos
            )
            self.result_label.bind(
                pos=lambda w, v: setattr(self._result_bg, 'pos', w.pos),
                size=lambda w, v: setattr(self._result_bg, 'size', w.size),
            )

        scroll_view.add_widget(self.result_label)
        root.add_widget(scroll_view)

        # ── 自动加载默认文件 ──
        # Android 上没有默认文件，跳过；桌面端尝试加载
        if platform not in ('android', 'ios'):
            default_dir = os.path.dirname(os.path.abspath(__file__))
            default_file = os.path.join(default_dir, '文件.txt')
            if os.path.exists(default_file):
                Clock.schedule_once(lambda dt: self.load_file(default_file, '文件.txt'), 0.5)

        return root

    # ─────────── 文件处理 ───────────

    def load_file(self, file_path, display_name=None):
        """加载并分析文件"""
        name = display_name or os.path.basename(file_path)
        self.status_label.text = f'正在分析: {name} ...'

        numbers = read_numbers_from_file(file_path)
        if numbers is None:
            from kivy.uix.popup import Popup
            popup = Popup(
                title='错误',
                content=Label(text='文件未找到有效数据或数据不足8行！'),
                size_hint=(0.7, 0.3),
            )
            popup.open()
            self.status_label.text = '就绪 | 文件读取失败'
            return

        self.all_numbers = numbers[:]
        self.display_analysis(numbers, name)
        self.status_label.text = f'就绪 | 文件: {name}'

    def display_analysis(self, numbers, source='数据'):
        """显示分析结果"""
        numbers = numbers[-8:]
        result = analyze_numbers(numbers)
        cond = result['condition']

        lines = []
        lines.append(f'[b]📂 来源:[/b] {source}')
        lines.append(f'[b]📊 数据:[/b] {numbers} (最后8行)')
        lines.append('=' * 40)
        lines.append('')
        lines.append(f'[color=4ec9b0][b]🔮 预测结果: 【{cond["output"]}】[/b][/color]')

        # 详细数据
        lines.append('')
        lines.append('[b]逐行奇偶:[/b]')
        for i, (n, p) in enumerate(result['per_line'], 1):
            c = '#4ec9b0' if p == '奇数' else '#569cd6'
            lines.append(f'  {i}. {n} → [color={c}]{p}[/color]')

        lines.append('')
        lines.append(f'[b]条件分析:[/b]')
        lines.append(f'  第3-6行奇数数: {cond["odd_3_6"]}')
        lines.append(f'  第3-7行奇数数: {cond["odd_3_7"]}')
        lines.append(f'  第4-7行奇数数: {cond["odd_4_7"]}')
        lines.append(f'  第4-8行奇数数: {cond["odd_4_8"]}')
        status_text = '✅ 条件满足' if cond['met'] else '❌ 条件不满足'
        lines.append(f'  {status_text} → 预测: [color=4ec9b0][b]{cond["output"]}[/b][/color]')

        self.result_label.text = '\n'.join(lines)

    # ─────────── 弹窗 ───────────

    def open_file_dialog(self, btn):
        """打开文件选择对话框"""
        if platform in ('android', 'ios'):
            # 移动端：使用 plyer 的文件选择器
            try:
                from plyer import filechooser
                filechooser.open_file(
                    on_selection=self.handle_file_selection,
                    filters=[('支持的文件', '*.txt', '*.xlsx', '*.xls', '*.xlsm', '*.xlsb')],
                )
            except Exception:
                popup = Popup(
                    title='提示',
                    content=Label(
                        text='Android 上请将数据文件放入\n手机存储 / 目录后手动输入数据。',
                    ),
                    size_hint=(0.8, 0.4),
                )
                popup.open()
        else:
            # 桌面端：使用 FileChooser
            content = FileDialogContent()
            popup = Popup(title='选择数据文件', content=content, size_hint=(0.9, 0.8))
            content.bind(on_select=lambda _: self._on_filedialog_select(content, popup))
            content.bind(on_cancel=lambda _: popup.dismiss())
            popup.open()

    def _on_filedialog_select(self, content, popup):
        filechooser = content.ids.filechooser
        if filechooser.selection:
            fpath = filechooser.selection[0]
            ext = os.path.splitext(fpath)[1].lower()
            if ext not in ('.txt', '.xlsx', '.xls', '.xlsm', '.xlsb'):
                from kivy.uix.popup import Popup as KivyPopup
                error_popup = KivyPopup(
                    title='格式不支持',
                    content=Label(text='请选择 .txt 或 Excel 文件！'),
                    size_hint=(0.7, 0.3),
                )
                error_popup.open()
                return
            popup.dismiss()
            self.load_file(fpath)

    def handle_file_selection(self, selection):
        """移动端文件选择回调"""
        if selection:
            self.load_file(selection[0])

    def open_manual_input(self, btn):
        """打开手动输入对话框"""
        if not self.all_numbers:
            from kivy.uix.popup import Popup as KivyPopup
            warn = KivyPopup(
                title='提示',
                content=Label(text='请先选择数据文件作为基础数据！'),
                size_hint=(0.7, 0.3),
            )
            warn.open()
            return

        content = ManualInputContent()
        popup = Popup(
            title='输入最新数据',
            content=content,
            size_hint=(0.8, 0.35),
            auto_dismiss=False,
        )

        def do_confirm(_):
            raw = content.ids.num_input.text.strip()
            if not raw:
                from kivy.uix.popup import Popup as KivyPopup
                err = KivyPopup(
                    title='错误',
                    content=Label(text='请输入一个数字！'),
                    size_hint=(0.6, 0.25),
                )
                err.open()
                return
            try:
                new_num = int(raw)
            except ValueError:
                from kivy.uix.popup import Popup as KivyPopup
                err = KivyPopup(
                    title='错误',
                    content=Label(text=f'无法解析: {raw}'),
                    size_hint=(0.6, 0.25),
                )
                err.open()
                return

            popup.dismiss()
            self.all_numbers.append(new_num)
            self.display_analysis(self.all_numbers, f'文件 + 追加 ({new_num})')
            self.status_label.text = f'就绪 | 已追加: {new_num}'

        content.bind(on_confirm=do_confirm)
        content.bind(on_cancel=lambda _: popup.dismiss())

        # 绑定回车键
        content.ids.num_input.bind(
            on_text_validate=lambda w: do_confirm(None)
        )

        popup.open()
        # 自动弹出键盘
        Clock.schedule_once(lambda dt: content.ids.num_input.focus, 0.3)


if __name__ == '__main__':
    JudgeParityApp().run()
