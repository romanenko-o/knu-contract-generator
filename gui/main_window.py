"""Графічний інтерфейс для створення контрактів і керування їхніми даними."""

import customtkinter as ctk
import os
import shutil
from tkinter import TclError
import tkinter.font as tkfont
from tkinter import filedialog, messagebox
from core.database import DataManager
from core.document_builder import fill_docx_template

# Налаштування зовнішнього вигляду
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

OPTIONMENU_STYLE = {
    "height": 24,
    "corner_radius": 6,
    "dynamic_resizing": False,
    "fg_color": ("gray95", "gray20"),
    "button_color": ("gray85", "gray30"),
    "button_hover_color": ("gray75", "gray40"),
    "text_color": ("black", "white"),
    "dropdown_fg_color": ("gray95", "gray15"),
    "dropdown_hover_color": "#8F9397",
    "dropdown_text_color": ("black", "white"),
    "font": ("Segoe UI", 15),
    "dropdown_font": ("Segoe UI", 15)
}

FONT_SIZE = 15
HEADER_FONT_SIZE = 18
TEXT_FONT_SIZE = 15
CARD_COLOR = ("gray95", "gray13")
BORDER_COLOR = ("gray85", "gray20")
UNIT_GENITIVE_CONTEXT = (
    "Фрагмент контракту: «Криворізький національний університет, "
    "в особі ректора [ПІБ у родовому відмінку], який діє на підставі Статуту...»"
)


def format_employee_short_name(employee_name: str) -> str:
    """Формує коротке ім'я працівника: ім'я та прізвище великими літерами."""
    name_parts = employee_name.split()
    if len(name_parts) < 2:
        return employee_name
    return f"{name_parts[1]} {name_parts[0].upper()}"


class ContractApp(ctk.CTk):
    """Головне вікно генератора контрактів і засоби керування його вкладками."""

    def __init__(self):
        """Створює вікно, підключає базу даних і налаштовує вкладки."""
        super().__init__()

        self.db = DataManager()
        position_rows = self.db.get_positions()
        self.positions_data = {
            name: (name_gen, name_ins)
            for name, name_gen, name_ins in position_rows
        }

        # Налаштування головного вікна
        self.title("Генератор контрактів")
        self.geometry("900x600")
        self.minsize(800, 600)

        self.tabview = ctk.CTkTabview(self)
        
        self.tabview.pack(fill="both", expand=True, padx=20, pady=20)

        # 3. Налаштовуємо розмір та шрифт шапки
        self.tabview._segmented_button.configure(
            font=ctk.CTkFont(size=15, weight="bold"),
            height=40,
            corner_radius=8
        )

        # Створюємо вкладки
        self.tab_create = self.tabview.add("  Створення контракту  ")
        self.tab_management = self.tabview.add("  Підписанти  ")
        self.tab_structure = self.tabview.add("  Структура  ")
        self.tab_batch = self.tabview.add("  Пакетне завантаження  ")

        # Викликаємо функції заповнення кожної вкладки
        self.setup_create_tab()
        self.setup_management_tab()
        self.setup_structure_tab()
        self._enable_clipboard_paste(self)

    def _paste_from_clipboard(self, event):
        """Вставляє текст із буфера обміну в активне текстове поле."""
        widget = event.widget
        widget_class = widget.winfo_class()
        if widget_class == "Entry" and widget.cget("state") != "normal":
            return "break"
        if widget_class == "Text" and widget.cget("state") == "disabled":
            return "break"

        try:
            text = widget.clipboard_get()
        except TclError:
            self.bell()
            return "break"

        if widget_class == "Entry":
            if widget.selection_present():
                widget.delete("sel.first", "sel.last")
        elif widget_class == "Text" and widget.tag_ranges("sel"):
            widget.delete("sel.first", "sel.last")

        widget.insert("insert", text)
        if widget is self.emp_info._textbox:
            self.after_idle(self._resize_employee_info)
        return "break"

    def _enable_clipboard_paste(self, parent):
        """Вмикає вставлення з буфера обміну для дочірніх полів введення."""
        for widget in parent.winfo_children():
            if isinstance(widget, ctk.CTkEntry):
                text_widget = widget._entry
            elif isinstance(widget, ctk.CTkTextbox):
                text_widget = widget._textbox
            elif isinstance(widget, ctk.CTkComboBox):
                text_widget = widget._entry
            else:
                self._enable_clipboard_paste(widget)
                continue

            text_widget.bind("<Control-KeyPress>", self._paste_from_control_key)
            text_widget.bind(
                "<Shift-KeyPress-Insert>", self._paste_from_clipboard
            )
            text_widget.bind(
                "<<Paste>>", self._paste_from_clipboard, add="+"
            )

    def _paste_from_control_key(self, event):
        """Обробляє комбінацію клавіш Ctrl+V у віджетах інтерфейсу."""
        if event.keysym.lower() == "v" or event.keycode == 86:
            return self._paste_from_clipboard(event)
        return None

    # ==========================================
    # ВКЛАДКА: Створення контракту
    # ==========================================
    def setup_create_tab(self):
        """Створює вкладку з формою підготовки нового контракту."""

        self.tab_create.grid_columnconfigure(0, weight=15)  # Лівий відступ
        self.tab_create.grid_columnconfigure(1, weight=70)  # Основний контент
        self.tab_create.grid_columnconfigure(2, weight=15)  # Правий відступ
    
        self.tab_create.grid_rowconfigure(0, weight=1)      # Розтягнення по висоті

        # Скрол-фрейм розміщуємо СТРОГО в 1-й (центральній) колонці
        scroll = ctk.CTkScrollableFrame(self.tab_create, fg_color="transparent")
        scroll.grid(row=0, column=1, sticky="nsew", pady=10)
        scroll.grid_columnconfigure(0, weight=1)

        # Кольори для "карток" залежно від теми (Світла/Темна)
        #CARD_COLOR = ("gray95", "gray13")
        #BORDER_COLOR = ("gray85", "gray20")

        # --- Шаблон документа ---
        ctk.CTkLabel(scroll, text="Шаблон документа", font=ctk.CTkFont(size=15, weight="bold")).grid(row=0, column=0, sticky="w", padx=10, pady=(10, 5))
        
        frame1 = ctk.CTkFrame(scroll, fg_color=CARD_COLOR, border_width=1, border_color=BORDER_COLOR, corner_radius=8)
        frame1.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 15))

        available_templates = self.get_available_templates()
        self.template_var = ctk.StringVar(value=available_templates[0])
        self.on_template_select(available_templates[0]) # Встановлюємо шлях для першого елемента

        self.template_menu = ctk.CTkOptionMenu(
            frame1, 
            values=available_templates,
            variable=self.template_var,
            command=self.on_template_select,
            width=250,
            **{**OPTIONMENU_STYLE, "corner_radius": 4}
        )
        self.template_menu.pack(side="left", padx=(20, 10), pady=20)

        self.btn_refresh_templates = ctk.CTkButton(
            frame1, 
            text="Оновити ⟳", 
            width=100, 
            fg_color=("gray90", "gray25"), 
            text_color=("black", "white"), 
            hover_color=("gray80", "gray35"),
            command=self.refresh_templates_list
        )
        self.btn_refresh_templates.pack(side="left", padx=10, pady=20)
        
        self.btn_select_template = ctk.CTkButton(
            frame1, text="Завантажити шаблон", 
            width=200, 
            fg_color=("gray90", "gray25"), 
            text_color=("black", "white"), 
            hover_color=("gray80", "gray35"),
            command=self.load_template_file
            )
        self.btn_select_template.pack(side="left", padx=20, pady=20)

        # Кнопка: Відкрити папку з шаблонами
        self.btn_open_folder = ctk.CTkButton(
            frame1, 
            text="Відкрити папку з шаблонами", 
            width=200, 
            fg_color=("gray90", "gray25"), 
            text_color=("black", "white"), 
            hover_color=("gray80", "gray35"),
            command=self.open_templates_folder
        )
        self.btn_open_folder.pack(side="left", padx=10, pady=20)

        # ---- Вибір підрозділу ----
        frame2 = ctk.CTkFrame(scroll, fg_color=CARD_COLOR, border_width=1, border_color=BORDER_COLOR, corner_radius=8)
        frame2.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 15))
        
        self.lbl_extra_option = ctk.CTkLabel(
            frame2, 
            text="Оберіть підрозділ:", 
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.lbl_extra_option.pack(side="left", padx=(20, 10), pady=20)

        units_list = self.db.get_organization_names() or ["Немає доступних підрозділів"]

        self.main_unit_combobox = ctk.CTkComboBox(
            frame2,
            values=units_list,
            width=self.calculate_combobox_width(units_list),
            command=self.on_unit_selected
        )
        self.main_unit_combobox.pack(side="left")

        # --- Основна інформація ---
        ctk.CTkLabel(scroll, text="Основна інформація", font=ctk.CTkFont(size=HEADER_FONT_SIZE, weight="bold")).grid(row=3, column=0, sticky="w", padx=10, pady=(10, 5))
        
        frame3 = ctk.CTkFrame(scroll, fg_color=CARD_COLOR, border_width=1, border_color=BORDER_COLOR, corner_radius=8)
        frame3.grid(row=4, column=0, sticky="ew", padx=10, pady=(0, 15))
        frame3.grid_columnconfigure(0, minsize=200, weight=0)
        frame3.grid_columnconfigure(1, weight=1)

        # ---- ПІБ ----
        ctk.CTkLabel(frame3, text="ПІБ працівника:", font=ctk.CTkFont(size=FONT_SIZE, weight="bold")).grid(row=0, column=0, sticky="e", padx=(20, 15), pady=(20, 10))
        self.emp_name = ctk.CTkEntry(frame3, fg_color="transparent", font=ctk.CTkFont(size=FONT_SIZE))
        self.emp_name.grid(row=0, column=1, sticky="ew", padx=(0, 20), pady=(20, 10))

        # --- Науковий ступінь ----
        ctk.CTkLabel(frame3, text="Науковий ступінь:", font=ctk.CTkFont(size=FONT_SIZE, weight="bold")).grid(row=1, column=0, sticky="e", padx=(20, 15), pady=10)
        self.emp_degree = ctk.CTkComboBox(
            frame3, 
            values=["Без ступеня", "Кандидат технічних наук", "Доктор наук"],
            font=ctk.CTkFont(size=FONT_SIZE)
        )
        self.emp_degree.set("Без ступеня")
        self.emp_degree.grid(row=1, column=1, sticky="ew", padx=(0, 20), pady=10)

        # --- Вчене звання ---
        ctk.CTkLabel(frame3, text="Вчене звання:", font=ctk.CTkFont(size=FONT_SIZE, weight="bold")).grid(row=2, column=0, sticky="e", padx=(20, 15), pady=10)
        self.emp_academic_title = ctk.CTkComboBox(
            frame3, 
            values=["Без звання", "Доцент", "Професор"],
            font=ctk.CTkFont(size=FONT_SIZE)
        )
        self.emp_academic_title.set("Без звання")
        self.emp_academic_title.grid(row=2, column=1, sticky="ew", padx=(0, 20), pady=10)

        # --- Посада ---
        ctk.CTkLabel(frame3, text="Посада:", font=ctk.CTkFont(size=FONT_SIZE, weight="bold")).grid(row=3, column=0, sticky="e", padx=(20, 15), pady=10)
        border_frame = ctk.CTkFrame(
            frame3,
            border_width=2,
            border_color="#838689",
            corner_radius=8,
            fg_color="transparent"
        )
        border_frame.grid(row=3, column=1, sticky="ew", padx=(0, 20), pady=10)

        self.emp_position = ctk.CTkOptionMenu(
            border_frame,
            values=list(self.positions_data),
            **{**OPTIONMENU_STYLE, "corner_radius": 4}  
        )
        self.emp_position.pack(fill="both", expand=True, padx=4, pady=4)

        # ---- Ставка ----
        ctk.CTkLabel(frame3, text="Ставка:", font=ctk.CTkFont(size=FONT_SIZE, weight="bold")).grid(row=4, column=0, sticky="e", padx=(20, 15), pady=10)
        self.emp_rate = ctk.CTkEntry(frame3, placeholder_text="напр. 1,0 або 0,75", fg_color="transparent", font=ctk.CTkFont(size=FONT_SIZE))
        self.emp_rate.grid(row=4, column=1, sticky="ew", padx=(0, 20), pady=10)

        # ---- Кафедра -----
        ctk.CTkLabel(frame3, text="Кафедра:", font=ctk.CTkFont(size=FONT_SIZE, weight="bold")).grid(row=5, column=0, sticky="e", padx=(20, 15), pady=(10, 20))
        
        selected_org = self.main_unit_combobox.get()
        departments_list = self.db.get_departments_list(selected_org)

        # Оновлюємо випадаючий список кафедр
        self.emp_department = ctk.CTkComboBox(
            frame3, 
            values=departments_list, 
            fg_color=("gray95", "gray20"), 
            font=ctk.CTkFont(size=FONT_SIZE),
            dropdown_font=ctk.CTkFont(size=FONT_SIZE),
            command=self.update_management_signatories
            )
        self.emp_department.set("")
        self.emp_department.grid(row=5, column=1, sticky="ew", padx=(0, 20), pady=(10, 20))

        # --- Термін дії ---
        ctk.CTkLabel(scroll, text="Термін дії", font=ctk.CTkFont(size=HEADER_FONT_SIZE, weight="bold")).grid(row=5, column=0, sticky="w", padx=10, pady=(10, 5))
        
        frame4 = ctk.CTkFrame(scroll, fg_color=CARD_COLOR, border_width=1, border_color=BORDER_COLOR, corner_radius=8)
        frame4.grid(row=6, column=0, sticky="ew", padx=10, pady=(0, 15))
        frame4.grid_columnconfigure((1, 3), weight=1)

        ctk.CTkLabel(frame4, text="Дата початку:", font=ctk.CTkFont(size=FONT_SIZE, weight="bold")).grid(row=0, column=0, sticky="e", padx=(20, 15), pady=20)
        self.start_date = ctk.CTkEntry(frame4, placeholder_text="29.08.2026", fg_color="transparent", font=ctk.CTkFont(size=FONT_SIZE))
        self.start_date.grid(row=0, column=1, sticky="ew", padx=(0, 20), pady=20)

        ctk.CTkLabel(frame4, text="Дата закінчення:", font=ctk.CTkFont(size=FONT_SIZE, weight="bold")).grid(row=0, column=2, sticky="e", padx=(20, 15), pady=20)
        self.end_date = ctk.CTkEntry(frame4, placeholder_text="30.07.2027", fg_color="transparent", font=ctk.CTkFont(size=FONT_SIZE))
        self.end_date.grid(row=0, column=3, sticky="ew", padx=(0, 20), pady=20)

        # --- Персональні та контактні дані ---
        ctk.CTkLabel(scroll, text="Персональні та контактні дані", font=ctk.CTkFont(size=HEADER_FONT_SIZE, weight="bold")).grid(row=7, column=0, sticky="w", padx=10, pady=(10, 5))
        
        frame5 = ctk.CTkFrame(scroll, fg_color=CARD_COLOR, border_width=1, border_color=BORDER_COLOR, corner_radius=8)
        frame5.grid(row=8, column=0, sticky="ew", padx=10, pady=(0, 15))
        frame5.grid_columnconfigure(0, weight=1)

        #ctk.CTkLabel(frame5, text="Паспортні дані:", font=ctk.CTkFont(size=FONT_SIZE, weight="bold")).grid(row=0, column=0, sticky="e", padx=(20, 15), pady=(20, 10))
        self.emp_info = ctk.CTkTextbox(
            frame5,
            height=60,
            wrap="word",
            font=ctk.CTkFont(size=FONT_SIZE),
        )
        self.emp_info.grid(row=0, column=0, sticky="ew", padx=20, pady=20)
        self.emp_info._textbox.bind(
            "<KeyRelease>", self._schedule_employee_info_resize, add="+"
        )
        self.emp_info._textbox.bind(
            "<<Paste>>", self._schedule_employee_info_resize, add="+"
        )
        self.emp_info._textbox.bind(
            "<<Cut>>", self._schedule_employee_info_resize, add="+"
        )
        self.emp_info._textbox.bind(
            "<Configure>", self._schedule_employee_info_resize, add="+"
        )
        

        # --- Кнопки дії (Низ сторінки) ---
        btn_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        btn_frame.grid(row=9, column=0, pady=25)

        self.btn_preview = ctk.CTkButton(btn_frame, text="Попередній перегляд", fg_color="gray", width=180, height=35, font=ctk.CTkFont(size=FONT_SIZE, weight="bold"))
        self.btn_preview.pack(side="left", padx=10)

        self.btn_create = ctk.CTkButton(btn_frame, text="Створити контракт", width=180, height=35, fg_color="#1f538d", font=ctk.CTkFont(size=FONT_SIZE, weight="bold"), command=self.create_contract)
        self.btn_create.pack(side="left", padx=10)

    def _schedule_employee_info_resize(self, _event=None):
        """Відкладає зміну висоти поля даних працівника до оновлення інтерфейсу."""
        if _event is not None and _event.widget is self.emp_info._textbox:
            self.after_idle(self._resize_employee_info)

    def _resize_employee_info(self):
        """Підбирає висоту текстового поля відповідно до кількості рядків."""
        text_widget = self.emp_info._textbox
        content_height = text_widget.count(
            "1.0", "end", "ypixels", "update"
        )
        new_height = max(60, content_height + 20)

        if self.emp_info.cget("height") != new_height:
            self.emp_info.configure(height=new_height)


    # ==============================================================
    # ВКЛАДКА: Підписанти (Керівництво)
    # ==============================================================

    def setup_management_tab(self):
        """Створює вкладку перегляду та редагування даних підписантів."""
        for widget in self.tab_management.winfo_children():
            widget.destroy()

        self.tab_management.grid_columnconfigure(0, weight=15)
        self.tab_management.grid_columnconfigure(1, weight=70)
        self.tab_management.grid_columnconfigure(2, weight=15)
        self.tab_management.grid_rowconfigure(0, weight=1)

        main_container = ctk.CTkFrame(self.tab_management, fg_color="transparent")
        main_container.grid(row=0, column=1, sticky="nsew", pady=10)

        ctk.CTkLabel(
            main_container, 
            text="Редагування загальноуніверситетського керівництва", 
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=(10, 20))

        scroll = ctk.CTkScrollableFrame(main_container, fg_color="transparent")
        scroll.pack(fill="both", expand=True, pady=5)

        # Заголовки стовпців
        header_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        header_frame.pack(fill="x", padx=5, pady=5)
        header_frame.grid_columnconfigure(0, weight=1)
        header_frame.grid_columnconfigure(1, weight=1)
        header_frame.grid_columnconfigure(2, weight=0)

        ctk.CTkLabel(header_frame, text="Посада", font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, sticky="w", padx=10)
        ctk.CTkLabel(header_frame, text="ПІБ Підписанта", font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=1, sticky="w", padx=10)

        # Загальне керівництво з БД
        signatories = self.db.get_all_signatories()
        
        for source_type, sig_id, role, full_name in signatories:
            row_frame = ctk.CTkFrame(scroll, fg_color=("gray95", "gray13"), corner_radius=6)
            row_frame.pack(fill="x", padx=5, pady=5)
            row_frame.grid_columnconfigure(0, weight=1)
            row_frame.grid_columnconfigure(1, weight=1)
            row_frame.grid_columnconfigure(2, weight=0)
            row_frame.grid_columnconfigure(3, weight=0)

            role_entry = ctk.CTkEntry(row_frame, fg_color="transparent", font=ctk.CTkFont(size=14))
            role_entry.insert(0, role)
            role_entry.configure(state="disabled")
            role_entry.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

            name_entry = ctk.CTkEntry(row_frame, fg_color="transparent", font=ctk.CTkFont(size=14))
            name_entry.insert(0, full_name)
            name_entry.configure(state="disabled")
            name_entry.grid(row=0, column=1, sticky="ew", padx=10, pady=10)
            saved_values = {"role": role, "name": full_name}

            btn_action = ctk.CTkButton(row_frame, text="РЕДАГУВАТИ", width=140)
            btn_cancel = ctk.CTkButton(
                row_frame,
                text="СКАСУВАТИ",
                width=110,
                fg_color="gray35",
                hover_color="gray45",
            )
            if source_type == "org":
                btn_action.configure(text="У СТРУКТУРІ", state="disabled")
                btn_cancel.destroy()
            else:
                def_fg, def_hv = btn_action.cget("fg_color"), btn_action.cget("hover_color")
                btn_action.configure(
                    command=lambda s_type=source_type, s_id=sig_id, r_e=role_entry, n_e=name_entry, b=btn_action, c=btn_cancel, saved=saved_values, fg=def_fg, hv=def_hv:
                    self.toggle_signatory_edit(
                        s_type, s_id, r_e, n_e, b, c, saved, fg, hv
                    )
                )
                btn_cancel.configure(
                    command=lambda r_e=role_entry, n_e=name_entry, b=btn_action, c=btn_cancel, saved=saved_values, fg=def_fg, hv=def_hv:
                    self.cancel_signatory_edit(
                        r_e, n_e, b, c, saved, fg, hv
                    )
                )
            btn_action.grid(row=0, column=2, padx=10, pady=10)
            if source_type != "org":
                btn_cancel.grid(row=0, column=3, padx=(5, 10), pady=10)
                btn_cancel.grid_remove()

        # 2. Рядок: Завідувач кафедри
        row_head = ctk.CTkFrame(scroll, fg_color=("gray95", "gray13"), corner_radius=6)
        row_head.pack(fill="x", padx=5, pady=5)
        row_head.grid_columnconfigure(0, weight=1)
        row_head.grid_columnconfigure(1, weight=1)
        row_head.grid_columnconfigure(2, weight=0)

        self.mgt_head_role_entry = ctk.CTkEntry(row_head, fg_color="transparent", font=ctk.CTkFont(size=14))
        self.mgt_head_role_entry.insert(0, "Завідувач кафедри")
        self.mgt_head_role_entry.configure(state="disabled")
        self.mgt_head_role_entry.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        self.mgt_head_entry = ctk.CTkEntry(row_head, fg_color="transparent", font=ctk.CTkFont(size=14), state="readonly")
        self.mgt_head_entry.grid(row=0, column=1, sticky="ew", padx=10, pady=10)

        self.btn_edit_head = ctk.CTkButton(row_head, text="РЕДАГУВАТИ", width=140)
        def_fg_h, def_hv_h = self.btn_edit_head.cget("fg_color"), self.btn_edit_head.cget("hover_color")
        self.btn_edit_head.configure(command=lambda: self.toggle_head_edit(def_fg_h, def_hv_h))
        self.btn_edit_head.grid(row=0, column=2, padx=10, pady=10)

        # 3. Рядок: Декан факультету
        row_dean = ctk.CTkFrame(scroll, fg_color=("gray95", "gray13"), corner_radius=6)
        row_dean.pack(fill="x", padx=5, pady=5)
        row_dean.grid_columnconfigure(0, weight=1)
        row_dean.grid_columnconfigure(1, weight=1)
        row_dean.grid_columnconfigure(2, weight=0)

        self.mgt_dean_role_entry = ctk.CTkEntry(row_dean, fg_color="transparent", font=ctk.CTkFont(size=14))
        self.mgt_dean_role_entry.insert(0, "Декан")
        self.mgt_dean_role_entry.configure(state="disabled")
        self.mgt_dean_role_entry.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        self.mgt_dean_entry = ctk.CTkEntry(row_dean, fg_color="transparent", font=ctk.CTkFont(size=14), state="readonly")
        self.mgt_dean_entry.grid(row=0, column=1, sticky="ew", padx=10, pady=10)

        self.btn_edit_dean = ctk.CTkButton(row_dean, text="РЕДАГУВАТИ", width=140)
        def_fg_d, def_hv_d = self.btn_edit_dean.cget("fg_color"), self.btn_edit_dean.cget("hover_color")
        self.btn_edit_dean.configure(command=lambda: self.toggle_dean_edit(def_fg_d, def_hv_d))
        self.btn_edit_dean.grid(row=0, column=2, padx=10, pady=10)

        # Первинне завантаження даних
        try:
            self.update_management_signatories()
        except AttributeError:
            pass

    # ==========================================
    # ВКЛАДКА: Структура (Факультети і Кафедри)
    # ==========================================

    def setup_structure_tab(self):
        """Створює вкладку керування підрозділами, факультетами й кафедрами."""
        for widget in self.tab_structure.winfo_children():
            widget.destroy()

        # Налаштування сітки для центрування контенту
        self.tab_structure.grid_columnconfigure(0, weight=1)
        self.tab_structure.grid_columnconfigure(1, weight=8)
        self.tab_structure.grid_columnconfigure(2, weight=1)
        self.tab_structure.grid_rowconfigure(0, weight=1)

        main_container = ctk.CTkScrollableFrame(self.tab_structure, fg_color="transparent")
        main_container.grid(row=0, column=1, sticky="nsew", pady=15, padx=10)

        # ------------------------------------------
        # ФРЕЙМ 1: Вибір елементів (Підрозділ / Факультет / Кафедра)
        # ------------------------------------------
        selection_frame = ctk.CTkFrame(main_container, fg_color=("gray95", "gray13"), corner_radius=8)
        selection_frame.pack(fill="x", pady=(0, 15))
        selection_frame.grid_columnconfigure(1, weight=1)

        # --- Рядок 1: Підрозділ (Організація) ---
        ctk.CTkLabel(selection_frame, text="Оберіть підрозділ:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, sticky="w", padx=15, pady=(15, 8))
        self.cmb_unit = ctk.CTkComboBox(selection_frame, command=self.on_unit_changed)
        self.cmb_unit.grid(row=0, column=1, sticky="ew", padx=(0, 10), pady=(15, 8))
        
        self.btn_edit_unit = ctk.CTkButton(selection_frame, text="✏️ Редагувати", width=110, command=lambda: self.start_editing("unit"))
        self.btn_edit_unit.grid(row=0, column=2, sticky="e", padx=(0, 15), pady=(15, 8))

        # --- Рядок 2: Факультет ---
        ctk.CTkLabel(selection_frame, text="Оберіть факультет:", font=ctk.CTkFont(weight="bold")).grid(row=1, column=0, sticky="w", padx=15, pady=8)
        self.cmb_faculty = ctk.CTkComboBox(selection_frame, command=self.on_faculty_changed)
        self.cmb_faculty.grid(row=1, column=1, sticky="ew", padx=(0, 10), pady=8)
        
        self.btn_edit_faculty = ctk.CTkButton(selection_frame, text="✏️ Редагувати", width=110, command=lambda: self.start_editing("faculty"))
        self.btn_edit_faculty.grid(row=1, column=2, sticky="e", padx=(0, 15), pady=8)

        # --- Рядок 3: Кафедра ---
        ctk.CTkLabel(selection_frame, text="Оберіть кафедру:", font=ctk.CTkFont(weight="bold")).grid(row=2, column=0, sticky="w", padx=15, pady=(8, 10))
        self.cmb_dept = ctk.CTkComboBox(selection_frame, command=self.on_dept_changed)
        self.cmb_dept.grid(row=2, column=1, sticky="ew", padx=(0, 10), pady=(8, 10))
        
        self.btn_edit_dept = ctk.CTkButton(selection_frame, text="✏️ Редагувати", width=110, command=lambda: self.start_editing("dept"))
        self.btn_edit_dept.grid(row=2, column=2, sticky="e", padx=(0, 15), pady=(8, 10))

        # --- Рядок 4: ЄДИНА кнопка ДОДАТИ ---
        self.btn_add_general = ctk.CTkButton(
            selection_frame, 
            text="➕ ДОДАТИ ЕЛЕМЕНТ", 
            fg_color="#10b981", 
            hover_color="#059669",
            height=36,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.open_add_choice_dialog
        )
        self.btn_add_general.grid(row=3, column=0, columnspan=3, sticky="w", padx=15, pady=(35, 15))

        # ------------------------------------------
        # ФРЕЙМ 2: Властивості та Редагування
        # ------------------------------------------
        self.editor_frame = ctk.CTkFrame(main_container, fg_color=("gray95", "gray13"), corner_radius=8)
        self.editor_frame.pack(fill="x", pady=5)

        self.active_editing_type = None
        self.render_editor_placeholder()

        # Первинне завантаження ланцюжка даних
        self.load_structure_data()

    # ==========================================
    # КАСКАДНЕ ЗАВАНТАЖЕННЯ ТА ЗВ'ЯЗОК З БД
    # ==========================================

    def load_structure_data(self):
        """Завантажує організації та відображає їх у списку вибору."""
        raw_orgs = self.db.get_organizations() or []
        # Значення містить ідентифікатор, назву, посаду та складові ПІБ керівника.
        self.orgs_data = {org[1]: org for org in raw_orgs}
        org_names = list(self.orgs_data.keys())

        if org_names:
            self.cmb_unit.configure(values=org_names)
            self.cmb_unit.set(org_names[0])
            self.on_unit_changed(org_names[0])
        else:
            self.cmb_unit.configure(values=["Немає підрозділів"])
            self.cmb_unit.set("Немає підрозділів")
            self.on_unit_changed(None)

    def on_unit_changed(self, selected_unit):
        """Завантажує факультети та кафедри вибраної організації."""
        organization = self.orgs_data.get(selected_unit)
        if not organization:
            self.faculties_data = {}
            self.departments_data = {}
            self.cmb_faculty.configure(values=["Немає факультетів"])
            self.cmb_faculty.set("Немає факультетів")
            self.cmb_dept.configure(values=["Немає кафедр"])
            self.cmb_dept.set("Немає кафедр")
            self.on_dept_changed(None)
            return

        raw_facs = self.db.get_faculties_by_organization(organization[0]) or []
        # Значення містить ідентифікатор, назву, посаду та складові ПІБ керівника.
        self.faculties_data = {fac[1]: fac for fac in raw_facs}
        faculty_names = list(self.faculties_data.keys())

        if faculty_names:
            self.cmb_faculty.configure(values=faculty_names)
            self.cmb_faculty.set(faculty_names[0])
            self.on_faculty_changed(faculty_names[0])
        else:
            self.cmb_faculty.configure(values=["Немає факультетів"])
            self.cmb_faculty.set("Немає факультетів")
            self.on_faculty_changed(None)

    def on_faculty_changed(self, selected_faculty):
        """Завантажує кафедри вибраного факультету."""
        fac_tuple = self.faculties_data.get(selected_faculty)
        
        if fac_tuple and len(fac_tuple) > 0:
            fac_id = fac_tuple[0]  # Отримуємо ідентифікатор факультету.
            raw_depts = self.db.get_departments_by_faculty(fac_id) or []
        else:
            organization = self.orgs_data.get(self.cmb_unit.get())
            raw_depts = (
                self.db.get_departments_without_faculty(organization[0])
                if organization else []
            )

        # Значення містить ідентифікатор, назву, посаду та складові ПІБ керівника.
        self.departments_data = {dept[1]: dept for dept in raw_depts}
        dept_names = list(self.departments_data.keys())

        if dept_names:
            self.cmb_dept.configure(values=dept_names)
            self.cmb_dept.set(dept_names[0])
            self.on_dept_changed(dept_names[0])
        else:
            self.cmb_dept.configure(values=["Немає кафедр"])
            self.cmb_dept.set("Немає кафедр")
            self.on_dept_changed(None)

    def on_dept_changed(self, selected_dept):
        """Оновлює форму редагування кафедри після зміни вибраного пункту."""
        if self.active_editing_type == "dept":
            self.start_editing("dept")

    # ==========================================
    # ФОРМА РЕДАГУВАННЯ (ФРЕЙМ 2)
    # ==========================================

    def render_editor_placeholder(self):
        """Показує підказку, якщо жоден елемент структури не редагується."""
        for w in self.editor_frame.winfo_children():
            w.destroy()
        self.active_editing_type = None

        lbl = ctk.CTkLabel(
            self.editor_frame, 
            text="ℹ️ Оберіть елемент вище та натисніть «Редагувати» для зміни його властивостей",
            font=ctk.CTkFont(size=13),
            text_color="gray50"
        )
        lbl.pack(pady=25)

    @staticmethod
    def _toggle_patronymic(entry, enabled_var):
        """Вмикає або вимикає поле по батькові відповідно до прапорця."""
        if enabled_var.get():
            entry.configure(state="normal")
        else:
            entry.configure(state="normal")
            entry.delete(0, "end")
            entry.configure(state="disabled")

    def _bind_uppercase_surname(self, entry):
        """Прив'язує перетворення прізвища на великі літери до подій поля."""
        entry.bind("<KeyRelease>", lambda _event: self._uppercase_entry(entry), add="+")
        entry.bind("<FocusOut>", lambda _event: self._uppercase_entry(entry), add="+")

    def _bind_capitalized_surname(self, entry):
        """Прив'язує форматування прізвища з великої літери до подій поля."""
        entry.bind("<KeyRelease>", lambda _event: self._capitalize_entry(entry), add="+")
        entry.bind("<FocusOut>", lambda _event: self._capitalize_entry(entry), add="+")

    @staticmethod
    def _capitalize_entry(entry):
        """Замінює першу літеру введеного значення на велику."""
        value = entry.get()
        capitalized_value = value.capitalize()
        if value != capitalized_value:
            cursor_position = entry.index("insert")
            entry.delete(0, "end")
            entry.insert(0, capitalized_value)
            entry.icursor(min(cursor_position, len(capitalized_value)))

    @staticmethod
    def _uppercase_entry(entry):
        """Перетворює все введене значення на великі літери."""
        value = entry.get()
        uppercase_value = value.upper()
        if value != uppercase_value:
            cursor_position = entry.index("insert")
            entry.delete(0, "end")
            entry.insert(0, uppercase_value)
            entry.icursor(min(cursor_position, len(uppercase_value)))

    def start_editing(self, entity_type: str):
        """Відкриває форму редагування вибраного елемента структури."""
        self.active_editing_type = entity_type
        
        for w in self.editor_frame.winfo_children():
            w.destroy()

        titles = {
            "unit": "підрозділу",
            "faculty": "факультету",
            "dept": "кафедри"
        }

        # 1. Витягуємо збережені дані з відповідного словника
        if entity_type == "unit":
            curr_name = self.cmb_unit.get()
            data = self.orgs_data.get(curr_name)
            if not data: self.render_editor_placeholder(); return
            e_id, e_name = data[0], data[1]
            e_role = data[2] or ""
            e_surname, e_first_name, e_patronymic = data[3:6]
            e_surname_gen, e_first_name_gen, e_patronymic_gen = data[6:9]

        elif entity_type == "faculty":
            curr_name = self.cmb_faculty.get()
            data = self.faculties_data.get(curr_name)
            if not data: self.render_editor_placeholder(); return
            e_id, e_name = data[0], data[1]
            e_role = data[2] or "Декан"
            e_surname, e_first_name, e_patronymic = data[3:6]
            e_name_gen = (data[6] or "") if len(data) > 6 else ""

        elif entity_type == "dept":
            curr_name = self.cmb_dept.get()
            data = self.departments_data.get(curr_name)
            if not data: self.render_editor_placeholder(); return
            e_id, e_name = data[0], data[1]
            e_role = data[2] or "Завідувач кафедри"
            e_surname, e_first_name, e_patronymic = data[3:6]

        self.current_editing_id = e_id  # Зберігаємо ідентифікатор для збереження або видалення.

        # 2. Заголовок
        lbl_title = ctk.CTkLabel(
            self.editor_frame, 
            text=f"Редагування {titles[entity_type]}: «{e_name}»", 
            font=ctk.CTkFont(size=14, weight="bold")
        )
        lbl_title.pack(anchor="w", padx=15, pady=(15, 10))

        # 3. Поля вводу
        form = ctk.CTkFrame(self.editor_frame, fg_color="transparent")
        form.pack(fill="x", padx=15, pady=5)
        form.grid_columnconfigure(1, weight=1)

        # Назва
        ctk.CTkLabel(form, text=f"Назва {titles[entity_type]}:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, sticky="w", pady=6, padx=(0, 10))
        self.entry_edit_name = ctk.CTkEntry(form)
        self.entry_edit_name.insert(0, e_name)
        self.entry_edit_name.grid(row=0, column=1, sticky="ew", pady=6)

        if entity_type == "faculty":
            ctk.CTkLabel(form, text="Назва факультету (родовий):").grid(
                row=1, column=0, sticky="w", pady=6, padx=(0, 10)
            )
            self.entry_edit_faculty_name_gen = ctk.CTkEntry(
                form, placeholder_text="Родовий відмінок назви факультету"
            )
            self.entry_edit_faculty_name_gen.insert(0, e_name_gen)
            self.entry_edit_faculty_name_gen.grid(
                row=1, column=1, sticky="ew", pady=6
            )

        # Посада керівника
        role_row = 2 if entity_type == "faculty" else 1
        ctk.CTkLabel(form, text="Посада керівника:").grid(row=role_row, column=0, sticky="w", pady=6, padx=(0, 10))
        self.entry_edit_role = ctk.CTkEntry(form, placeholder_text="Посада керівника")
        self.entry_edit_role.insert(0, e_role)
        self.entry_edit_role.grid(row=role_row, column=1, sticky="ew", pady=6)

        # Складові ПІБ керівника
        name_fields_start_row = role_row + 1
        name_fields = (
            ("Прізвище:", "entry_edit_surname", e_surname, "Прізвище"),
            ("Ім'я:", "entry_edit_first_name", e_first_name, "Ім'я"),
        )
        for row, (label, attribute, value, placeholder) in enumerate(name_fields, start=name_fields_start_row):
            ctk.CTkLabel(form, text=label).grid(row=row, column=0, sticky="w", pady=6, padx=(0, 10))
            entry = ctk.CTkEntry(form, placeholder_text=placeholder)
            entry.insert(0, value or "")
            if attribute == "entry_edit_surname":
                self._bind_uppercase_surname(entry)
            entry.grid(row=row, column=1, sticky="ew", pady=6)
            setattr(self, attribute, entry)
        patronymic_row = name_fields_start_row + 2
        self.edit_patronymic_enabled = ctk.BooleanVar(value=bool(e_patronymic))
        ctk.CTkCheckBox(
            form,
            text="Вказати по батькові",
            variable=self.edit_patronymic_enabled,
            command=lambda: self._toggle_patronymic(
                self.entry_edit_patronymic, self.edit_patronymic_enabled
            ),
        ).grid(row=patronymic_row, column=0, sticky="w", pady=6, padx=(0, 10))
        self.entry_edit_patronymic = ctk.CTkEntry(
            form,
            placeholder_text="По батькові (необов'язково)",
            state="normal" if e_patronymic else "disabled",
        )
        self.entry_edit_patronymic.insert(0, e_patronymic or "")
        self.entry_edit_patronymic.grid(row=patronymic_row, column=1, sticky="ew", pady=6)
        if entity_type == "unit":
            ctk.CTkLabel(
                form,
                text=UNIT_GENITIVE_CONTEXT,
                font=ctk.CTkFont(size=13),
                text_color="gray50",
                justify="left",
                wraplength=550,
            ).grid(row=5, column=0, columnspan=2, sticky="ew", pady=(8, 2))
            genitive_fields = (
                ("Прізвище (родовий):", "entry_edit_surname_gen", e_surname_gen),
                ("Ім'я (родовий):", "entry_edit_first_name_gen", e_first_name_gen),
                ("По батькові (родовий):", "entry_edit_patronymic_gen", e_patronymic_gen),
            )
            for row, (label, attribute, value) in enumerate(genitive_fields, start=6):
                ctk.CTkLabel(form, text=label).grid(
                    row=row, column=0, sticky="w", pady=6, padx=(0, 10)
                )
                entry = ctk.CTkEntry(form)
                entry.insert(0, value or "")
                if attribute == "entry_edit_surname_gen":
                    self._bind_capitalized_surname(entry)
                entry.grid(row=row, column=1, sticky="ew", pady=6)
                setattr(self, attribute, entry)

        # 4. Кнопки дій
        btn_box = ctk.CTkFrame(self.editor_frame, fg_color="transparent")
        btn_box.pack(fill="x", padx=15, pady=(15, 15))

        btn_delete = ctk.CTkButton(
            btn_box, 
            text="🗑️ Видалити", 
            fg_color="#ef4444", 
            hover_color="#dc2626", 
            command=self.delete_current_entity
        )
        btn_delete.pack(side="left")

        btn_cancel = ctk.CTkButton(
            btn_box, 
            text="Скасувати", 
            fg_color="gray35", 
            hover_color="gray45", 
            command=self.render_editor_placeholder
        )
        btn_cancel.pack(side="right", padx=(10, 0))

        btn_save = ctk.CTkButton(
            btn_box, 
            text="💾 Зберегти зміни", 
            fg_color="#10b981", 
            hover_color="#059669", 
            command=self.save_entity_changes
        )
        btn_save.pack(side="right")
        self._enable_clipboard_paste(self.editor_frame)

    # ==========================================
    # ЗБЕРЕЖЕННЯ ТА ВИДАЛЕННЯ З БД
    # ==========================================

    def save_entity_changes(self):
        """Перевіряє та зберігає зміни після підтвердження користувача."""
        new_name = self.entry_edit_name.get().strip()
        new_role = self.entry_edit_role.get().strip()
        head_surname = self.entry_edit_surname.get().strip()
        head_first_name = self.entry_edit_first_name.get().strip()
        head_patronymic = (
            self.entry_edit_patronymic.get().strip()
            if self.edit_patronymic_enabled.get() else ""
        )

        # Перевірка на порожнє поле
        if not new_name:
            messagebox.showwarning("Увага", "Назва не може бути порожньою!")
            return

        # ❓ Вікно підтвердження збереження
        confirm = messagebox.askyesno(
            title="Підтвердження збереження", 
            message="Ви дійсно бажаєте зберегти внесені зміни?"
        )
        if not confirm:
            return  # Якщо натиснули "Ні" — перериваємо операцію

        rec_id = self.current_editing_id
        success = False

        # Оновлення у БД
        if self.active_editing_type == "unit":
            success = self.db.update_organization(
                rec_id, new_name, new_role, head_surname, head_first_name,
                head_patronymic,
                self.entry_edit_surname_gen.get().strip().capitalize(),
                self.entry_edit_first_name_gen.get().strip(),
                self.entry_edit_patronymic_gen.get().strip(),
            )
            
        elif self.active_editing_type == "faculty":
            success = self.db.update_faculty(
                rec_id, new_name, new_role, head_surname, head_first_name,
                head_patronymic, self.entry_edit_faculty_name_gen.get().strip()
            )
            
        elif self.active_editing_type == "dept":
            success = self.db.update_department(
                rec_id, new_name, new_role, head_surname, head_first_name, head_patronymic
            )

        if not success:
            messagebox.showerror("Помилка", "Не вдалося зберегти зміни.")
            return
        messagebox.showinfo("Успіх", "Зміни успішно збережено!")

        # Оновлюємо інтерфейс
        self.load_structure_data()
        self.render_editor_placeholder()

    def delete_current_entity(self):
        """Запитує підтвердження та видаляє вибраний елемент структури."""
        rec_id = getattr(self, 'current_editing_id', None)
        if not rec_id: 
            return

        titles = {
            "unit": "підрозділ",
            "faculty": "факультет",
            "dept": "кафедру"
        }
        entity_label = titles.get(self.active_editing_type, "елемент")

        # ⚠️ Вікно підтвердження видалення з попередженням
        confirm = messagebox.askyesno(
            title="Підтвердження видалення", 
            message=f"Ви дійсно бажаєте видалити цей {entity_label}?\n\nЦю дію неможливо буде скасувати!",
            icon='warning'
        )
        if not confirm:
            return  # Якщо натиснули "Ні" — перериваємо операцію

        # Видаляємо запис через менеджер бази даних.
        delete_success = False
        if self.active_editing_type == "unit":
            delete_success = self.db.delete_organization(rec_id)
        elif self.active_editing_type == "faculty":
            delete_success = self.db.delete_faculty(rec_id)
        elif self.active_editing_type == "dept":
            delete_success = self.db.delete_department(rec_id)

        # ℹ️ Повідомлення про успішне видалення
        if delete_success:
            messagebox.showinfo("Успіх", f"Обраний {entity_label} успішно видалено!")
        else:
            messagebox.showerror("Помилка", f"Не вдалося видалити {entity_label}.")

        # Оновлюємо інтерфейс
        self.load_structure_data()
        self.render_editor_placeholder()

    def start_adding(self, entity_type: str):
        """Відкриває форму додавання нового елемента структури."""
        self.active_editing_type = f"add_{entity_type}"
        
        for w in self.editor_frame.winfo_children():
            w.destroy()

        titles = {
            "unit": "підрозділу",
            "faculty": "факультету",
            "dept": "кафедри"
        }

        # Перевірка наявності підрозділу, якщо додаємо факультет або кафедру
        if entity_type in ["faculty", "dept"] and not getattr(self, "orgs_data", None):
            messagebox.showwarning("Увага", "Спочатку додайте хоча б один підрозділ!")
            self.render_editor_placeholder()
            return

        lbl_title = ctk.CTkLabel(
            self.editor_frame, 
            text=f"Додавання: Новий/а {titles[entity_type]}", 
            font=ctk.CTkFont(size=14, weight="bold")
        )
        lbl_title.pack(anchor="w", padx=15, pady=(15, 10))

        form = ctk.CTkFrame(self.editor_frame, fg_color="transparent")
        form.pack(fill="x", padx=15, pady=5)
        form.grid_columnconfigure(1, weight=1)

        row_idx = 0

        # Вибір батьківського підрозділу (для факультетів та кафедр)
        if entity_type in ["faculty", "dept"]:
            ctk.CTkLabel(form, text="Підрозділ:").grid(row=row_idx, column=0, sticky="w", pady=6, padx=(0, 10))
            self.add_unit_combo = ctk.CTkComboBox(form, values=list(self.orgs_data.keys()), state="readonly")
            if self.cmb_unit.get() in self.orgs_data:
                self.add_unit_combo.set(self.cmb_unit.get())
            self.add_unit_combo.grid(row=row_idx, column=1, sticky="ew", pady=6)
            row_idx += 1

        # Вибір батьківського факультету (тільки для кафедр)
        if entity_type == "dept":
            ctk.CTkLabel(form, text="Факультет:").grid(row=row_idx, column=0, sticky="w", pady=6, padx=(0, 10))
            self.add_fac_combo = ctk.CTkComboBox(form, state="readonly")
            self.add_fac_combo.grid(row=row_idx, column=1, sticky="ew", pady=6)
            row_idx += 1
            
            # Динамічне оновлення списку факультетів при зміні підрозділу
            def update_faculties(selected_unit_name):
                """Оновлює список факультетів для вибраного підрозділу."""
                if selected_unit_name not in self.orgs_data: return
                org_id = self.orgs_data[selected_unit_name][0]
                facs = self.db.get_faculties_by_organization(org_id)
                fac_names = ["Без факультету"] + [f[1] for f in facs]
                self.add_fac_combo.configure(values=fac_names)
                
                # Підставляємо поточний вибраний факультет з головного меню, якщо він є
                current_fac = self.cmb_faculty.get()
                if current_fac in fac_names:
                    self.add_fac_combo.set(current_fac)
                else:
                    self.add_fac_combo.set("Без факультету")

            self.add_unit_combo.configure(command=update_faculties)
            update_faculties(self.add_unit_combo.get())

        # Назва
        ctk.CTkLabel(form, text=f"Назва {titles[entity_type]}:", font=ctk.CTkFont(weight="bold")).grid(row=row_idx, column=0, sticky="w", pady=6, padx=(0, 10))
        self.entry_add_name = ctk.CTkEntry(form, placeholder_text="Введіть назву")
        self.entry_add_name.grid(row=row_idx, column=1, sticky="ew", pady=6)
        row_idx += 1

        if entity_type == "faculty":
            ctk.CTkLabel(form, text="Назва факультету (родовий):").grid(
                row=row_idx, column=0, sticky="w", pady=6, padx=(0, 10)
            )
            self.entry_add_faculty_name_gen = ctk.CTkEntry(
                form, placeholder_text="Родовий відмінок назви факультету"
            )
            self.entry_add_faculty_name_gen.grid(
                row=row_idx, column=1, sticky="ew", pady=6
            )
            row_idx += 1

        # Посада керівника (встановлюємо значення за умовчанням)
        default_role = "Керівник підрозділу"
        if entity_type == "faculty": default_role = "Декан"
        elif entity_type == "dept": default_role = "Завідувач кафедри"

        ctk.CTkLabel(form, text="Посада керівника:").grid(row=row_idx, column=0, sticky="w", pady=6, padx=(0, 10))
        self.entry_add_role = ctk.CTkEntry(form)
        self.entry_add_role.insert(0, default_role)
        self.entry_add_role.grid(row=row_idx, column=1, sticky="ew", pady=6)
        row_idx += 1

        self.add_name_entries = {}
        for label, key, placeholder in (
            ("Прізвище:", "surname", "Прізвище керівника"),
            ("Ім'я:", "first_name", "Ім'я керівника"),
        ):
            ctk.CTkLabel(form, text=label).grid(row=row_idx, column=0, sticky="w", pady=6, padx=(0, 10))
            entry = ctk.CTkEntry(form, placeholder_text=placeholder)
            if key == "surname":
                self._bind_uppercase_surname(entry)
            entry.grid(row=row_idx, column=1, sticky="ew", pady=6)
            self.add_name_entries[key] = entry
            row_idx += 1

        self.add_patronymic_enabled = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            form,
            text="Вказати по батькові",
            variable=self.add_patronymic_enabled,
            command=lambda: self._toggle_patronymic(
                self.add_name_entries["patronymic"], self.add_patronymic_enabled
            ),
        ).grid(row=row_idx, column=0, sticky="w", pady=6, padx=(0, 10))
        patronymic_entry = ctk.CTkEntry(
            form, placeholder_text="По батькові (необов'язково)", state="disabled"
        )
        patronymic_entry.grid(row=row_idx, column=1, sticky="ew", pady=6)
        self.add_name_entries["patronymic"] = patronymic_entry
        if entity_type == "unit":
            row_idx += 1
            ctk.CTkLabel(
                form,
                text=UNIT_GENITIVE_CONTEXT,
                font=ctk.CTkFont(size=13),
                text_color="gray50",
                justify="left",
                wraplength=550,
            ).grid(row=row_idx, column=0, columnspan=2, sticky="ew", pady=(8, 2))
            row_idx += 1
            self.add_gen_name_entries = {}
            for label, key in (
                ("Прізвище (родовий):", "surname"),
                ("Ім'я (родовий):", "first_name"),
                ("По батькові (родовий):", "patronymic"),
            ):
                ctk.CTkLabel(form, text=label).grid(
                    row=row_idx, column=0, sticky="w", pady=6, padx=(0, 10)
                )
                entry = ctk.CTkEntry(form)
                if key == "surname":
                    self._bind_capitalized_surname(entry)
                entry.grid(row=row_idx, column=1, sticky="ew", pady=6)
                self.add_gen_name_entries[key] = entry
                row_idx += 1

        # 4. Кнопки
        btn_box = ctk.CTkFrame(self.editor_frame, fg_color="transparent")
        btn_box.pack(fill="x", padx=15, pady=(15, 15))

        btn_cancel = ctk.CTkButton(btn_box, text="Скасувати", fg_color="gray35", hover_color="gray45", command=self.render_editor_placeholder)
        btn_cancel.pack(side="right", padx=(10, 0))

        btn_save = ctk.CTkButton(btn_box, text="➕ Зберегти", fg_color="#10b981", hover_color="#059669", command=self.save_new_entity)
        btn_save.pack(side="right")
        self._enable_clipboard_paste(self.editor_frame)

    def save_new_entity(self):
        """Перевіряє та додає новий елемент структури до бази даних."""
        new_name = self.entry_add_name.get().strip()
        new_role = self.entry_add_role.get().strip()
        head_surname = self.add_name_entries["surname"].get().strip()
        head_first_name = self.add_name_entries["first_name"].get().strip()
        head_patronymic = (
            self.add_name_entries["patronymic"].get().strip()
            if self.add_patronymic_enabled.get() else ""
        )

        if not new_name:
            messagebox.showwarning("Увага", "Назва не може бути порожньою!")
            return

        success = False
        
        if self.active_editing_type == "add_unit":
            success = self.db.add_organization(
                new_name, new_role, head_surname, head_first_name, head_patronymic,
                self.add_gen_name_entries["surname"].get().strip().capitalize(),
                self.add_gen_name_entries["first_name"].get().strip(),
                self.add_gen_name_entries["patronymic"].get().strip(),
            )
            
        elif self.active_editing_type == "add_faculty":
            unit_name = self.add_unit_combo.get()
            if unit_name not in self.orgs_data:
                return
            org_id = self.orgs_data[unit_name][0]
            success = self.db.add_faculty(
                org_id, new_name, new_role, head_surname, head_first_name,
                head_patronymic, self.entry_add_faculty_name_gen.get().strip()
            )
            
        elif self.active_editing_type == "add_dept":
            unit_name = self.add_unit_combo.get()
            if unit_name not in self.orgs_data:
                return
            org_id = self.orgs_data[unit_name][0]
            
            fac_name = self.add_fac_combo.get()
            fac_id = None
            if fac_name != "Без факультету":
                # Знаходимо ідентифікатор вибраного факультету.
                facs = self.db.get_faculties_by_organization(org_id)
                for f in facs:
                    if f[1] == fac_name:
                        fac_id = f[0]
                        break
            
            success = self.db.add_department(
                org_id, fac_id, new_name, new_role, head_surname, head_first_name, head_patronymic
            )

        if success:
            messagebox.showinfo("Успіх", "Новий елемент успішно додано!")
            self.load_structure_data()
            self.render_editor_placeholder()
        else:
            messagebox.showerror("Помилка", "Не вдалося додати елемент. Можливо, така назва вже існує.")

# ========================================== #
# ============= МЕТОДИ ===================== #
# ========================================== #

    def load_template_file(self):
        """Копіює вибраний файл шаблону до каталогу шаблонів."""
        file_path = filedialog.askopenfilename(
            title="Оберіть шаблон контракту",
            filetypes=[("Word Documents", "*.doc *.docx"), ("All Files", "*.*")]
        )
        
        if not file_path:
            return

        try:
            templates_dir = os.path.join(os.getcwd(), "templates")
            os.makedirs(templates_dir, exist_ok=True)
            
            filename = os.path.basename(file_path)
            destination_path = os.path.join(templates_dir, filename)
            
            shutil.copy2(file_path, destination_path)
            self.refresh_templates_list(filename)
            self.current_template_path = destination_path
            
            messagebox.showinfo("Успіх", f"Шаблон '{filename}' успішно завантажено!")
            
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося завантажити файл:\n{e}")

    def open_templates_folder(self):
        """Відкриває каталог із шаблонами у файловому менеджері."""
        templates_dir = os.path.join(os.getcwd(), "templates")
        os.makedirs(templates_dir, exist_ok=True)
        try:
            os.startfile(templates_dir)
        except AttributeError:
            import subprocess, platform
            if platform.system() == "Darwin":
                subprocess.call(["open", templates_dir])
            else:
                subprocess.call(["xdg-open", templates_dir])
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося відкрити папку:\n{e}")

    def load_faculty_details(self, faculty_name):
        """Показує дані декана та кафедр вибраного факультету."""
        for widget in self.fac_edit_frame.winfo_children():
            widget.destroy()
        for widget in self.dept_scroll.winfo_children():
            widget.destroy()

        # ЗМІНА: Розпаковуємо 4 елементи кортежу факультету
        faculty_data = self.faculties_data[faculty_name]
        fac_id, fac_name, dean_role = faculty_data[:3]
        dean_name = self.db.format_head_name(*faculty_data[3:6])

        # --- Рендер редактора факультету ---
        self.fac_edit_frame.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(self.fac_edit_frame, text="Назва факультету:").grid(row=0, column=0, sticky="e", padx=10, pady=(10, 5))
        fac_name_entry = ctk.CTkEntry(self.fac_edit_frame, fg_color="transparent")
        fac_name_entry.insert(0, fac_name)
        fac_name_entry.configure(state="disabled")
        fac_name_entry.grid(row=0, column=1, sticky="ew", padx=10, pady=(10, 5))

        ctk.CTkLabel(self.fac_edit_frame, text="ПІБ Декана:").grid(row=1, column=0, sticky="e", padx=10, pady=(5, 10))
        dean_name_entry = ctk.CTkEntry(self.fac_edit_frame, fg_color="transparent")
        dean_name_entry.insert(0, dean_name or "")
        dean_name_entry.configure(state="disabled")
        dean_name_entry.grid(row=1, column=1, sticky="ew", padx=10, pady=(5, 10))

        btn_fac_action = ctk.CTkButton(self.fac_edit_frame, text="РЕДАГУВАТИ", width=150)
        def_fg_fac = btn_fac_action.cget("fg_color")
        def_hv_fac = btn_fac_action.cget("hover_color")

        btn_fac_action.configure(
            command=lambda f_id=fac_id, f_e=fac_name_entry, d_e=dean_name_entry, b=btn_fac_action, fg=def_fg_fac, hv=def_hv_fac: 
            self.toggle_faculty_edit(f_id, f_e, d_e, b, fg, hv)
        )
        btn_fac_action.grid(row=0, column=2, rowspan=2, padx=20)

        # --- Рендер редактора кафедр ---
        departments = self.db.get_departments_by_faculty(fac_id)
        
        # ЗМІНА: Розпаковуємо 4 елементи кортежу кафедри
        for dept_id, dept_name, head_role, surname, first_name, patronymic in departments:
            head_name = self.db.format_head_name(surname, first_name, patronymic)
            row_frame = ctk.CTkFrame(self.dept_scroll, fg_color=("gray95", "gray13"), corner_radius=6)
            row_frame.pack(fill="x", pady=5)
            row_frame.grid_columnconfigure(1, weight=1)

            ctk.CTkLabel(row_frame, text="Кафедра:").grid(row=0, column=0, sticky="e", padx=10, pady=(10, 5))
            dept_name_entry = ctk.CTkEntry(row_frame, fg_color="transparent")
            dept_name_entry.insert(0, dept_name)
            dept_name_entry.configure(state="disabled")
            dept_name_entry.grid(row=0, column=1, sticky="ew", padx=10, pady=(10, 5))

            ctk.CTkLabel(row_frame, text="Завідувач:").grid(row=1, column=0, sticky="e", padx=10, pady=(5, 10))
            head_name_entry = ctk.CTkEntry(row_frame, fg_color="transparent")
            head_name_entry.insert(0, head_name or "")
            head_name_entry.configure(state="disabled")
            head_name_entry.grid(row=1, column=1, sticky="ew", padx=10, pady=(5, 10))

            btn_dept_action = ctk.CTkButton(row_frame, text="РЕДАГУВАТИ", width=140)
            def_fg_dept = btn_dept_action.cget("fg_color")
            def_hv_dept = btn_dept_action.cget("hover_color")

            btn_dept_action.configure(
                command=lambda d_id=dept_id, dn_e=dept_name_entry, hn_e=head_name_entry, b=btn_dept_action, fg=def_fg_dept, hv=def_hv_dept: 
                self.toggle_department_edit(d_id, dn_e, hn_e, b, fg, hv)
            )
            btn_dept_action.grid(row=0, column=2, rowspan=2, padx=15)

    # --- Логіка перемикання станів і збереження ---

    def toggle_faculty_edit(self, fac_id, fac_entry, dean_entry, btn, def_fg, def_hv):
        """Перемикає редагування назви факультету й даних декана."""
        if btn.cget("text") == "РЕДАГУВАТИ":
            fac_entry.configure(state="normal")
            dean_entry.configure(state="normal")
            btn.configure(text="ЗБЕРЕГТИ ЗМІНИ", fg_color="#28a745", hover_color="#218838")
        else:
            if self.save_faculty(fac_id, fac_entry.get(), dean_entry.get()):
                fac_entry.configure(state="disabled")
                dean_entry.configure(state="disabled")
                btn.configure(text="РЕДАГУВАТИ", fg_color=def_fg, hover_color=def_hv)

    def toggle_department_edit(self, dept_id, dept_entry, head_entry, btn, def_fg, def_hv):
        """Перемикає редагування назви кафедри та даних її завідувача."""
        if btn.cget("text") == "РЕДАГУВАТИ":
            dept_entry.configure(state="normal")
            head_entry.configure(state="normal")
            btn.configure(text="ЗБЕРЕГТИ ЗМІНИ", fg_color="#28a745", hover_color="#218838")
        else:
            if self.save_department(dept_id, dept_entry.get(), head_entry.get()):
                dept_entry.configure(state="disabled")
                head_entry.configure(state="disabled")
                btn.configure(text="РЕДАГУВАТИ", fg_color=def_fg, hover_color=def_hv)

    def save_faculty(self, fac_id, faculty_name, dean_name):
        """Зберігає назву факультету та оновлені дані декана."""
        try:
            current = next(fac for fac in self.db.get_all_faculties() if fac[0] == fac_id)
            surname, first_name, patronymic = self.db._split_legacy_head_name(dean_name)
            self.db.update_faculty(
                fac_id, faculty_name, current[2], surname, first_name, patronymic
            )
            messagebox.showinfo("Успіх", "Дані факультету оновлено!")
            
            self.faculties_data = {fac[1]: fac for fac in self.db.get_all_faculties()}
            self.fac_combobox.configure(values=list(self.faculties_data.keys()))
            self.fac_combobox.set(faculty_name)
            return True
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося оновити факультет:\n{e}")
            return False

    def save_department(self, dept_id, dept_name, head_name):
        """Зберігає назву кафедри та оновлені дані завідувача."""
        try:
            current = next(
                (
                    dept for faculty in self.db.get_all_faculties()
                    for dept in self.db.get_departments_by_faculty(faculty[0])
                    if dept[0] == dept_id
                ),
                None,
            )
            if current is None:
                raise ValueError(f"Кафедру з ID {dept_id} не знайдено")
            surname, first_name, patronymic = self.db._split_legacy_head_name(head_name)
            self.db.update_department(
                dept_id, dept_name, current[2], surname, first_name, patronymic
            )
            messagebox.showinfo("Успіх", "Дані кафедри успішно оновлено!")
            return True
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося оновити кафедру:\n{e}")
            return False
        
    def toggle_signatory_edit(
        self, source_type, sig_id, role_entry, name_entry, btn, cancel_btn,
        saved_values, def_fg, def_hv
    ):
        """Перемикає режим редагування посади та імені підписанта."""
        if btn.cget("text") == "РЕДАГУВАТИ":
            role_entry.configure(state="normal")
            name_entry.configure(state="normal")
            btn.configure(text="ЗБЕРЕГТИ ЗМІНИ", fg_color="green", hover_color="darkgreen")
            cancel_btn.grid()
        else:
            new_role = role_entry.get().strip()
            new_name = name_entry.get().strip()
            
            if not new_role or not new_name:
                messagebox.showerror("Помилка", "Поля не можуть бути порожніми")
                return
                
            if not self.db.update_signatory(source_type, sig_id, new_role, new_name):
                messagebox.showerror("Помилка", "Не вдалося зберегти зміни підписанта.")
                return
            saved_values.update(role=new_role, name=new_name)
            
            role_entry.configure(state="disabled")
            name_entry.configure(state="disabled")
            btn.configure(text="РЕДАГУВАТИ", fg_color=def_fg, hover_color=def_hv)
            cancel_btn.grid_remove()

    @staticmethod
    def cancel_signatory_edit(
        role_entry, name_entry, btn, cancel_btn, saved_values, def_fg, def_hv
    ):
        """Скасовує редагування підписанта та відновлює попередні значення."""
        role_entry.configure(state="normal")
        name_entry.configure(state="normal")
        role_entry.delete(0, "end")
        role_entry.insert(0, saved_values["role"])
        name_entry.delete(0, "end")
        name_entry.insert(0, saved_values["name"])
        role_entry.configure(state="disabled")
        name_entry.configure(state="disabled")
        btn.configure(text="РЕДАГУВАТИ", fg_color=def_fg, hover_color=def_hv)
        cancel_btn.grid_remove()

    def save_signatory(self, sig_id, role, full_name):
        """Зберігає змінені посаду та ім'я підписанта."""
        try:
            self.db.update_signatory(sig_id, role, full_name)
            messagebox.showinfo("Успіх", "Дані підписанта успішно оновлено!")
            return True
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося оновити дані:\n{e}")
            return False
        
    def get_available_templates(self):
        """Сканує каталог шаблонів і повертає доступні файли DOC та DOCX."""
        templates_dir = os.path.join(os.getcwd(), "templates")
        os.makedirs(templates_dir, exist_ok=True)
        
        files = [f for f in os.listdir(templates_dir) if f.endswith(('.docx', '.doc'))]
        return files if files else ["Шаблони відсутні"]

    def refresh_templates_list(self, new_template=None):
        """Оновлює список шаблонів і вибирає потрібний файл."""
        templates = self.get_available_templates()
        self.template_menu.configure(values=templates)
        
        if new_template and new_template in templates:
            self.template_var.set(new_template)
            self.on_template_select(new_template)
        elif templates[0] != "Шаблони відсутні":
            self.template_var.set(templates[0])
            self.on_template_select(templates[0])
        else:
            self.template_var.set("Шаблони відсутні")
            self.current_template_path = None

    def on_template_select(self, selected_filename):
        """Зберігає шлях до обраного шаблону при зміні в меню"""
        if selected_filename == "Шаблони відсутні":
            self.current_template_path = None
            return
            
        template_path = os.path.join(os.getcwd(), "templates", selected_filename)
        if not os.path.isfile(template_path):
            template_path = os.path.join(os.getcwd(), selected_filename)
        self.current_template_path = template_path if os.path.isfile(template_path) else None

    def create_contract(self):
        """Збирає дані форми та створює контракт за вибраним шаблоном DOCX."""
        if not self.current_template_path:
            messagebox.showerror("Помилка", "Оберіть шаблон DOCX перед створенням контракту.")
            return
        if not self.current_template_path.lower().endswith(".docx"):
            messagebox.showerror("Помилка", "Для автозаповнення потрібен файл шаблону .docx.")
            return

        employee_name = self.emp_name.get().strip()
        department_name = self.emp_department.get().strip()
        if not employee_name or not department_name:
            messagebox.showerror("Помилка", "Заповніть ПІБ працівника та оберіть кафедру.")
            return

        department_info = self.db.get_department_info(department_name)
        signatories = {role.strip().lower(): name for _, _, role, name in self.db.get_all_signatories()}
        organization_name = self.main_unit_combobox.get().strip()
        organization_info = next((item for item in self.db.get_organizations() if item[1] == organization_name), None)
        head_name = (
            self.db.format_short_head_name(
                organization_info[3], organization_info[4]
            )
            if organization_info else ""
        )
        if organization_info:
            signatories[organization_info[2].strip().lower()] = self.db.format_head_name(
                organization_info[3], organization_info[4], organization_info[5]
            )

        leader_role = "Декан"
        leader_name = department_info.get("faculty_head", "")
        if organization_info and organization_info[1].strip() == "Навчально-науковий інститут економіки і торгівлі імені Михайла Туган-Барановського":
            leader_role = "Директор"
            leader_name = self.db.format_short_head_name(organization_info[3], organization_info[4])
        rector_short_name, rector_full_name, rector_full_name_gen = self.db.get_rector_names()
        rector_legacy_name = signatories.get("ректор", "")

        replacements = {
            "EMP_FULL_NAME": employee_name,
            "EMP_SHORT_NAME": format_employee_short_name(employee_name),
            "EMP_DEGREE": self.emp_degree.get().strip(),
            "ACADEMIC_RANK": self.emp_academic_title.get().strip(),
            "POSITION": self.emp_position.get().strip(),
            "POSITION_GEN": self.positions_data.get(self.emp_position.get().strip(), ("", ""))[0],
            "POSITION_ACC": self.positions_data.get(self.emp_position.get().strip(), ("", ""))[0],
            "POSITION_INS": self.positions_data.get(self.emp_position.get().strip(), ("", ""))[1],
            "POSITION_PREPOSITION": (
                "зі" if self.emp_position.get().strip() == "Старший викладач" else "з"
            ),
            "DEPARTMENT_NAME": department_name,
            "FACULTY_NAME": department_info.get("faculty_name", ""),
            "FACULTY_NAME_GEN": department_info.get("faculty_name_gen", ""),
            "FACULTY_HEAD_ROLE": leader_role,
            "FACULTY_HEAD_ROLE_GEN": leader_role,
            "WORK_RATE": self.emp_rate.get().strip(),
            "START_DATE": self.start_date.get().strip(),
            "END_DATE": self.end_date.get().strip(),
            "RECTOR_NAME": rector_short_name or rector_legacy_name,
            "RECTOR_NAME_GEN": rector_full_name_gen,
            "RECTOR_FULL_NAME": rector_full_name or rector_legacy_name,
            "RECTOR_FULL_NAME_GEN": rector_full_name_gen,
            "VICE_RECTOR_NAME": signatories.get("перший проректор", ""),
            "ACCOUNTANT_NAME": signatories.get("головний бухгалтер", ""),
            "METHODIST_HEAD_NAME": signatories.get("завідувач навчально-методичного відділу", ""),
            "LAW_HEAD_NAME": signatories.get("начальник юридичного відділу", ""),
            "PLAN_HEAD_NAME": signatories.get("начальник планового відділу", ""),
            "HR_HEAD_NAME": signatories.get("начальник відділу кадрів", ""),
            "DEAN_NAME": leader_name,
            "DEP_HEAD_NAME": department_info.get("department_head", ""),
            "HEAD_NAME": head_name,
        }
        output_path = filedialog.asksaveasfilename(
            title="Зберегти заповнений контракт",
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx")],
            initialfile=f"Контракт_{employee_name}.docx",
        )
        if not output_path:
            return

        try:
            fill_docx_template(self.current_template_path, output_path, replacements)
            messagebox.showinfo("Успіх", f"Контракт збережено:\n{output_path}")
            os.startfile(output_path)
        except Exception as error:
            messagebox.showerror("Помилка", f"Не вдалося створити контракт:\n{error}")

    def open_structure_editor(self):
        """Відкриває вікно для додавання та видалення елементів структури"""
        editor = ctk.CTkToplevel(self)
        editor.title("Управління структурою")
        editor.geometry("500x450")
        
        editor.transient(self) 
        editor.grab_set()      

        tabview = ctk.CTkTabview(editor)
        tabview.pack(fill="both", expand=True, padx=15, pady=15)
        
        tab_units = tabview.add("Підрозділи")
        tab_facs = tabview.add("Факультети")
        tab_depts = tabview.add("Кафедри")

        # ==========================================
        # Вкладка 1: Підрозділи
        # ==========================================
        ctk.CTkLabel(tab_units, text="➕ Додати новий підрозділ:", font=ctk.CTkFont(weight="bold")).pack(pady=(10,5))
        self.unit_add_entry = ctk.CTkEntry(tab_units, placeholder_text="Назва (напр. ННІЕТ)", width=350)
        self.unit_add_entry.pack(pady=5)
        ctk.CTkButton(tab_units, text="Додати підрозділ", fg_color="#28a745", hover_color="#218838").pack(pady=(5, 20))
        
        ctk.CTkLabel(tab_units, text="❌ Видалити підрозділ:", font=ctk.CTkFont(weight="bold")).pack(pady=(10,5))
        self.unit_del_combo = ctk.CTkComboBox(tab_units, values=["КНУ", "ННІЕТ"], width=350)
        self.unit_del_combo.pack(pady=5)
        ctk.CTkButton(tab_units, text="Видалити", fg_color="#dc3545", hover_color="#c82333").pack(pady=5)

        # ==========================================
        # Вкладка 2: Факультети
        # ==========================================
        ctk.CTkLabel(tab_facs, text="➕ Додати новий факультет:", font=ctk.CTkFont(weight="bold")).pack(pady=(10,5))
        self.fac_add_entry = ctk.CTkEntry(tab_facs, placeholder_text="Назва факультету", width=350)
        self.fac_add_entry.pack(pady=5)
        ctk.CTkButton(tab_facs, text="Додати факультет", fg_color="#28a745", hover_color="#218838").pack(pady=(5, 20))

        ctk.CTkLabel(tab_facs, text="❌ Видалити факультет:", font=ctk.CTkFont(weight="bold")).pack(pady=(10,5))
        self.fac_del_combo = ctk.CTkComboBox(tab_facs, values=list(self.faculties_data.keys()), width=350) 
        self.fac_del_combo.pack(pady=5)
        ctk.CTkButton(tab_facs, text="Видалити", fg_color="#dc3545", hover_color="#c82333").pack(pady=5)

        # ==========================================
        # Вкладка 3: Кафедри
        # ==========================================
        ctk.CTkLabel(tab_depts, text="➕ Додати нову кафедру:", font=ctk.CTkFont(weight="bold")).pack(pady=(10,5))
        self.dept_add_entry = ctk.CTkEntry(tab_depts, placeholder_text="Назва кафедри", width=350)
        self.dept_add_entry.pack(pady=5)
        ctk.CTkButton(tab_depts, text="Додати кафедру", fg_color="#28a745", hover_color="#218838").pack(pady=(5, 20))

        ctk.CTkLabel(tab_depts, text="❌ Видалити кафедру:", font=ctk.CTkFont(weight="bold")).pack(pady=(10,5))
        self.dept_del_combo = ctk.CTkComboBox(tab_depts, values=["Оберіть кафедру..."], width=350)
        self.dept_del_combo.pack(pady=5)
        ctk.CTkButton(tab_depts, text="Видалити", fg_color="#dc3545", hover_color="#c82333").pack(pady=5)

        self._enable_clipboard_paste(editor)

    def get_signatories_by_dept_name(self, dept_name):
        """Знайти Завідувача та Декана за назвою кафедри"""
        department_info = self.db.get_department_info(dept_name)
        if department_info:
            return department_info["department_head"], department_info["faculty_head"]
        return "", ""

    def update_dept_signatories(self, selected_dept_name):
        """Автоматично оновлює поля Завідувача та Декана при зміні кафедри"""
        head_name, dean_name = self.get_signatories_by_dept_name(selected_dept_name)

        if hasattr(self, 'emp_dept_head'):
            self.emp_dept_head.configure(state="normal")
            self.emp_dept_head.delete(0, "end")
            self.emp_dept_head.insert(0, head_name if head_name else "Не вказано")
            self.emp_dept_head.configure(state="readonly")

        if hasattr(self, 'emp_dean'):
            self.emp_dean.configure(state="normal")
            self.emp_dean.delete(0, "end")
            self.emp_dean.insert(0, dean_name if dean_name else "Не вказано")
            self.emp_dean.configure(state="readonly")

    def update_management_signatories(self, selected_dept_name=None):
        """Оновлює поля й ідентифікатори завідувача та керівника факультету."""
        if not selected_dept_name and hasattr(self, 'emp_department'):
            selected_dept_name = self.emp_department.get()

        head_name, head_role = "", "Завідувач кафедри"
        dean_name, dean_role = "", "Декан"
        self.current_dept_id, self.current_dept_name = None, None
        self.current_fac_id, self.current_fac_name = None, None

        if selected_dept_name:
            department_info = self.db.get_department_info(selected_dept_name)
            if department_info:
                head_name = department_info["department_head"]
                head_role = department_info["department_role"]
                dean_name = department_info["faculty_head"]
                dean_role = department_info["faculty_role"] or "Декан"
                self.current_dept_id = department_info["department_id"]
                self.current_dept_name = department_info["department_name"]
                self.current_fac_id = department_info["faculty_id"]
                self.current_fac_name = department_info["faculty_name"] or None

        # Оновлюємо поля у вкладці Підписанти (якщо вони не редагуються в цей момент)
        if hasattr(self, 'mgt_head_entry') and hasattr(self, 'mgt_head_role_entry'):
            if self.btn_edit_head.cget("text") == "РЕДАГУВАТИ":
                self.mgt_head_role_entry.configure(state="normal")
                self.mgt_head_role_entry.delete(0, "end")
                self.mgt_head_role_entry.insert(0, head_role)
                self.mgt_head_role_entry.configure(state="disabled")
                
                self.mgt_head_entry.configure(state="normal")
                self.mgt_head_entry.delete(0, "end")
                self.mgt_head_entry.insert(0, head_name)
                self.mgt_head_entry.configure(state="disabled")

        if hasattr(self, 'mgt_dean_entry') and hasattr(self, 'mgt_dean_role_entry'):
            if self.btn_edit_dean.cget("text") == "РЕДАГУВАТИ":
                self.mgt_dean_role_entry.configure(state="normal")
                self.mgt_dean_role_entry.delete(0, "end")
                self.mgt_dean_role_entry.insert(0, dean_role)
                self.mgt_dean_role_entry.configure(state="disabled")
                
                self.mgt_dean_entry.configure(state="normal")
                self.mgt_dean_entry.delete(0, "end")
                self.mgt_dean_entry.insert(0, dean_name)
                self.mgt_dean_entry.configure(state="disabled")

    def toggle_head_edit(self, def_fg, def_hv):
        """Перемикає режим редагування лише для ПОСАДИ завідувача кафедри"""
        if self.btn_edit_head.cget("text") == "РЕДАГУВАТИ":
            self.mgt_head_role_entry.configure(state="normal")
            self.mgt_head_role_entry.focus()
            self.btn_edit_head.configure(text="ЗБЕРЕГТИ ЗМІНИ", fg_color="#28a745", hover_color="#218838")
        else:
            self.mgt_head_role_entry.configure(state="disabled")
            self.btn_edit_head.configure(text="РЕДАГУВАТИ", fg_color=def_fg, hover_color=def_hv)

    def toggle_dean_edit(self, def_fg, def_hv):
        """Перемикає режим редагування лише для ПОСАДИ декана факультету"""
        if self.btn_edit_dean.cget("text") == "РЕДАГУВАТИ":
            self.mgt_dean_role_entry.configure(state="normal")
            self.mgt_dean_role_entry.focus()
            self.btn_edit_dean.configure(text="ЗБЕРЕГТИ ЗМІНИ", fg_color="#28a745", hover_color="#218838")
        else:
            self.mgt_dean_role_entry.configure(state="disabled")
            self.btn_edit_dean.configure(text="РЕДАГУВАТИ", fg_color=def_fg, hover_color=def_hv)


    def calculate_combobox_width(self, values: list[str], min_width: int = 150, padding: int = 10) -> int:
        """Обчислює ширину списку за найдовшим значенням і заданими відступами."""
        font = tkfont.Font(family="Segoe UI", size=13)
        max_text_px = max((font.measure(text) for text in values), default=100)
        return max(min_width, max_text_px + padding)

    def on_unit_selected(self, selected_org: str):
        """Оновлює список кафедр залежно від обраного підрозділу"""
        # 1. Отримуємо список кафедр для обраного підрозділу
        departments_list = self.db.get_departments_list(selected_org)

        # 2. Оновлюємо параметри наявного віджета.
        self.emp_department.configure(values=departments_list)

        # 3. Встановлюємо значення за умовчанням
        if departments_list:
            self.emp_department.set("")  # Або можна залишити ""
            # Одразу оновлюємо підписантів для першої кафедри
            self.update_management_signatories(departments_list[0])
        else:
            self.emp_department.set("Немає кафедр")
            self.update_management_signatories(None)

    def open_add_choice_dialog(self):
        """Вікно вибору елемента для додавання"""
        dialog = ctk.CTkToplevel(self)
        dialog.title("Додавання до структури")
        dialog.geometry("320x220")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        # Центрування вікна відносно головного
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - (320 // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (220 // 2)
        dialog.geometry(f"+{x}+{y}")

        ctk.CTkLabel(
            dialog, 
            text="Оберіть, що саме додати:", 
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(pady=(15, 10))

        def select_option(entity_type):
            """Починає додавання вибраного типу елемента структури."""
            dialog.destroy()
            self.start_adding(entity_type)

        ctk.CTkButton(
            dialog, 
            text="🏢 Підрозділ", 
            width=220, 
            height=32,
            fg_color="#10b981", 
            hover_color="#059669",
            command=lambda: select_option("unit")
        ).pack(pady=5)

        ctk.CTkButton(
            dialog, 
            text="🏛️ Факультет", 
            width=220, 
            height=32,
            fg_color="#10b981", 
            hover_color="#059669",
            command=lambda: select_option("faculty")
        ).pack(pady=5)

        ctk.CTkButton(
            dialog, 
            text="📚 Кафедру", 
            width=220, 
            height=32,
            fg_color="#10b981", 
            hover_color="#059669",
            command=lambda: select_option("dept")
        ).pack(pady=5)
