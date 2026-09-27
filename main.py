import tkinter as tk
import os
import glob
import sys
from docx import Document
from docx.shared import Inches, Cm, Pt
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from tkinter import messagebox
from tkinter import ttk

from equipment.data import equipment
from prices import prices

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

window = tk.Tk()

window.title("Генератор КП")
window.geometry("700x700")

from PIL import Image, ImageTk

logo_path = os.path.join(BASE_DIR, "logo.png")

img = Image.open(logo_path)
img = img.resize((300, 130))   # можно менять размер

logo = ImageTk.PhotoImage(img)

logo_label = tk.Label(window, image=logo)
logo_label.pack(pady=15)

selected_equipment = []
equipment_vars = []

price_mode = tk.StringVar(value="Автоматически")

def replace_text(paragraph, company, price):

    replacements = {
        "{{company}}": company,
        "{{price}}": price,
    }

    for run in paragraph.runs:
        for old_text, new_text in replacements.items():
            if old_text in run.text:
                run.text = run.text.replace(old_text, new_text)

def choose_equipment():
    equipment_window = tk.Toplevel(window)

    equipment_window.title("Навесное оборудование")
    equipment_window.geometry("600x600")

    tk.Label(
        equipment_window,
        text="Выберите навесное оборудование",
        font=("Arial", 12, "bold")
    ).pack(pady=10)

    # Область со списком
    list_frame = tk.Frame(equipment_window)
    list_frame.pack(
        fill="both",
        expand=True,
        padx=10,
        pady=10
    )

    # Canvas
    canvas = tk.Canvas(list_frame)

    # Полоса прокрутки
    scrollbar = tk.Scrollbar(
        list_frame,
        orient="vertical",
        command=canvas.yview
    )

    # Внутренний Frame
    scrollable_frame = tk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window(
        (0, 0),
        window=scrollable_frame,
        anchor="nw"
    )

    canvas.configure(
        yscrollcommand=scrollbar.set
    )

    canvas.bind_all(
    "<MouseWheel>",
    lambda event: canvas.yview_scroll(
        int(-1 * (event.delta / 120)),
        "units"
    )
)

    canvas.pack(
        side="left",
        fill="both",
        expand=True
    )

    scrollbar.pack(
        side="right",
        fill="y"
    )

    equipment_vars.clear()

    for item in equipment:
        var = tk.BooleanVar()

        checkbox = tk.Checkbutton(
            scrollable_frame,
            text=f'{item["name"]} — {item["price"]:,} ₸',
            variable=var,
            anchor="w"
        )

        checkbox.pack(
            fill="x",
            padx=10,
            pady=4
        )

        equipment_vars.append(
            (var, item)
        )

    def save_equipment():

        selected_equipment.clear()

        for var, item in equipment_vars:
            if var.get():
                selected_equipment.append(item)

        update_equipment_label()

        equipment_window.destroy()

        canvas.unbind_all("<MouseWheel>")

    # КНОПКА
    save_button = tk.Button(
        equipment_window,
        text="ГОТОВО",
        command=save_equipment,
        width=20
    )

    save_button.pack(
        pady=10
    )

def update_equipment_label():
    if not selected_equipment:
        equipment_label.config(
            text="Оборудование не выбрано"
        )
        update_price()
        return

    names = ", ".join(
        item["name"]
        for item in selected_equipment
    )

    equipment_label.config(
        text=f"Выбрано:\n{names}",
        wraplength=500
    )

    update_price()

def update_price():
    if price_mode.get() == "Автоматически":

        template = template_box.get()
        base_price = prices.get(template, 0)

        equipment_price = sum(
            item["price"]
            for item in selected_equipment
        )

        total = base_price + equipment_price

        price_entry.config(state="normal")
        price_entry.delete(0, tk.END)
        price_entry.insert(0, f"{total:,} ₸")
        price_entry.config(state="readonly")

    else:

        price_entry.config(state="normal")
        price_entry.delete(0, tk.END)

def add_equipment_block(doc):
    if not selected_equipment:
        return

    marker = "{{EQUIPMENT_BLOCK}}"

    for paragraph in doc.paragraphs:

        if marker not in paragraph.text:
            continue

        # Очищаем маркер
        paragraph.text = ""

        # Заголовок
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        heading = paragraph.add_run("НАВЕСКИ В КОМПЛЕКТЕ")
        heading.bold = True
        heading.font.size = Pt(11)

        # Создаём таблицу
        table = doc.add_table(
            rows=len(selected_equipment),
            cols=2
        )

        table.autofit = False

        # Принудительно задаём ширину
        for row in table.rows:

            row.cells[0].width = Cm(8.75)
            row.cells[1].width = Cm(8.75)

        # Заполняем таблицу
        for i, item in enumerate(selected_equipment):

            row = table.rows[i]

            name_cell = row.cells[0]
            image_cell = row.cells[1]

            # Название
            name_cell.text = item["name"]

            name_paragraph = name_cell.paragraphs[0]
            name_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

            # Фото
            image_path = item.get("image")

            if image_path and os.path.exists(image_path):

                image_paragraph = image_cell.paragraphs[0]
                image_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

                run = image_paragraph.add_run()

                run.add_picture(
                    image_path,
                    width=Cm(4.5)
                )

            # Центрирование по вертикали
            name_cell.vertical_alignment = (
                WD_CELL_VERTICAL_ALIGNMENT.CENTER
            )

            image_cell.vertical_alignment = (
                WD_CELL_VERTICAL_ALIGNMENT.CENTER
            )

        # ==========================================
        # ДЕЛАЕМ ГРАНИЦЫ ТАБЛИЦЫ
        # ==========================================

        tbl = table._tbl
        tblPr = tbl.tblPr

        borders = tblPr.first_child_found_in("w:tblBorders")

        if borders is None:
            from docx.oxml import OxmlElement

            borders = OxmlElement("w:tblBorders")
            tblPr.append(borders)

        from docx.oxml import OxmlElement

        for border_name in (
            "top",
            "left",
            "bottom",
            "right",
            "insideH",
            "insideV"
        ):

            border = borders.find(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
                + border_name
            )

            if border is None:
                border = OxmlElement(
                    f"w:{border_name}"
                )
                borders.append(border)

            border.set(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val",
                "single"
            )

            border.set(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz",
                "8"
            )

            border.set(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}space",
                "0"
            )

            border.set(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color",
                "000000"
            )

        # ==========================================
        # ВСТАВЛЯЕМ ТАБЛИЦУ НА МЕСТО МАРКЕРА
        # ==========================================
        
        from docx.oxml import OxmlElement

        tblPr = table._tbl.tblPr

        tblInd = OxmlElement("w:tblInd")

        tblInd.set(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w",
            "600"
            )
        tblInd.set(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type",
            "dxa"
            )
        tblPr.append(tblInd)

        paragraph._p.addnext(table._tbl)

        return

    messagebox.showerror(
        "Ошибка",
        "В шаблоне не найден {{EQUIPMENT_BLOCK}}"
    )

def create_document():

    company = company_entry.get()
    price = price_entry.get()

    if not company or not price:
        messagebox.showerror(
            "Ошибка",
            "Заполните все поля!"
        )
        return

    template = template_box.get()

    # Путь к выбранному шаблону
    doc_path = os.path.join(
        BASE_DIR,
        "templates",
        f"{template}.docx"
    )

    if not os.path.exists(doc_path):
        messagebox.showerror(
            "Ошибка",
            f"Не найден шаблон:\n{doc_path}"
        )
        return

    doc = Document(doc_path)

    # Добавляем выбранные навески
    add_equipment_block(doc)

    # Заменяем данные в обычном тексте
    for paragraph in doc.paragraphs:
        replace_text(
            paragraph,
            company,
            price
        )

    # Заменяем данные в таблицах
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_text(
                        paragraph,
                        company,
                        price
                    )

    # Создаём папку output, если её нет
    output_dir = os.path.join(
        BASE_DIR,
        "output"
    )

    os.makedirs(
        output_dir,
        exist_ok=True
    )

    # Безопасное имя файла
    safe_company = company

    for char in '\\/:*?"<>|':
        safe_company = safe_company.replace(
            char,
            ""
        )

    path = os.path.join(
        output_dir,
        f"КП_{safe_company}.docx"
    )

    try:
        doc.save(path)

        os.startfile(path)

        messagebox.showinfo(
            "Готово!",
            "Коммерческое предложение успешно создано!"
        )

        company_entry.delete(
            0,
            tk.END
        )

    except Exception as error:
        messagebox.showerror(
            "Ошибка",
            f"Не удалось создать КП:\n\n{error}"
        )

templates = glob.glob("templates/*.docx")

template_names = []

for file in templates:
    name = os.path.basename(file)
    name = name.replace(".docx", "")
    template_names.append(name)

# Шаблон КП
tk.Label(window, text="Шаблон КП").pack()



template_box = ttk.Combobox(
    window,
    values=template_names,
    state="readonly",
    width=37
)

template_box.current(0)
template_box.pack(pady=5)

template_box.bind(
    "<<ComboboxSelected>>",
    lambda event: update_price()
)

tk.Label(
    window,
    text="Навесное оборудование"
).pack(pady=(10, 0))

equipment_button = tk.Button(
    window,
    text="Выбрать оборудование",
    command=choose_equipment
)

equipment_button.pack(pady=5)

equipment_label = tk.Label(
    window,
    text="Оборудование не выбрано",
    wraplength=450
)

equipment_label.pack(pady=5)

# Компания
tk.Label(window, text="Компания").pack()
company_entry = tk.Entry(window, width=40)
company_entry.pack(pady=5)

tk.Label(
    window,
    text="Расчёт стоимости"
).pack(pady=(10, 0))

price_mode_box = ttk.Combobox(
    window,
    textvariable=price_mode,
    values=["Автоматически", "Вручную"],
    state="readonly",
    width=37
)

price_mode_box.pack(pady=5)

price_mode_box.bind(
    "<<ComboboxSelected>>",
    lambda event: update_price()
)

tk.Label(
    window,
    text="Итоговая стоимость"
).pack()

price_entry = tk.Entry(
    window,
    width=40
)

price_entry.pack(pady=5)

button = tk.Button(
    window,
    text="Создать КП",
    command=create_document
)
button.pack(pady=20)

window.mainloop()