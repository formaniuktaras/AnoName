import os
import tkinter as tk
from tkinter import filedialog, messagebox

from anonymizer import process_docx
from generators import generator_full_random, generator_init_random


class AnonApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("DocxAnon")
        self.root.resizable(False, False)

        icon_path = os.path.join(os.path.dirname(__file__), "assets", "icon.ico")
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except Exception:
                pass

        self.input_path_var = tk.StringVar(value="")
        self.output_dir_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Очікування")

        self.replace_full_var = tk.BooleanVar(value=True)
        self.replace_initials_var = tk.BooleanVar(value=True)
        self.process_tables_var = tk.BooleanVar(value=True)
        self.process_headers_var = tk.BooleanVar(value=True)

        self._build_ui()

    def _build_ui(self) -> None:
        frame = tk.Frame(self.root, padx=12, pady=12)
        frame.grid(row=0, column=0, sticky="nsew")

        file_button = tk.Button(frame, text="Обрати DOCX", command=self.select_file)
        file_button.grid(row=0, column=0, sticky="w")

        file_entry = tk.Entry(frame, textvariable=self.input_path_var, width=60, state="readonly")
        file_entry.grid(row=0, column=1, padx=(8, 0), sticky="w")

        output_label = tk.Label(frame, text="Папка збереження:")
        output_label.grid(row=1, column=0, pady=(10, 0), sticky="w")

        output_entry = tk.Entry(frame, textvariable=self.output_dir_var, width=60)
        output_entry.grid(row=1, column=1, padx=(8, 0), pady=(10, 0), sticky="w")

        output_button = tk.Button(frame, text="Обрати папку", command=self.select_output_dir)
        output_button.grid(row=2, column=1, pady=(6, 0), sticky="w")

        options_frame = tk.LabelFrame(frame, text="Опції", padx=10, pady=8)
        options_frame.grid(row=3, column=0, columnspan=2, pady=(12, 0), sticky="ew")

        tk.Checkbutton(
            options_frame,
            text="Замінювати повні ПІБ",
            variable=self.replace_full_var,
        ).grid(row=0, column=0, sticky="w")

        tk.Checkbutton(
            options_frame,
            text="Замінювати ПІБ з ініціалами",
            variable=self.replace_initials_var,
        ).grid(row=1, column=0, sticky="w")

        tk.Checkbutton(
            options_frame,
            text="Обробляти таблиці",
            variable=self.process_tables_var,
        ).grid(row=2, column=0, sticky="w")

        tk.Checkbutton(
            options_frame,
            text="Обробляти headers / footers",
            variable=self.process_headers_var,
        ).grid(row=3, column=0, sticky="w")

        start_button = tk.Button(
            frame,
            text="Почати анонімізацію",
            command=self.start_anonymization,
            width=25,
        )
        start_button.grid(row=4, column=0, columnspan=2, pady=(12, 0))

        status_bar = tk.Label(
            frame,
            textvariable=self.status_var,
            anchor="w",
            relief="sunken",
            width=70,
        )
        status_bar.grid(row=5, column=0, columnspan=2, pady=(12, 0), sticky="ew")

    def select_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Оберіть DOCX файл",
            filetypes=[("DOCX files", "*.docx")],
        )
        if not file_path:
            return
        self.input_path_var.set(file_path)
        self.output_dir_var.set(os.path.dirname(file_path))

    def select_output_dir(self) -> None:
        directory = filedialog.askdirectory(title="Оберіть папку для збереження")
        if not directory:
            return
        self.output_dir_var.set(directory)

    def start_anonymization(self) -> None:
        input_path = self.input_path_var.get().strip()
        output_dir = self.output_dir_var.get().strip()

        if not input_path:
            messagebox.showerror("Помилка", "Файл не вибрано.")
            return

        if not input_path.lower().endswith(".docx"):
            messagebox.showerror("Помилка", "Файл має бути у форматі .docx.")
            return

        if not os.path.exists(input_path):
            messagebox.showerror("Помилка", "Файл не знайдено.")
            return

        if not output_dir:
            messagebox.showerror("Помилка", "Не вказано папку для збереження.")
            return

        if not os.path.isdir(output_dir):
            messagebox.showerror("Помилка", "Папка для збереження не існує.")
            return

        base_name = os.path.splitext(os.path.basename(input_path))[0]
        output_path = os.path.join(output_dir, f"{base_name}_anon.docx")

        self.status_var.set("Обробка...")
        self.root.update_idletasks()

        try:
            process_docx(
                input_path=input_path,
                output_path=output_path,
                generator_full=generator_full_random,
                generator_init=generator_init_random,
                replace_full=self.replace_full_var.get(),
                replace_initials=self.replace_initials_var.get(),
                process_tables=self.process_tables_var.get(),
                process_headers_footers=self.process_headers_var.get(),
            )
        except PermissionError:
            self.status_var.set("Помилка")
            messagebox.showerror(
                "Помилка",
                "Немає доступу до файлу. Закрийте Word або перевірте права на запис.",
            )
            return
        except Exception:
            self.status_var.set("Помилка")
            messagebox.showerror(
                "Помилка",
                "Сталася помилка під час читання або збереження файлу.",
            )
            return

        self.status_var.set("Готово")
        messagebox.showinfo("Готово", f"Файл збережено: {output_path}")


def main() -> None:
    root = tk.Tk()
    app = AnonApp(root)
    _ = app
    root.mainloop()


if __name__ == "__main__":
    main()
