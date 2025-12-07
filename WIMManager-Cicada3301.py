import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import scrolledtext
import subprocess
import threading
import os
import sys
import urllib.request
import zipfile
import tempfile
import webbrowser

APP_TITLE = "WIM Manager v1.0 Cicada3301"
HEADER_TEXT = "Cicada3301"
DEFAULT_INDEX = "1"   # индекс по умолчанию


def is_windows():
    return os.name == "nt"

def resource_path(relative_path: str) -> str:
    """
    Для доступа к файлам (icon, wimlib и т.п.) как в обычном .py, так и внутри .exe (PyInstaller).
    """
    try:
        # когда запущено из exe
        base_path = sys._MEIPASS  # type: ignore
    except Exception:
        # когда запускаем обычный .py
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def shutil_which(cmd):
    try:
        from shutil import which
        return which(cmd)
    except Exception:
        paths = os.environ.get("PATH", "").split(os.pathsep)
        for path in paths:
            full = os.path.join(path, cmd)
            if os.path.isfile(full):
                return full
        return None


class WimManagerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("900x520")
        self.root.minsize(880, 520)
        try:
            icon_file = resource_path("C:\\Users\\denis\\Desktop\\777\\logo.ico")
            # Для дебага можно посмотреть путь:
            # print("ICON PATH:", icon_file, "exists:", os.path.exists(icon_file))
            self.root.iconbitmap(icon_file)
        except Exception as e:
            print(f"Не удалось установить иконку: {e}")
        self.theme_var = tk.StringVar(value="dark")      # dark / light
        self.backend_var = tk.StringVar(value="dism")    # auto / dism / wimlib

        self.wim_path_var = tk.StringVar()
        self.mount_path_var = tk.StringVar()
        self.index_var = tk.StringVar(value=DEFAULT_INDEX)
        self.status_var = tk.StringVar(value="Готов к работе")

        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass

        self.apply_theme()       # настроим стили под тему
        self.build_ui()
        self.center_window(900, 520)

        self.log("Приложение запущено.")
        self.detect_tools()

    # -------------------------------------------------------- UI / ТЕМА

    def center_window(self, width, height):
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - width) // 2
        y = (sh - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def apply_theme(self, *_):
        theme = self.theme_var.get()

        if theme == "light":
            bg = "#f4f4f7"
            fg = "#111827"
            accent = "#0078d4"
            entry_bg = "#ffffff"
            button_bg = "#e5e7eb"
            log_bg = "#ffffff"
            log_fg = "#111827"
        else:  # dark
            bg = "#1e1e2f"
            fg = "#ffffff"
            accent = "#00ffc6"
            entry_bg = "#25253a"
            button_bg = "#2c2c44"
            log_bg = "#111827"
            log_fg = "#e5e7eb"

        self.root.configure(bg=bg)

        self.style.configure("TFrame", background=bg)
        self.style.configure("TLabel", background=bg, foreground=fg, font=("Segoe UI", 10))
        self.style.configure("Header.TLabel", background=bg, foreground=accent,
                             font=("Segoe UI Semibold", 20))

        # Entry
        self.style.configure("Path.TEntry",
                             fieldbackground=entry_bg,
                             foreground=fg,
                             bordercolor=button_bg)

        # Buttons
        self.style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=6)
        self.style.map("Accent.TButton",
                       background=[("active", accent), ("!disabled", accent)],
                       foreground=[("!disabled", "#000000")])

        self.style.configure("Secondary.TButton", font=("Segoe UI", 10), padding=6)
        self.style.map("Secondary.TButton",
                       background=[("active", button_bg), ("!disabled", button_bg)],
                       foreground=[("!disabled", fg)])

        # Progress
        self.style.configure("Horizontal.TProgressbar",
                             troughcolor=entry_bg,
                             bordercolor=entry_bg)

        # Лог, если уже создан
        if hasattr(self, "log_text"):
            self.log_text.configure(bg=log_bg, fg=log_fg, insertbackground=log_fg)

    def build_ui(self):
        header = ttk.Label(self.root, text=HEADER_TEXT, style="Header.TLabel", anchor="center")
        header.pack(fill="x", pady=(10, 5))

        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill="x", padx=20, pady=(0, 5))

        # Тема
        ttk.Label(top_frame, text="Тема:").grid(row=0, column=0, sticky="w")
        theme_cb = ttk.Combobox(top_frame, width=10,
                                values=["dark", "light"],
                                textvariable=self.theme_var,
                                state="readonly")
        theme_cb.grid(row=0, column=1, padx=(4, 15))
        theme_cb.bind("<<ComboboxSelected>>", self.apply_theme)

        # Бэкенд
        ttk.Label(top_frame, text="Бэкенд:").grid(row=0, column=2, sticky="w")
        backend_cb = ttk.Combobox(
            top_frame,
            width=12,
            values=["dism", "wimlib"],
            textvariable=self.backend_var,
            state="readonly"
        )
        backend_cb.grid(row=0, column=3, padx=(4, 15))

        # Кнопки установки DISM и wimlib
        install_frame = ttk.Frame(top_frame)
        install_frame.grid(row=0, column=4, padx=(10, 0))

        ttk.Button(
            install_frame,
            text="Установить DISM",
            style="Secondary.TButton",
            command=self.install_dism
        ).grid(row=0, column=0, padx=4)

        ttk.Button(
            install_frame,
            text="Установить wimlib",
            style="Secondary.TButton",
            command=self.install_wimlib
        ).grid(row=0, column=1, padx=4)

        # Основной фрейм с полями и кнопками
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        main_frame.columnconfigure(1, weight=1)

        # WIM
        ttk.Label(main_frame, text="WIM-файл:").grid(row=0, column=0, sticky="w", pady=5, padx=(0, 8))
        wim_entry = ttk.Entry(main_frame, textvariable=self.wim_path_var, style="Path.TEntry")
        wim_entry.grid(row=0, column=1, sticky="ew", pady=5)
        ttk.Button(main_frame, text="Выбрать", style="Secondary.TButton",
                   command=self.choose_wim).grid(row=0, column=2, sticky="w", padx=(8, 0), pady=5)

        # Mount dir
        ttk.Label(main_frame, text="Папка монтирования:").grid(row=1, column=0, sticky="w", pady=5,
                                                               padx=(0, 8))
        mount_entry = ttk.Entry(main_frame, textvariable=self.mount_path_var, style="Path.TEntry")
        mount_entry.grid(row=1, column=1, sticky="ew", pady=5)
        ttk.Button(main_frame, text="Выбрать", style="Secondary.TButton",
                   command=self.choose_mount_dir).grid(row=1, column=2, sticky="w", padx=(8, 0), pady=5)

        # Index
        ttk.Label(main_frame, text="Индекс образа:").grid(row=2, column=0, sticky="w",
                                                          pady=5, padx=(0, 8))
        idx_entry = ttk.Entry(main_frame, textvariable=self.index_var, width=6)
        idx_entry.grid(row=2, column=1, sticky="w", pady=5)

        ttk.Button(main_frame, text="Показать индексы WIM",
                   style="Secondary.TButton",
                   command=self.show_wim_indexes).grid(row=2, column=2, sticky="w",
                                                       padx=(8, 0), pady=5)

        # Кнопки операций
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=3, column=0, columnspan=3, pady=(10, 5))

        ttk.Button(btn_frame, text="Монтировать", style="Accent.TButton",
                   command=self.mount_wim).grid(row=0, column=0, padx=8)

        ttk.Button(btn_frame, text="Размонтировать (Commit)", style="Secondary.TButton",
                   command=lambda: self.unmount_wim(discard=False)).grid(row=0, column=1, padx=8)

        ttk.Button(btn_frame, text="Размонтировать (Discard)", style="Secondary.TButton",
                   command=lambda: self.unmount_wim(discard=True)).grid(row=0, column=2, padx=8)

        ttk.Button(btn_frame, text="Смонтированные WIM", style="Secondary.TButton",
                   command=self.show_mounted_wim).grid(row=0, column=3, padx=8)

        # Прогресс
        self.progress = ttk.Progressbar(main_frame, mode="indeterminate",
                                        style="Horizontal.TProgressbar")
        self.progress.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(5, 2))

        # Статус
        status_label = ttk.Label(main_frame, textvariable=self.status_var, anchor="w")
        status_label.grid(row=5, column=0, columnspan=3, sticky="w", pady=(0, 5))

        # Лог
        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=6, column=0, columnspan=3, sticky="nsew", pady=(5, 0))
        main_frame.rowconfigure(6, weight=1)

        self.log_text = scrolledtext.ScrolledText(log_frame, wrap="word", height=10)
        self.log_text.pack(fill="both", expand=True)
        self.log_text.configure(state="disabled")
        self.apply_theme()  # обновим цвета лог-окна под текущую тему

    # -------------------------------------------------------- ВСПОМОГАТЕЛЬНОЕ

    def log(self, text: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def detect_tools(self):
        dism = shutil_which("dism.exe") if is_windows() else None
        wimlib = shutil_which("wimlib-imagex.exe") or shutil_which("wimlib-imagex")

        msg_parts = []
        if dism:
            msg_parts.append("DISM найден")
        else:
            msg_parts.append("DISM не найден (или не в PATH)")

        if wimlib:
            msg_parts.append("wimlib-imagex найден")
        else:
            msg_parts.append("wimlib-imagex не найден")

        self.log(" / ".join(msg_parts))
        self.log("Для обновления статуса можно перезапустить программу после установки.")

    def get_backend(self):
        mode = self.backend_var.get()
        dism = shutil_which("dism.exe") if is_windows() else None
        wimlib = shutil_which("wimlib-imagex.exe") or shutil_which("wimlib-imagex")

        if mode == "dism":
            if not dism:
                raise RuntimeError("DISM не найден. Выберите другой бэкенд или добавьте DISM в PATH.")
            return "dism"
        if mode == "wimlib":
            if not wimlib:
                raise RuntimeError("wimlib-imagex не найден. Установите его или выберите другой бэкенд.")
            return "wimlib"

        # auto
        if wimlib:
            return "wimlib"
        if dism:
            return "dism"
        raise RuntimeError("Не найден ни DISM, ни wimlib-imagex.")

    # -------------------------------------------------------- УСТАНОВКА

    def install_dism(self):
        self.log("Открывается страница загрузки ADK (DISM)...")
        messagebox.showinfo(
            "DISM",
            "DISM входит в состав Windows ADK.\nСейчас откроется страница загрузки ADK."
        )
        webbrowser.open("https://learn.microsoft.com/en-us/windows-hardware/get-started/adk-install")

    def install_wimlib(self):
        self.log("Попытка установить wimlib-imagex...")

        try:
            url = "https://wimlib.net/downloads/wimlib-1.14.4-windows-x86_64-bin.zip"
            tmp_zip = os.path.join(tempfile.gettempdir(), "wimlib-imagex.zip")

            self.log(f"Скачивание {url} -> {tmp_zip}")
            urllib.request.urlretrieve(url, tmp_zip)

            target_dir = os.path.join(os.getcwd(), "wimlib")
            os.makedirs(target_dir, exist_ok=True)

            self.log(f"Распаковка в {target_dir}")
            with zipfile.ZipFile(tmp_zip, "r") as z:
                z.extractall(target_dir)

            # ищем бинарник
            bin_path = None
            for root_dir, dirs, files in os.walk(target_dir):
                if "wimlib-imagex.exe" in files:
                    bin_path = root_dir
                    break

            if not bin_path:
                messagebox.showerror("Ошибка", "Не удалось найти wimlib-imagex после распаковки.")
                self.log("wimlib-imagex.exe не найден после распаковки.")
                return

            # добавляем в PATH для текущего процесса
            os.environ["PATH"] = bin_path + os.pathsep + os.environ.get("PATH", "")
            self.log(f"wimlib-imagex найден в: {bin_path}")
            self.log("Путь добавлен во временный PATH текущего процесса.")

            messagebox.showinfo(
                "wimlib-imagex",
                f"wimlib-imagex установлен в:\n{bin_path}\n\n"
                f"В этом запуске программы уже можно выбрать бэкенд 'wimlib'."
            )

            # обновим детект
            self.detect_tools()

        except Exception as e:
            messagebox.showerror("Ошибка установки wimlib", str(e))
            self.log(f"Ошибка при установке wimlib: {e}")

    # -------------------------------------------------------- ОБРАБОТЧИКИ

    def choose_wim(self):
        filename = filedialog.askopenfilename(
            title="Выбрать WIM-файл",
            filetypes=[("WIM файлы", "*.wim"), ("Все файлы", "*.*")]
        )
        if filename:
            self.wim_path_var.set(filename)

    def choose_mount_dir(self):
        dirname = filedialog.askdirectory(title="Выбрать папку для монтирования")
        if dirname:
            self.mount_path_var.set(dirname)

    def mount_wim(self):
        wim = self.wim_path_var.get().strip()
        mount_dir = self.mount_path_var.get().strip()
        index = self.index_var.get().strip() or DEFAULT_INDEX

        if not wim or not os.path.isfile(wim):
            messagebox.showwarning("Внимание", "Выберите корректный WIM-файл.")
            return
        if not mount_dir or not os.path.isdir(mount_dir):
            messagebox.showwarning("Внимание", "Выберите существующую папку монтирования.")
            return

        try:
            backend = self.get_backend()
        except RuntimeError as e:
            messagebox.showerror("Ошибка", str(e))
            return

        if backend == "dism":
            if not is_windows():
                messagebox.showerror("Ошибка", "DISM поддерживается только в Windows.")
                return
            cmd = [
                "dism", "/English", "/Mount-Wim",
                f"/WimFile:{wim}",
                f"/index:{index}",
                f"/MountDir:{mount_dir}",
            ]
        else:  # wimlib
            cmd = ["wimlib-imagex", "mount", wim, index, mount_dir]

        self.run_command_async(cmd, f"Монтирование ({backend})")

    def unmount_wim(self, discard=False):
        mount_dir = self.mount_path_var.get().strip()
        if not mount_dir or not os.path.isdir(mount_dir):
            messagebox.showwarning("Внимание", "Выберите корректную папку монтирования.")
            return

        try:
            backend = self.get_backend()
        except RuntimeError as e:
            messagebox.showerror("Ошибка", str(e))
            return

        if backend == "dism":
            if not is_windows():
                messagebox.showerror("Ошибка", "DISM поддерживается только в Windows.")
                return
            cmd = [
                "dism", "/English", "/Unmount-Wim",
                f"/MountDir:{mount_dir}",
                "/Discard" if discard else "/Commit",
            ]
        else:
            cmd = ["wimlib-imagex", "unmount", mount_dir]
            if not discard:
                cmd.append("--commit")

        suffix = "Discard" if discard else "Commit"
        self.run_command_async(cmd, f"Размонтирование ({backend}, {suffix})")

    def show_mounted_wim(self):
        # Определяем через DISM – он есть почти в любой Windows
        if not is_windows() or not shutil_which("dism.exe"):
            messagebox.showerror("Ошибка", "DISM недоступен, невозможно прочитать список смонтированных WIM.")
            return

        cmd = ["dism", "/English", "/Get-MountedWimInfo"]
        self.run_command_async(cmd, "Список смонтированных WIM", show_message=False)

    def show_wim_indexes(self):
        wim = self.wim_path_var.get().strip()
        if not wim or not os.path.isfile(wim):
            messagebox.showwarning("Внимание", "Сначала выберите WIM-файл.")
            return

        # Для списка индексов пробуем выбранный бэкенд, а если не получится – DISM.
        try:
            backend = self.get_backend()
        except RuntimeError:
            backend = "dism"

        if backend == "wimlib" and shutil_which("wimlib-imagex"):
            cmd = ["wimlib-imagex", "info", wim]
            title = "Индексы WIM (wimlib)"
        else:
            if not is_windows() or not shutil_which("dism.exe"):
                messagebox.showerror("Ошибка", "Нет ни wimlib-imagex, ни DISM для показа индексов.")
                return
            cmd = ["dism", "/English", "/Get-WimInfo", f"/WimFile:{wim}"]
            title = "Индексы WIM (DISM)"

        self.run_command_async(cmd, title, show_message=False)

    # -------------------------------------------------------- ВЫПОЛНЕНИЕ КОМАНД

    def start_progress(self, text):
        self.status_var.set(text)
        self.progress.start(10)

    def stop_progress(self, text="Готово"):
        self.progress.stop()
        self.status_var.set(text)

    def run_command_async(self, cmd, action_name, show_message=True):
        self.start_progress(f"{action_name}...")

        def worker():
            try:
                creationflags = 0
                if is_windows():
                    creationflags = subprocess.CREATE_NO_WINDOW  # type: ignore

                self.log(f">>> {action_name}: {' '.join(cmd)}")
                result = subprocess.run(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    creationflags=creationflags
                )
                code = result.returncode
                output = result.stdout
                self.log(output.strip() or "<пустой вывод>")
                success = (code == 0)
            except Exception as e:
                code = None
                output = str(e)
                self.log(f"Исключение: {output}")
                success = False

            def on_complete():
                if code == 740:
                    self.stop_progress("Нужны права администратора")
                    messagebox.showerror(
                        "Ошибка 740",
                        "Операция требует прав администратора.\n"
                        "Запустите программу (или консоль) от имени администратора."
                    )
                    return

                if success:
                    self.stop_progress(f"{action_name} завершено")
                    if show_message:
                        messagebox.showinfo("Готово", f"{action_name} успешно завершено.")
                else:
                    self.stop_progress(f"{action_name} завершилось с ошибкой")
                    if show_message:
                        messagebox.showerror(
                            "Ошибка",
                            f"{action_name} завершилось с ошибкой.\n"
                            f"Код: {code}\n\n{output}"
                        )

            self.root.after(0, on_complete)

        threading.Thread(target=worker, daemon=True).start()


def main():
    root = tk.Tk()
    app = WimManagerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
