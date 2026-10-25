from __future__ import annotations

import math
import os
import shutil
import subprocess
import tempfile
import threading
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter as tk

import pymupdf
from PIL import Image, ImageOps
from pdf2docx import Converter
from tkinterdnd2 import DND_FILES, TkinterDnD


APP_NAME = "ClaraConvert"
IMAGE_FORMATS = {"PNG": "PNG", "JPG": "JPEG", "WEBP": "WEBP"}
SAME_FORMAT_EXTENSIONS = {
    "PDF": {".pdf"},
    "DOCX": {".docx"},
    "PNG": {".png"},
    "JPG": {".jpg", ".jpeg"},
    "WEBP": {".webp"},
}
SUPPORTED_IMAGES = {
    ".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff", ".gif"
}
FORMATS_BY_EXTENSION = {
    ".pdf": ("DOCX", "PNG", "JPG"),
    ".docx": ("PDF", "PNG", "JPG"),
    **{extension: ("PDF", "PNG", "JPG", "WEBP") for extension in SUPPORTED_IMAGES},
}
COLORS = {
    "window": "#002e34",
    "surface": "#004443",
    "surface_alt": "#002e34",
    "text": "#E8FFF6",
    "muted": "#9FD8C4",
    "accent": "#00c16c",
    "accent_hover": "#19d982",
    "entry": "#002e34",
    "border": "#00755c",
    "progress": "#90ff17",
    "error": "#FF6B6B",
    "disabled": "#00755c",
}


def find_libreoffice() -> str | None:
    executable = shutil.which("soffice") or shutil.which("libreoffice")
    if executable:
        return executable
    if os.name == "nt":
        candidates = (
            Path(os.environ.get("PROGRAMFILES", "C:/Program Files"))
            / "LibreOffice" / "program" / "soffice.exe",
            Path(os.environ.get("PROGRAMFILES(X86)", "C:/Program Files (x86)"))
            / "LibreOffice" / "program" / "soffice.exe",
        )
        for candidate in candidates:
            if candidate.is_file():
                return str(candidate)
    return None


def docx_to_pdf(source: Path, destination: Path) -> None:
    executable = find_libreoffice()
    if not executable:
        raise RuntimeError(
            "Para converter Word (DOCX) em PDF, instale o LibreOffice neste "
            "computador. Depois, abra o ClaraConvert novamente."
        )

    with tempfile.TemporaryDirectory(prefix="claraconvert-") as profile_dir:
        profile = Path(profile_dir).as_uri()
        result = subprocess.run(
            [
                executable,
                "--headless",
                f"-env:UserInstallation={profile}",
                "--convert-to",
                "pdf",
                "--outdir",
                str(destination.parent),
                str(source),
            ],
            capture_output=True,
            text=True,
            timeout=180,
            check=False,
        )
    generated = destination.parent / f"{source.stem}.pdf"
    if result.returncode != 0 or not generated.is_file():
        detail = (result.stderr or result.stdout).strip()
        raise RuntimeError(
            "O LibreOffice não conseguiu converter o documento."
            + (f"\n\nDetalhes: {detail}" if detail else "")
        )
    if generated != destination:
        generated.replace(destination)


def render_pdf(source: Path, output_dir: Path, image_format: str) -> list[Path]:
    document = pymupdf.open(source)
    if document.page_count == 0:
        raise ValueError("O PDF não contém páginas.")
    outputs = []
    for index, page in enumerate(document, start=1):
        pixmap = page.get_pixmap(matrix=pymupdf.Matrix(2, 2), alpha=False)
        image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
        suffix = "jpg" if image_format == "JPEG" else image_format.lower()
        output = output_dir / f"{source.stem}_pagina_{index:03d}.{suffix}"
        save_image(image, output, image_format)
        outputs.append(output)
    document.close()
    return outputs


def save_image(image: Image.Image, destination: Path, image_format: str) -> None:
    image = ImageOps.exif_transpose(image)
    if image_format == "JPEG":
        if image.mode not in ("RGB", "L"):
            background = Image.new("RGB", image.size, "white")
            if "A" in image.getbands():
                background.paste(image, mask=image.getchannel("A"))
            else:
                background.paste(image.convert("RGB"))
            image = background
        image.save(destination, format="JPEG", quality=94, optimize=True)
    else:
        image.save(destination, format=image_format)


def is_same_format(source: Path, target_format: str) -> bool:
    return source.suffix.lower() in SAME_FORMAT_EXTENSIONS.get(target_format, set())


def convert_file(source: Path, output_dir: Path, target_format: str) -> list[Path]:
    if is_same_format(source, target_format):
        raise ValueError(f"O arquivo já está no formato {target_format}.")
    output_dir.mkdir(parents=True, exist_ok=True)
    extension = source.suffix.lower()
    outputs: list[Path]

    if extension == ".pdf" and target_format == "DOCX":
        destination = output_dir / f"{source.stem}.docx"
        converter = Converter(str(source))
        try:
            converter.convert(str(destination))
        finally:
            converter.close()
        outputs = [destination]
    elif extension == ".docx" and target_format == "PDF":
        destination = output_dir / f"{source.stem}.pdf"
        docx_to_pdf(source, destination)
        outputs = [destination]
    elif extension == ".docx" and target_format in IMAGE_FORMATS:
        with tempfile.TemporaryDirectory(prefix="claraconvert-") as temp_dir:
            pdf_path = Path(temp_dir) / f"{source.stem}.pdf"
            docx_to_pdf(source, pdf_path)
            outputs = render_pdf(pdf_path, output_dir, IMAGE_FORMATS[target_format])
    elif extension == ".pdf" and target_format in IMAGE_FORMATS:
        outputs = render_pdf(source, output_dir, IMAGE_FORMATS[target_format])
    elif extension in SUPPORTED_IMAGES and target_format == "PDF":
        destination = output_dir / f"{source.stem}.pdf"
        with Image.open(source) as image:
            frames = []
            for frame_index in range(getattr(image, "n_frames", 1)):
                image.seek(frame_index)
                frame = ImageOps.exif_transpose(image.copy()).convert("RGB")
                frames.append(frame)
            if not frames:
                raise ValueError("Não foi possível ler a imagem.")
            frames[0].save(
                destination,
                "PDF",
                save_all=True,
                append_images=frames[1:],
                resolution=150.0,
            )
        outputs = [destination]
    elif extension in SUPPORTED_IMAGES and target_format in IMAGE_FORMATS:
        image_format = IMAGE_FORMATS[target_format]
        suffix = "jpg" if image_format == "JPEG" else image_format.lower()
        destination = output_dir / f"{source.stem}.{suffix}"
        with Image.open(source) as image:
            save_image(image.copy(), destination, image_format)
        outputs = [destination]
    else:
        raise ValueError("Esta combinação de arquivo e formato não é suportada.")

    missing = [str(path) for path in outputs if not path.is_file()]
    if missing:
        raise RuntimeError("A conversão terminou sem gerar todos os arquivos esperados.")
    return outputs


class ClaraConvertApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"{APP_NAME} | Conversor de arquivos")
        self.root.geometry("1020x820")
        self.root.minsize(900, 760)
        self.root.configure(bg=COLORS["window"])
        self.source_paths: list[Path] = []
        self.source_item_paths: dict[str, Path] = {}
        self.is_converting = False
        self.progress_value = 0.0
        self.progress_total = 0
        self.destination_dir = Path.home() / "Downloads"
        self.history_paths: dict[str, Path] = {}

        self._configure_style()
        self._build_interface()
        self._update_formats()

    def _configure_style(self) -> None:
        self.root.option_add("*TCombobox*Listbox.background", COLORS["entry"])
        self.root.option_add("*TCombobox*Listbox.foreground", COLORS["text"])
        self.root.option_add("*TCombobox*Listbox.selectBackground", COLORS["border"])
        self.root.option_add("*TCombobox*Listbox.selectForeground", "#FFFFFF")
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(
            "TCombobox",
            fieldbackground=COLORS["entry"],
            background=COLORS["surface"],
            foreground=COLORS["text"],
            arrowcolor=COLORS["text"],
            bordercolor=COLORS["border"],
            lightcolor=COLORS["border"],
            darkcolor=COLORS["border"],
            focuscolor=COLORS["progress"],
            padding=8,
        )
        style.map(
            "TCombobox",
            fieldbackground=[
                ("readonly", COLORS["entry"]),
                ("focus", COLORS["entry"]),
                ("disabled", COLORS["surface"]),
            ],
            foreground=[
                ("readonly", COLORS["text"]),
                ("disabled", COLORS["muted"]),
            ],
            selectbackground=[("readonly", COLORS["border"])],
            selectforeground=[("readonly", COLORS["text"])],
            arrowcolor=[("active", COLORS["progress"])],
        )
        style.configure(
            "History.Treeview",
            background=COLORS["window"],
            fieldbackground=COLORS["window"],
            foreground=COLORS["text"],
            rowheight=32,
            bordercolor=COLORS["border"],
            lightcolor=COLORS["border"],
            darkcolor=COLORS["border"],
            focuscolor=COLORS["progress"],
            font=("Segoe UI", 9),
        )
        style.configure(
            "History.Treeview.Heading",
            background=COLORS["border"],
            foreground=COLORS["text"],
            font=("Segoe UI Semibold", 9),
            padding=(8, 7),
            relief="flat",
        )
        style.map(
            "History.Treeview.Heading",
            background=[("active", COLORS["border"])],
            foreground=[("active", COLORS["text"])],
        )
        style.map(
            "History.Treeview",
            background=[("selected", COLORS["border"])],
            foreground=[("selected", "#FFFFFF")],
        )
        style.configure(
            "Vertical.TScrollbar",
            background=COLORS["border"],
            troughcolor=COLORS["surface"],
            arrowcolor=COLORS["text"],
            bordercolor=COLORS["surface"],
            lightcolor=COLORS["surface"],
            darkcolor=COLORS["surface"],
            width=12,
        )
        style.map(
            "Vertical.TScrollbar",
            background=[("active", COLORS["accent"]), ("pressed", COLORS["accent"])],
            arrowcolor=[("active", COLORS["text"])],
        )

    def _label(self, parent: tk.Widget, text: str, **kwargs: object) -> tk.Label:
        kwargs.setdefault("bg", COLORS["surface"])
        kwargs.setdefault("fg", COLORS["text"])
        kwargs.setdefault("font", ("Segoe UI", 10))
        return tk.Label(parent, text=text, **kwargs)

    def _rounded_panel(
        self, parent: tk.Widget, color: str, radius: int = 18, inset: int = 8
    ) -> tuple[tk.Canvas, tk.Frame]:
        canvas = tk.Canvas(
            parent,
            bg=parent.cget("bg"),
            highlightthickness=0,
            borderwidth=0,
        )
        content = tk.Frame(canvas, bg=color)
        window = canvas.create_window(
            inset, inset, anchor="nw", window=content
        )

        def resize(event: tk.Event) -> None:
            width = max(event.width - inset * 2, 1)
            height = max(event.height - inset * 2, 1)
            canvas.coords(window, inset, inset)
            canvas.itemconfigure(window, width=width, height=height)
            canvas.delete("rounded-panel")
            corner = min(radius, event.width // 2, event.height // 2)
            right = event.width - 1
            bottom = event.height - 1
            rounded_corners = (
                (right - corner, 1 + corner, -math.pi / 2, 0),
                (right - corner, bottom - corner, 0, math.pi / 2),
                (1 + corner, bottom - corner, math.pi / 2, math.pi),
                (1 + corner, 1 + corner, math.pi, math.pi * 1.5),
            )
            points = []
            for center_x, center_y, start_angle, end_angle in rounded_corners:
                for step in range(13):
                    angle = start_angle + (end_angle - start_angle) * step / 12
                    points.extend(
                        (
                            center_x + corner * math.cos(angle),
                            center_y + corner * math.sin(angle),
                        )
                    )
            canvas.create_polygon(
                *points,
                smooth=False,
                fill=color,
                outline=COLORS["border"],
                width=1,
                tags="rounded-panel",
            )
            canvas.tag_lower("rounded-panel")

        def update_height(_event: tk.Event) -> None:
            canvas.configure(height=content.winfo_reqheight() + inset * 2)

        canvas.bind("<Configure>", resize)
        content.bind("<Configure>", update_height)
        return canvas, content

    def _build_interface(self) -> None:
        outer = tk.Frame(self.root, bg=COLORS["window"], padx=24, pady=24)
        outer.pack(fill="both", expand=True)

        header = tk.Frame(outer, bg=COLORS["window"])
        header.pack(fill="x", pady=(0, 20))
        tk.Label(
            header,
            text="ClaraConvert",
            bg=COLORS["window"],
            fg=COLORS["text"],
            font=("Segoe UI Semibold", 23),
        ).pack(anchor="w")
        tk.Label(
            header,
            text="Conversões simples, sem complicação.",
            bg=COLORS["window"],
            fg=COLORS["muted"],
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(3, 0))

        body = tk.Frame(outer, bg=COLORS["window"])
        body.pack(fill="both", expand=True)
        body.grid_columnconfigure(0, weight=3, uniform="panels")
        body.grid_columnconfigure(1, weight=2, uniform="panels")
        body.grid_rowconfigure(0, weight=1)

        left_canvas, card = self._rounded_panel(body, COLORS["surface"], radius=18)
        left_canvas.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        card.configure(padx=24, pady=24)

        right_canvas, history_panel = self._rounded_panel(
            body, COLORS["surface"], radius=18
        )
        right_canvas.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        history_panel.configure(padx=16, pady=24)

        self._label(card, "1. Escolha os arquivos").pack(anchor="w")
        input_frame = tk.Frame(card, bg=COLORS["surface"])
        input_frame.pack(fill="x", pady=(9, 16))
        self.input_tree = ttk.Treeview(
            input_frame,
            columns=("file", "type"),
            show="headings",
            style="History.Treeview",
            selectmode="browse",
            height=6,
        )
        self.input_tree.heading("file", text="Arquivo selecionado")
        self.input_tree.heading("type", text="Tipo")
        self.input_tree.column("file", width=310, anchor="w")
        self.input_tree.column("type", width=65, anchor="center", stretch=False)
        self.input_tree.bind("<Double-1>", self._open_selected_input)
        self.input_empty_label = tk.Label(
            input_frame,
            text="Arraste arquivos aqui",
            bg=COLORS["window"],
            fg=COLORS["muted"],
            font=("Segoe UI", 10),
        )
        self.input_empty_label.place(relx=0.5, rely=0.5, anchor="center")
        self.input_empty_label.drop_target_register(DND_FILES)
        self.input_empty_label.dnd_bind("<<Drop>>", self._drop_files)
        input_scrollbar = ttk.Scrollbar(
            input_frame, orient="vertical", command=self.input_tree.yview
        )
        self.input_tree.configure(yscrollcommand=input_scrollbar.set)
        self.input_tree.pack(side="left", fill="both", expand=True)
        input_scrollbar.pack(side="right", fill="y")

        input_actions = tk.Frame(card, bg=COLORS["surface"])
        input_actions.pack(fill="x", pady=(0, 18))
        self.choose_button = self._button(
            input_actions,
            "Selecionar arquivo(s)",
            self._choose_files,
            variant="secondary",
        )
        self.choose_button.pack(side="left", ipady=7, padx=(0, 8))
        self.remove_file_button = self._button(
            input_actions,
            "Remover selecionado",
            self._remove_selected_input,
            variant="tertiary",
        )
        self.remove_file_button.pack(side="left", ipady=7, padx=(0, 8))
        self.clear_files_button = self._button(
            input_actions,
            "Limpar tudo",
            self._clear_input_files,
            variant="tertiary",
        )
        self.clear_files_button.pack(side="left", ipady=7)
        for widget in (input_frame, self.input_tree):
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", self._drop_files)

        self._label(card, "2. Formato desejado").pack(anchor="w")
        format_row = tk.Frame(card, bg=COLORS["surface"])
        format_row.pack(fill="x", pady=(8, 0))
        self.format_var = tk.StringVar()
        self.format_box = ttk.Combobox(
            format_row,
            textvariable=self.format_var,
            state="readonly",
            font=("Segoe UI", 10),
            width=30,
        )
        self.format_box.pack(side="left", anchor="w")
        self.format_var.set("Escolha o formato")
        self.convert_button = self._button(
            format_row, "Converter arquivo", self._convert
        )
        self.convert_button.pack(side="right", ipady=9)
        self.format_note = self._label(
            card,
            "Selecione primeiro um arquivo para ver os formatos disponíveis.",
            fg=COLORS["muted"],
            font=("Segoe UI", 9),
            justify="left",
        )
        self.format_note.pack(anchor="w", pady=(8, 8))

        progress_area = tk.Frame(card, bg=COLORS["surface"])
        self.progress_area = progress_area
        progress_area.pack(fill="x", pady=(0, 8))
        self.progress_label = tk.Label(
            progress_area,
            text="0 de 0",
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=("Segoe UI Semibold", 9),
            width=10,
            anchor="w",
        )
        self.progress_label.grid(row=0, column=0, padx=(0, 10))
        self.progress = tk.Canvas(
            progress_area,
            height=10,
            width=240,
            bg=COLORS["surface"],
            highlightthickness=0,
            borderwidth=0,
        )
        self.progress.grid(row=0, column=1, sticky="ew")
        progress_area.grid_columnconfigure(1, weight=1)
        self.progress.bind("<Configure>", self._draw_progress)
        progress_area.pack_forget()

        self._label(history_panel, "Arquivos convertidos").pack(anchor="w")
        tk.Label(
            history_panel,
            text="Dê duplo clique ou selecione e abra.",
            bg=COLORS["surface"],
            fg=COLORS["muted"],
            font=("Segoe UI", 9),
            anchor="w",
        ).pack(anchor="w", pady=(4, 12))
        history_frame = tk.Frame(history_panel, bg=COLORS["surface"])
        history_frame.pack(fill="both", expand=True)
        self.history_tree = ttk.Treeview(
            history_frame,
            columns=("file", "type"),
            show="headings",
            style="History.Treeview",
        )
        self.history_tree.heading("file", text="Arquivo")
        self.history_tree.heading("type", text="Tipo")
        self.history_tree.column("file", width=210, anchor="w")
        self.history_tree.column("type", width=65, anchor="center", stretch=False)
        self.history_empty_label = tk.Label(
            history_frame,
            text="Arquivos convertidos aparecerão aqui",
            bg=COLORS["window"],
            fg=COLORS["muted"],
            font=("Segoe UI", 9),
        )
        self.history_empty_label.place(relx=0.5, rely=0.5, anchor="center")
        history_scrollbar = ttk.Scrollbar(
            history_frame, orient="vertical", command=self.history_tree.yview
        )
        self.history_tree.configure(yscrollcommand=history_scrollbar.set)
        self.history_tree.pack(side="left", fill="both", expand=True)
        history_scrollbar.pack(side="right", fill="y")
        self.history_tree.bind("<Double-1>", self._open_selected_output)
        self.open_file_button = self._button(
            history_panel,
            "Abrir arquivo",
            self._open_selected_output,
            variant="secondary",
        )
        self.open_file_button.pack(anchor="e", pady=(12, 0), ipady=5)

        self.status_var = tk.StringVar(value="Pronto para começar")
        self.status_label = tk.Label(
            outer,
            textvariable=self.status_var,
            bg=COLORS["window"],
            fg=COLORS["muted"],
            font=("Segoe UI", 9),
            anchor="w",
        )
        self.status_label.pack(fill="x", pady=(20, 0))
        tk.Label(
            outer,
            text="Dica: a conversão de Word para PDF/imagens requer o LibreOffice instalado.",
            bg=COLORS["window"],
            fg=COLORS["muted"],
            font=("Segoe UI", 8),
            anchor="w",
        ).pack(fill="x", pady=(8, 0))

    def _button(
        self,
        parent: tk.Widget,
        text: str,
        command: object,
        variant: str = "primary",
    ) -> tk.Button:
        if variant == "primary":
            background = COLORS["accent"]
            foreground = COLORS["window"]
            hover = COLORS["accent_hover"]
            borderwidth = 0
            relief = "flat"
        elif variant == "secondary":
            background = COLORS["border"]
            foreground = COLORS["text"]
            hover = COLORS["accent"]
            borderwidth = 0
            relief = "flat"
        else:
            background = COLORS["surface"]
            foreground = COLORS["text"]
            hover = COLORS["border"]
            borderwidth = 1
            relief = "solid"
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=background,
            fg=foreground,
            activebackground=hover,
            activeforeground=(
                COLORS["window"] if variant in {"primary", "secondary"} else foreground
            ),
            disabledforeground="#FFFFFF",
            relief=relief,
            borderwidth=borderwidth,
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            highlightcolor=COLORS["progress"],
            cursor="hand2",
            font=("Segoe UI Semibold", 9),
            padx=14,
        )

    @staticmethod
    def _draw_rounded_rectangle(
        canvas: tk.Canvas,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        color: str,
        outline: str = "",
        width: int = 1,
        radius: float = 8,
    ) -> None:
        radius = min(radius, (x2 - x1) / 2, (y2 - y1) / 2)
        points: list[float] = []
        for center_x, center_y, start_angle, end_angle in (
            (x2 - radius, y1 + radius, -math.pi / 2, 0),
            (x2 - radius, y2 - radius, 0, math.pi / 2),
            (x1 + radius, y2 - radius, math.pi / 2, math.pi),
            (x1 + radius, y1 + radius, math.pi, math.pi * 1.5),
        ):
            for step in range(7):
                angle = start_angle + (end_angle - start_angle) * step / 6
                points.extend(
                    (
                        center_x + radius * math.cos(angle),
                        center_y + radius * math.sin(angle),
                    )
                )
        canvas.create_polygon(
            *points,
            fill=color,
            outline=outline,
            width=width,
            smooth=True,
            splinesteps=32,
        )

    def _draw_progress(self, _event: tk.Event | None = None) -> None:
        canvas = self.progress
        canvas.delete("all")
        width = canvas.winfo_width()
        height = canvas.winfo_height()
        if width < 40 or height < 6:
            return

        self._draw_rounded_rectangle(
            canvas,
            0,
            0,
            width,
            height,
            COLORS["surface"],
            radius=5,
        )
        fill_width = width * min(max(self.progress_value, 0.0), 1.0)
        if fill_width >= 3:
            self._draw_rounded_rectangle(
                canvas,
                0,
                0,
                fill_width,
                height,
                COLORS["progress"],
                radius=5,
            )

    def _update_progress(self, processed: int, total: int) -> None:
        self.progress_value = processed / total if total else 0.0
        self.progress_label.configure(text=f"{processed} de {total}")
        self._draw_progress()

    def _set_status(self, message: str, kind: str = "muted") -> None:
        color = {
            "success": COLORS["progress"],
            "error": COLORS["error"],
            "muted": COLORS["muted"],
            "normal": COLORS["text"],
        }[kind]
        self.status_var.set(message)
        self.status_label.configure(fg=color)

    def _choose_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Selecione um ou mais arquivos para converter",
            filetypes=[
                ("Arquivos aceitos", "*.pdf *.docx *.png *.jpg *.jpeg *.webp *.bmp *.tif *.tiff *.gif"),
                ("PDF", "*.pdf"),
                ("Word (DOCX)", "*.docx"),
                ("Imagens", "*.png *.jpg *.jpeg *.webp *.bmp *.tif *.tiff *.gif"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        if paths:
            self._set_files([Path(path) for path in paths])

    def _drop_files(self, event: tk.Event) -> str:
        if self.is_converting:
            return "break"
        try:
            dropped_paths = [Path(path) for path in self.root.tk.splitlist(event.data)]
        except tk.TclError as error:
            messagebox.showerror(
                "Não foi possível ler os arquivos",
                f"Os arquivos arrastados não puderam ser identificados.\n\n{error}",
            )
            return "break"

        accepted = [
            path for path in dropped_paths
            if path.is_file() and path.suffix.lower() in FORMATS_BY_EXTENSION
        ]
        rejected = [path for path in dropped_paths if path not in accepted]
        combined = [*self.source_paths, *accepted]
        self._set_files(combined)
        if rejected:
            names = "\n".join(path.name for path in rejected[:5])
            if len(rejected) > 5:
                names += f"\n... e mais {len(rejected) - 5} arquivo(s)"
            messagebox.showwarning(
                "Alguns itens foram ignorados",
                "São aceitos arquivos PDF, DOCX e imagens.\n\n" + names,
            )
        return "copy"

    def _set_files(self, paths: list[Path]) -> None:
        supported = [
            path for path in paths
            if path.is_file() and path.suffix.lower() in FORMATS_BY_EXTENSION
        ]
        unique: list[Path] = []
        seen: set[str] = set()
        for path in supported:
            identity = str(path.resolve()).casefold()
            if identity not in seen:
                unique.append(path)
                seen.add(identity)
        self.source_paths = unique
        for item in self.input_tree.get_children():
            self.input_tree.delete(item)
        self.source_item_paths.clear()
        for path in unique:
            item = self.input_tree.insert(
                "",
                "end",
                values=(path.name, path.suffix[1:].upper()),
                tags=("even" if len(self.source_item_paths) % 2 == 0 else "odd",),
            )
            self.source_item_paths[item] = path
        self.input_tree.tag_configure("even", background=COLORS["window"])
        self.input_tree.tag_configure("odd", background=COLORS["surface"])
        if unique:
            self.input_empty_label.place_forget()
        else:
            self.input_empty_label.place(relx=0.5, rely=0.5, anchor="center")
        self._update_formats()
        if unique:
            self._set_status(f"{len(unique)} arquivo(s) selecionado(s)")

    def _remove_selected_input(self) -> None:
        selection = self.input_tree.selection()
        if not selection:
            messagebox.showinfo(APP_NAME, "Selecione um arquivo para remover.")
            return
        removed = self.source_item_paths.pop(selection[0])
        self._set_files([path for path in self.source_paths if path != removed])
        self._set_status(f"Arquivo removido: {removed.name}")

    def _clear_input_files(self) -> None:
        if not self.source_paths:
            return
        self._set_files([])
        self._set_status("Seleção de arquivos limpa")

    def _update_formats(self) -> None:
        if not self.source_paths:
            options = []
        else:
            common_formats = set(FORMATS_BY_EXTENSION[self.source_paths[0].suffix.lower()])
            for source in self.source_paths[1:]:
                common_formats.intersection_update(
                    FORMATS_BY_EXTENSION[source.suffix.lower()]
                )
            options = [
                target
                for target in FORMATS_BY_EXTENSION[self.source_paths[0].suffix.lower()]
                if target in common_formats
            ]
        self.format_box["values"] = options
        if options:
            self.format_var.set(options[0])
            if any(source.suffix.lower() == ".pdf" for source in self.source_paths):
                note = (
                    "Formatos disponíveis para todos os arquivos. Cada página de PDF "
                    "vira uma imagem; PDFs digitalizados não recebem OCR."
                )
            elif any(source.suffix.lower() == ".docx" for source in self.source_paths):
                note = (
                    "Formatos disponíveis para todos os arquivos. Word para PDF ou "
                    "imagens requer LibreOffice instalado."
                )
            else:
                note = "Formatos disponíveis para todos os arquivos selecionados."
            self.format_note.configure(text=note)
        else:
            self.format_var.set("Escolha o formato")
            if self.source_paths:
                self.format_note.configure(
                    text="Os arquivos selecionados não têm um formato de saída em comum."
                )
            else:
                self.format_note.configure(
                    text="Selecione arquivos ou arraste-os para a área acima."
                )

    def _convert(self) -> None:
        if self.is_converting:
            return
        if not self.source_paths:
            messagebox.showinfo(APP_NAME, "Selecione um arquivo antes de converter.")
            return
        target = self.format_var.get()
        if not target or target == "Escolha o formato":
            messagebox.showinfo(APP_NAME, "Escolha o formato de saída.")
            return
        same_format_sources = [
            source for source in self.source_paths if is_same_format(source, target)
        ]
        if same_format_sources and len(self.source_paths) == 1:
            filenames = "\n".join(source.name for source in same_format_sources[:5])
            if len(same_format_sources) > 5:
                filenames += (
                    f"\n... e mais {len(same_format_sources) - 5} arquivo(s)"
                )
            messagebox.showerror(
                "Formato já selecionado",
                f"Este(s) arquivo(s) já está(ão) no formato {target}:\n\n"
                f"{filenames}\n\nNenhum arquivo foi convertido.",
            )
            return
        selected_dir = filedialog.askdirectory(
            title="Escolha onde salvar os arquivos convertidos",
            initialdir=str(self.destination_dir),
            mustexist=True,
            parent=self.root,
        )
        if not selected_dir:
            return
        output_dir = Path(selected_dir)
        self.destination_dir = output_dir

        self.is_converting = True
        self.progress_value = 0.0
        self.progress_total = len(self.source_paths)
        self.progress_area.pack(fill="x", pady=(0, 8))
        self._update_progress(0, self.progress_total)
        self.convert_button.configure(
            state="disabled",
            text="Convertendo...",
            bg=COLORS["disabled"],
            fg="#FFFFFF",
        )
        self.choose_button.configure(state="disabled")
        self.remove_file_button.configure(state="disabled")
        self.clear_files_button.configure(state="disabled")
        self._set_status(
            f"Convertendo {self.progress_total} arquivo(s)...", "normal"
        )

        sources = self.source_paths.copy()
        worker = threading.Thread(
            target=self._convert_worker,
            args=(sources, output_dir, target),
            daemon=True,
        )
        worker.start()

    def _convert_worker(
        self, sources: list[Path], output_dir: Path, target: str
    ) -> None:
        outputs: list[Path] = []
        failures: list[tuple[Path, str]] = []
        reserved_names: set[str] = set()
        converted_count = 0
        total = len(sources)
        for index, source in enumerate(sources, start=1):
            self.root.after(
                0,
                lambda current=index, name=source.name, total=total: (
                    self._set_status(
                        f"Convertendo {current} de {total}: {name}", "normal"
                    )
                ),
            )
            try:
                with tempfile.TemporaryDirectory(
                    prefix="claraconvert-", dir=output_dir
                ) as temp_dir:
                    temporary_outputs = convert_file(source, Path(temp_dir), target)
                    for temporary_output in temporary_outputs:
                        destination = self._unique_output_path(
                            output_dir, temporary_output.name, reserved_names
                        )
                        temporary_output.replace(destination)
                        outputs.append(destination)
                converted_count += 1
            except Exception as error:
                failures.append((source, str(error)))
            if failures:
                converted_label = "convertido" if converted_count == 1 else "convertidos"
                failure_label = "falha" if len(failures) == 1 else "falhas"
                progress_message = (
                    f"{index} de {total} arquivos processados "
                    f"({converted_count} {converted_label}, "
                    f"{len(failures)} {failure_label})"
                )
            else:
                file_label = "arquivo" if total == 1 else "arquivos"
                converted_label = "convertido" if total == 1 else "convertidos"
                progress_message = (
                    f"{converted_count} de {total} {file_label} {converted_label}"
                )
            self.root.after(
                0,
                lambda current=index, count=total, message=progress_message: (
                    self._update_progress(current, count),
                    self._set_status(message, "normal"),
                ),
            )
        self.root.after(
            0, lambda: self._conversion_finished(outputs, failures)
        )

    @staticmethod
    def _unique_output_path(
        output_dir: Path, filename: str, reserved_names: set[str]
    ) -> Path:
        original = Path(filename)
        candidate = output_dir / original.name
        number = 2
        while candidate.exists() or candidate.name.casefold() in reserved_names:
            candidate = output_dir / f"{original.stem}_{number}{original.suffix}"
            number += 1
        reserved_names.add(candidate.name.casefold())
        return candidate

    def _finish_conversion(self) -> None:
        self.is_converting = False
        self.progress_total = 0
        self.progress_value = 0.0
        self.progress_area.pack_forget()
        self._draw_progress()
        self.convert_button.configure(
            state="normal",
            text="Converter arquivo",
            bg=COLORS["accent"],
            fg=COLORS["window"],
        )
        self.choose_button.configure(state="normal")
        self.remove_file_button.configure(state="normal")
        self.clear_files_button.configure(state="normal")

    def _open_selected_input(self, _event: tk.Event | None = None) -> None:
        selection = self.input_tree.selection()
        if not selection:
            return

        source = self.source_item_paths[selection[0]]
        if not source.is_file():
            messagebox.showerror(
                APP_NAME, f"O arquivo não foi encontrado:\n\n{source}"
            )
            return
        try:
            os.startfile(str(source))
        except OSError as error:
            messagebox.showerror(
                APP_NAME,
                f"Não foi possível abrir o arquivo:\n\n{source}\n\n{error}",
            )

    def _open_selected_output(self, _event: tk.Event | None = None) -> None:
        selection = self.history_tree.selection()
        if not selection:
            if _event is None:
                messagebox.showinfo(
                    APP_NAME, "Selecione um arquivo convertido para abrir."
                )
            return

        output = self.history_paths[selection[0]]
        if not output.is_file():
            messagebox.showerror(
                APP_NAME, f"O arquivo não foi encontrado:\n\n{output}"
            )
            return
        try:
            os.startfile(str(output))
        except OSError as error:
            messagebox.showerror(
                APP_NAME,
                f"Não foi possível abrir o arquivo:\n\n{output}\n\n{error}",
            )

    def _conversion_finished(
        self, outputs: list[Path], failures: list[tuple[Path, str]]
    ) -> None:
        self._finish_conversion()
        for output in outputs:
            item_id = self.history_tree.insert(
                "",
                "end",
                values=(output.name, output.suffix[1:].upper()),
                tags=("even" if len(self.history_paths) % 2 == 0 else "odd",),
            )
            self.history_paths[item_id] = output
        self.history_tree.tag_configure("even", background=COLORS["window"])
        self.history_tree.tag_configure("odd", background=COLORS["surface"])
        if self.history_paths:
            self.history_empty_label.place_forget()
        self._set_status(
            f"Concluído: {len(outputs)} arquivo(s) gerado(s), "
            f"{len(failures)} falha(s)",
            "error" if failures else "success",
        )
        message = "\n".join(str(path) for path in outputs[:5])
        if len(outputs) > 5:
            message += f"\n... e mais {len(outputs) - 5} arquivo(s) gerado(s)"
        if failures:
            details = "\n".join(
                f"{source.name}: {error}" for source, error in failures[:5]
            )
            if len(failures) > 5:
                details += f"\n... e mais {len(failures) - 5} falha(s)"
            message += f"\n\nFalhas:\n{details}"
        if outputs:
            messagebox.showinfo(
                "Conversão concluída",
                f"Arquivos gerados: {len(outputs)}\n\n{message}",
            )
        else:
            messagebox.showerror("Falha na conversão", message or "Nenhum arquivo foi gerado.")


def main() -> None:
    root = TkinterDnD.Tk()
    ClaraConvertApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
