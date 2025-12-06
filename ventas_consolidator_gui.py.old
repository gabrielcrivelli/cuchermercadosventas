# ventas_consolidator_gui.py

import os
import threading
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

from core_consolidacion import (
    parsear_nombre_archivo,
    consolidar_datos,
    generar_reportes,
    MESES_ES,
)


class VentasConsolidatorGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Consolidador de Ventas")
        self.root.geometry("1100x700")
        self.root.minsize(900, 600)

        # --- Colores modo oscuro ---
        self.bg_color = "#1e1e1e"   # fondo principal
        self.fg_color = "#ffffff"   # texto
        self.accent_color = "#3a7bd5"

        self._setup_styles()

        # Datos de estado
        self.archivos = []  # lista de dicts: {ruta, mes, anio, sucursal}
        self.anio_var = tk.IntVar(value=datetime.now().year)
        self.meses_vars = {m: tk.BooleanVar(value=True) for m in MESES_ES}
        self.sucursales_extra_var = tk.StringVar()

        # Columnas disponibles / seleccionadas para 'Consolidado'
        self.columnas_disponibles_default = [
            "IdArticulo", "Marca", "Descripcion", "Departamento", "SubFamilia", "Familia",
            "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO",
            "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE",
            "OCTUBRE", "NOVIEMBRE", "DICIEMBRE",
            "TOTAL CORRIENTES", "TOTAL HIPER", "TOTAL CONSOLIDADO",
        ]

        self.columnas_seleccionadas_inicial = [
            "IdArticulo", "Marca", "Descripcion", "Departamento", "SubFamilia", "Familia",
            "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
            "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE",
            "TOTAL CORRIENTES", "TOTAL HIPER", "TOTAL CONSOLIDADO",
        ]

        # Flags de hojas a generar
        self.habilitar_ranking_var = tk.BooleanVar(value=True)
        self.habilitar_por_suc_var = tk.BooleanVar(value=True)
        self.habilitar_matriz_var = tk.BooleanVar(value=True)
        self.habilitar_evolucion_var = tk.BooleanVar(value=True)
        self.habilitar_especiales_var = tk.BooleanVar(value=True)
        self.habilitar_estacionalidad_var = tk.BooleanVar(value=True)

        # Filtros de categorías especiales
        self.especiales_deps_var = tk.StringVar()
        self.especiales_marcas_var = tk.StringVar()

        # Textos de ayuda
        self.help_hojas = {
            "ranking": (
                "Ranking de Ventas: lista cada artículo con la cantidad total vendida "
                "en el período, ordenado de mayor a menor."
            ),
            "sucursal": (
                "Ventas por Sucursal: muestra, para cada sucursal, cuánto se vendió por "
                "mes y el total del período."
            ),
            "matriz": (
                "Matriz por Departamento: tabla Departamento x Sucursal con las ventas "
                "totales y un total por departamento."
            ),
            "evolucion": (
                "Evolución Mensual: ventas por departamento mes a mes, útil para ver "
                "tendencias."
            ),
            "especiales": (
                "Categorías Especiales: igual que el consolidado, pero solo para los "
                "departamentos y/o marcas filtrados."
            ),
            "estacionalidad": (
                "Estacionalidad: para cada producto, calcula en qué mes del año se "
                "concentran más sus ventas y cuánto por encima del promedio mensual "
                "se encuentra ese mes (índice estacional)."
            ),
        }

        # Contenedor scroll general
        self._build_scroll_container()
        self._build_ui(self.content_frame)

    # ---------- Estilos ----------
    def _setup_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        self.root.configure(bg=self.bg_color)

        style.configure("TFrame", background=self.bg_color, foreground=self.fg_color)
        style.configure("TLabelframe", background=self.bg_color, foreground=self.fg_color)
        style.configure("TLabelframe.Label", background=self.bg_color, foreground=self.fg_color)
        style.configure("TLabel", background=self.bg_color, foreground=self.fg_color)
        style.configure("TCheckbutton", background=self.bg_color, foreground=self.fg_color)

        style.configure("TButton", background="#333333", foreground=self.fg_color)
        style.map(
            "TButton",
            background=[("active", "#444444")],
            foreground=[("disabled", "#777777")],
        )

        style.configure("TEntry", fieldbackground="#2b2b2b", foreground=self.fg_color)

        style.configure(
            "Treeview",
            background="#2b2b2b",
            fieldbackground="#2b2b2b",
            foreground=self.fg_color,
            bordercolor="#444444",
            borderwidth=0,
        )
        style.map(
            "Treeview",
            background=[("selected", "#444444")],
            foreground=[("selected", self.fg_color)],
        )
        style.configure("Treeview.Heading", background="#333333", foreground=self.fg_color)

    # ---------- Scroll general ----------
    def _build_scroll_container(self):
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(self.root, highlightthickness=0, bg=self.bg_color)
        self.v_scroll = ttk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.h_scroll = ttk.Scrollbar(self.root, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.v_scroll.grid(row=0, column=1, sticky="ns")
        self.h_scroll.grid(row=1, column=0, sticky="ew")

        self.content_frame = ttk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.content_frame, anchor="nw")

        self.content_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

    def _on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    # ---------- Construcción UI ----------
    def _build_ui(self, parent: ttk.Frame):
        # Top: año + meses
        frame_top = ttk.Frame(parent, padding=10)
        frame_top.pack(fill="x")

        ttk.Label(frame_top, text="Año:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame_top, textvariable=self.anio_var, width=6).grid(
            row=0, column=1, sticky="w", padx=(5, 20)
        )

        ttk.Label(frame_top, text="Meses a procesar:").grid(row=0, column=2, sticky="w")
        frame_meses = ttk.Frame(frame_top)
        frame_meses.grid(row=0, column=3, sticky="w")

        for i, mes in enumerate(MESES_ES):
            ttk.Checkbutton(
                frame_meses, text=mes.title(), variable=self.meses_vars[mes]
            ).grid(row=i // 4, column=i % 4, sticky="w")

        # Sucursales adicionales
        frame_suc_extra = ttk.Frame(parent, padding=(10, 0, 10, 10))
        frame_suc_extra.pack(fill="x")
        ttk.Label(frame_suc_extra, text="Sucursales adicionales (texto libre):").pack(side="left")
        ttk.Entry(frame_suc_extra, textvariable=self.sucursales_extra_var, width=50).pack(
            side="left", padx=5
        )

        # ---- Archivos Excel (arriba) ----
        frame_archivos = ttk.LabelFrame(parent, text="Archivos Excel", padding=10)
        frame_archivos.pack(fill="x", padx=10, pady=(0, 5))

        ttk.Button(frame_archivos, text="Agregar archivos...", command=self.agregar_archivos).pack(
            anchor="w"
        )

        tree_frame = ttk.Frame(frame_archivos)
        tree_frame.pack(fill="x", expand=False, pady=5)

        cols = ("ruta", "mes", "anio", "sucursal", "estado")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=6)

        for c in cols:
            self.tree.heading(c, text=c.upper())
            self.tree.column(c, width=150 if c != "ruta" else 400, anchor="w")

        tree_vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        tree_hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_vsb.set, xscrollcommand=tree_hsb.set)

        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_vsb.grid(row=0, column=1, sticky="ns")
        tree_hsb.grid(row=1, column=0, sticky="ew")

        # ---- Debajo: Configurar Excel Resultado (sin canvas interno) ----
        frame_opts = ttk.LabelFrame(parent, text="Configurar Excel Resultado", padding=10)
        frame_opts.pack(fill="both", expand=True, padx=10, pady=(0, 5))

        # Columnas disponibles / seleccionadas
        frame_cols = ttk.Frame(frame_opts)
        frame_cols.pack(side="left", fill="y", padx=(0, 20))

        ttk.Label(frame_cols, text="Columnas disponibles:").grid(row=0, column=0, sticky="w")
        ttk.Label(
            frame_cols, text="Columnas seleccionadas (orden de salida):"
        ).grid(row=0, column=2, sticky="w")

        self.lb_cols_disp = tk.Listbox(frame_cols, height=16, exportselection=False)
        self.lb_cols_sel = tk.Listbox(frame_cols, height=16, exportselection=False)

        for col in self.columnas_disponibles_default:
            self.lb_cols_disp.insert("end", col)
        for col in self.columnas_seleccionadas_inicial:
            self.lb_cols_sel.insert("end", col)

        self.lb_cols_disp.grid(row=1, column=0, sticky="nsw")
        self.lb_cols_sel.grid(row=1, column=2, sticky="nse")

        for lb in (self.lb_cols_disp, self.lb_cols_sel):
            lb.configure(
                bg=self.bg_color,
                fg=self.fg_color,
                selectbackground="#444444",
                selectforeground=self.fg_color,
                highlightbackground=self.bg_color,
                relief="solid",
                borderwidth=1,
            )

        frame_btn_cols_mid = ttk.Frame(frame_cols)
        frame_btn_cols_mid.grid(row=1, column=1, padx=5)

        ttk.Button(frame_btn_cols_mid, text="Agregar →", command=self._col_agregar).pack(
            fill="x", pady=2
        )
        ttk.Button(frame_btn_cols_mid, text="← Quitar", command=self._col_quitar).pack(
            fill="x", pady=2
        )

        frame_btn_cols_right = ttk.Frame(frame_cols)
        frame_btn_cols_right.grid(row=1, column=3, padx=5, sticky="n")
        ttk.Button(frame_btn_cols_right, text="Subir", command=self._mover_col_arriba).pack(
            fill="x", pady=2
        )
        ttk.Button(frame_btn_cols_right, text="Bajar", command=self._mover_col_abajo).pack(
            fill="x", pady=2
        )

        ttk.Label(frame_cols, text="Agregar columna manualmente:").grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(8, 0)
        )

        self.col_custom_var = tk.StringVar()
        frame_custom = ttk.Frame(frame_cols)
        frame_custom.grid(row=3, column=0, columnspan=2, sticky="w")
        ttk.Entry(frame_custom, textvariable=self.col_custom_var, width=30).pack(side="left")
        ttk.Button(frame_custom, text="Agregar", command=self._col_agregar_custom).pack(
            side="left", padx=4
        )

        # Hojas a generar + ayuda
        frame_checks = ttk.Frame(frame_opts)
        frame_checks.pack(side="left", fill="y", padx=(10, 20))

        ttk.Label(frame_checks, text="Hojas a generar:").pack(anchor="w")

        def add_hoja_row(texto, var, clave_help, titulo_help):
            fila = ttk.Frame(frame_checks)
            fila.pack(anchor="w")
            ttk.Checkbutton(fila, text=texto, variable=var).pack(side="left")
            ttk.Button(
                fila, text="?", width=2,
                command=lambda: self._explicar_hoja(clave_help, titulo_help),
            ).pack(side="left", padx=3)

        add_hoja_row("Ranking de Ventas", self.habilitar_ranking_var,
                     "ranking", "Ranking de Ventas")
        add_hoja_row("Ventas por Sucursal", self.habilitar_por_suc_var,
                     "sucursal", "Ventas por Sucursal")
        add_hoja_row("Matriz por Departamento", self.habilitar_matriz_var,
                     "matriz", "Matriz por Departamento")
        add_hoja_row("Evolución Mensual", self.habilitar_evolucion_var,
                     "evolucion", "Evolución Mensual")
        add_hoja_row("Categorías Especiales", self.habilitar_especiales_var,
                     "especiales", "Categorías Especiales")
        add_hoja_row("Estacionalidad", self.habilitar_estacionalidad_var,
                     "estacionalidad", "Estacionalidad")

        # Filtros de categorías especiales
        frame_espec = ttk.Frame(frame_opts)
        frame_espec.pack(side="left", fill="both", expand=True, padx=(20, 0))

        ttk.Label(frame_espec, text="Categorías Especiales - filtros opcionales").pack(
            anchor="w"
        )
        ttk.Label(frame_espec, text="Departamentos (separados por coma):").pack(anchor="w")
        ttk.Entry(frame_espec, textvariable=self.especiales_deps_var, width=45).pack(
            anchor="w", pady=(0, 5)
        )
        ttk.Label(frame_espec, text="Marcas (separadas por coma):").pack(anchor="w")
        ttk.Entry(frame_espec, textvariable=self.especiales_marcas_var, width=45).pack(
            anchor="w"
        )

        # ---- Zona inferior: botón + log (altura reducida) ----
        frame_bottom = ttk.Frame(parent, padding=10)
        frame_bottom.pack(fill="x", expand=False)

        self.btn_procesar = ttk.Button(frame_bottom, text="Procesar", command=self.procesar_async)
        self.btn_procesar.pack(side="top", anchor="w")

        self.log_text = scrolledtext.ScrolledText(
            frame_bottom, height=4, state="disabled", wrap="word"
        )
        self.log_text.pack(fill="x", expand=False, padx=10, pady=(5, 0))
        self.log_text.configure(
            background=self.bg_color,
            foreground=self.fg_color,
            insertbackground=self.fg_color,
        )

    # ---------- Lógica UI ----------
    def _explicar_hoja(self, clave: str, titulo: str):
        texto = self.help_hojas.get(clave, "Sin descripción disponible.")
        messagebox.showinfo(titulo, texto)

    def log(self, msg: str):
        self.log_text.config(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def agregar_archivos(self):
        rutas = filedialog.askopenfilenames(
            title="Seleccionar archivos Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")],
        )
        if not rutas:
            return

        anio_sel = self.anio_var.get()
        self.archivos.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)

        for ruta in rutas:
            info = {"ruta": ruta, "mes": None, "anio": None, "sucursal": None}
            parsed = parsear_nombre_archivo(os.path.basename(ruta))
            estado = "OK"

            if parsed:
                mes, anio, sucursal = parsed
                info["mes"] = mes
                info["anio"] = anio
                info["sucursal"] = sucursal if sucursal else "DESCONOCIDA"
                if anio != anio_sel:
                    estado = f"AÑO {anio}≠{anio_sel}"
            else:
                estado = "No detectado"

            self.archivos.append(info)

            self.tree.insert(
                "", "end",
                values=(ruta, info["mes"] or "", info["anio"] or "",
                        info["sucursal"] or "", estado),
            )

        self.log(f"{len(rutas)} archivos agregados.")

    def validar(self) -> bool:
        if not self.archivos:
            messagebox.showerror("Error", "Debes agregar al menos un archivo.")
            return False

        anio = self.anio_var.get()
        if anio < 2000 or anio > 2028:
            messagebox.showerror("Error", "El año debe estar entre 2000 y 2028.")
            return False

        meses_sel = [m for m, v in self.meses_vars.items() if v.get()]
        if not meses_sel:
            messagebox.showerror("Error", "Debes seleccionar al menos un mes.")
            return False

        for info in self.archivos:
            if not (info["mes"] and info["anio"] and info["sucursal"]):
                messagebox.showerror(
                    "Error",
                    f"Archivo sin datos completos (mes/año/sucursal):\n{info['ruta']}",
                )
                return False

        if self.lb_cols_sel.size() == 0:
            messagebox.showerror(
                "Error", "Debes seleccionar al menos una columna para el Consolidado."
            )
            return False

        return True

    # Gestión columnas
    def _col_agregar(self):
        sel = self.lb_cols_disp.curselection()
        if not sel:
            return
        col = self.lb_cols_disp.get(sel[0])
        existentes = [self.lb_cols_sel.get(i) for i in range(self.lb_cols_sel.size())]
        if col not in existentes:
            self.lb_cols_sel.insert("end", col)

    def _col_quitar(self):
        sel = self.lb_cols_sel.curselection()
        if not sel:
            return
        self.lb_cols_sel.delete(sel[0])

    def _col_agregar_custom(self):
        col = self.col_custom_var.get().strip()
        if not col:
            return
        existentes = [self.lb_cols_sel.get(i) for i in range(self.lb_cols_sel.size())]
        if col not in existentes:
            self.lb_cols_sel.insert("end", col)
        self.col_custom_var.set("")

    def _mover_col_arriba(self):
        sel = self.lb_cols_sel.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx == 0:
            return
        texto = self.lb_cols_sel.get(idx)
        self.lb_cols_sel.delete(idx)
        self.lb_cols_sel.insert(idx - 1, texto)
        self.lb_cols_sel.selection_set(idx - 1)

    def _mover_col_abajo(self):
        sel = self.lb_cols_sel.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx == self.lb_cols_sel.size() - 1:
            return
        texto = self.lb_cols_sel.get(idx)
        self.lb_cols_sel.delete(idx)
        self.lb_cols_sel.insert(idx + 1, texto)
        self.lb_cols_sel.selection_set(idx + 1)

    def _get_columnas_consolidado(self):
        return [self.lb_cols_sel.get(i) for i in range(self.lb_cols_sel.size())]

    # Procesamiento
    def procesar_async(self):
        if not self.validar():
            return
        self.btn_procesar.config(state="disabled")
        t = threading.Thread(target=self._procesar, daemon=True)
        t.start()

    def _procesar(self):
        try:
            anio = self.anio_var.get()
            meses_sel = [m for m, v in self.meses_vars.items() if v.get()]

            archivos_filtrados = [
                a for a in self.archivos if a["anio"] == anio and a["mes"] in meses_sel
            ]
            if not archivos_filtrados:
                messagebox.showerror(
                    "Error",
                    "No hay archivos que coincidan con año y meses seleccionados.",
                )
                return

            self.log("Iniciando consolidación...")
            df = consolidar_datos(archivos_filtrados)
            self.log(f"Datos consolidados: {len(df)} filas.")

            ts = datetime.now().strftime("%Y%m%d_%H%M")
            default_name = f"Ventas_Consolidadas_Final_{ts}.xlsx"
            ruta_salida = filedialog.asksaveasfilename(
                title="Guardar archivo de salida",
                defaultextension=".xlsx",
                initialfile=default_name,
                filetypes=[("Excel", "*.xlsx")],
            )
            if not ruta_salida:
                self.log("Operación cancelada por el usuario.")
                return

            columnas_consolidado = self._get_columnas_consolidado()
            deps = [d.strip() for d in self.especiales_deps_var.get().split(",") if d.strip()]
            marcas = [m.strip() for m in self.especiales_marcas_var.get().split(",") if m.strip()]
            filtros_especiales = (
                {"departamentos": deps, "marcas": marcas} if (deps or marcas) else None
            )

            self.log(f"Generando reportes en {ruta_salida}...")
            generar_reportes(
                df,
                ruta_salida,
                columnas_consolidado=columnas_consolidado,
                habilitar_ranking=self.habilitar_ranking_var.get(),
                habilitar_por_sucursal=self.habilitar_por_suc_var.get(),
                habilitar_matriz=self.habilitar_matriz_var.get(),
                habilitar_evolucion=self.habilitar_evolucion_var.get(),
                habilitar_especiales=self.habilitar_especiales_var.get(),
                filtros_especiales=filtros_especiales,
                habilitar_estacionalidad=self.habilitar_estacionalidad_var.get(),
            )

            self.log("✅ Proceso completado.")
            messagebox.showinfo("Listo", f"Archivo generado:\n{ruta_salida}")
        except Exception as e:
            self.log(f"ERROR: {e}")
            messagebox.showerror("Error", str(e))
        finally:
            self.btn_procesar.config(state="normal")


if __name__ == "__main__":
    root = tk.Tk()
    app = VentasConsolidatorGUI(root)
    root.mainloop()
