import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import docx
from texthumanize import humanize

class ParaphraserApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Anti-AI Paraphraser")
        self.root.geometry("600x400")
        self.root.resizable(False, False)

        # --- Variables ---
        self.file_path = tk.StringVar()
        self.language = tk.StringVar(value="en") # Inglés por defecto

        # --- UI Elements ---
        tk.Label(root, text="Selecciona un archivo (.txt o .docx):", font=("Arial", 10)).pack(pady=5)
        
        frame_file = tk.Frame(root)
        frame_file.pack(pady=5, fill=tk.X, padx=20)
        tk.Entry(frame_file, textvariable=self.file_path, state='readonly').pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(frame_file, text="Examinar", command=self.browse_file).pack(side=tk.RIGHT, padx=5)

        tk.Label(root, text="Idioma del texto:", font=("Arial", 10)).pack(pady=10)
        frame_lang = tk.Frame(root)
        frame_lang.pack(pady=5)
        tk.Radiobutton(frame_lang, text="Inglés", variable=self.language, value="en").pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(frame_lang, text="Español", variable=self.language, value="es").pack(side=tk.LEFT, padx=10)

        self.process_btn = tk.Button(root, text="Procesar Archivo", command=self.process_file, bg="#4CAF50", fg="white", font=("Arial", 12))
        self.process_btn.pack(pady=20)

        self.status_label = tk.Label(root, text="Listo para procesar.", fg="blue")
        self.status_label.pack(pady=5)

        # Barra de progreso (indeterminada, ya que el procesamiento puede tardar)
        self.progress = ttk.Progressbar(root, mode='indeterminate')
        self.progress.pack(pady=10, fill=tk.X, padx=20)

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Selecciona un archivo",
            filetypes=[("Archivos de texto", "*.txt"), ("Documentos de Word", "*.docx")]
        )
        if filename:
            self.file_path.set(filename)

    def process_file(self):
        filepath = self.file_path.get()
        if not filepath:
            messagebox.showwarning("Advertencia", "Por favor, selecciona un archivo primero.")
            return

        if not os.path.exists(filepath):
            messagebox.showerror("Error", "El archivo seleccionado no existe.")
            return

        lang = self.language.get()
        self.status_label.config(text="Procesando... Esto puede tardar unos minutos.", fg="orange")
        self.process_btn.config(state=tk.DISABLED)
        self.progress.start()

        # Ejecutar en un hilo separado para no congelar la GUI
        self.root.after(100, lambda: self._run_processing(filepath, lang))

    def _run_processing(self, filepath, lang):
        try:
            filename, ext = os.path.splitext(filepath)
            output_path = f"{filename}_paraphrased{ext}"

            if ext.lower() == '.txt':
                self._process_txt(filepath, output_path, lang)
            elif ext.lower() == '.docx':
                self._process_docx(filepath, output_path, lang)
            else:
                messagebox.showerror("Error", "Formato de archivo no soportado. Usa .txt o .docx")
                return

            self.status_label.config(text=f"¡Éxito! Archivo guardado en: {output_path}", fg="green")
            messagebox.showinfo("Completado", f"El archivo procesado se ha guardado como:\n{output_path}")

        except Exception as e:
            self.status_label.config(text="Ocurrió un error.", fg="red")
            messagebox.showerror("Error", f"Ocurrió un error durante el procesamiento:\n{e}")
        finally:
            self.progress.stop()
            self.process_btn.config(state=tk.NORMAL)

    def _process_txt(self, input_path, output_path, lang):
        with open(input_path, 'r', encoding='utf-8', errors='ignore') as f:
            text = f.read()

        # Usar TextHumanize para reescribir el texto
        result = humanize(text, lang=lang)
        paraphrased_text = result.text if hasattr(result, 'text') else str(result)

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(paraphrased_text)

    def _process_docx(self, input_path, output_path, lang):
        doc = docx.Document(input_path)

        for para in doc.paragraphs:
            if not para.text.strip():
                continue

            original_text = para.text
            # Reescribir el texto del párrafo
            result = humanize(original_text, lang=lang)
            paraphrased_text = result.text if hasattr(result, 'text') else str(result)

            # Reemplazar el texto manteniendo el estilo del párrafo.
            # Esto elimina el formato a nivel de "run" (negrita, cursiva parcial),
            # pero conserva el estilo general del párrafo.
            if para.runs:
                # Mantener el primer run y modificar su texto, eliminar los demás
                para.runs[0].text = paraphrased_text
                for run in para.runs[1:]:
                    run.text = ""
            else:
                para.text = paraphrased_text

        doc.save(output_path)

if __name__ == "__main__":
    root = tk.Tk()
    app = ParaphraserApp(root)
    root.mainloop()