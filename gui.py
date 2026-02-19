"""
gui.py - Interface graphique pour le convertisseur DocX vers LaTeX
"""

import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
from tkinterdnd2 import TkinterDnD, DND_FILES
from pathlib import Path
import threading
import queue
import sys
import io
from datetime import datetime
import os

from converter import DocxToLatexConverter
from config import Config


class TextRedirector(io.StringIO):
    """Redirige les print vers l'interface graphique."""
    
    def __init__(self, text_widget, tag=""):
        super().__init__()
        self.text_widget = text_widget
        self.tag = tag
    
    def write(self, string):
        self.text_widget.configure(state='normal')
        self.text_widget.insert(tk.END, string, (self.tag,))
        self.text_widget.configure(state='disabled')
        self.text_widget.see(tk.END)
        self.text_widget.update()


class ConverterGUI:
    """Interface graphique pour le convertisseur DocX vers LaTeX."""
    
    def __init__(self):
        # Créer la fenêtre principale avec support du drag and drop
        self.root = TkinterDnD.Tk()
        self.root.title("Convertisseur DocX vers LaTeX/PDF")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # Configuration
        self.config = Config()
        self.converter = DocxToLatexConverter(self.config)
        
        # Files à traiter
        self.files_to_process = []
        self.processing = False
        self.process_queue = queue.Queue()
        
        # Configurer le style
        self.setup_styles()
        
        # Créer l'interface
        self.create_widgets()
        
        # Configurer le drag and drop
        self.setup_drag_drop()
        
        # Centrer la fenêtre
        self.center_window()
        
        # Rediriger les prints
        self.redirect_output()

    
    def show_output_location(self):
        """Affiche l'emplacement des fichiers de sortie."""
        output_dir = self.converter.output_base_dir
        
        # Créer un message formaté
        location_msg = f"\n{'='*50}\n"
        location_msg += " EMPLACEMENT DES FICHIERS DE SORTIE\n"
        location_msg += f"{'='*50}\n"
        location_msg += f"Dossier principal: {output_dir}\n"
        location_msg += f"Fichiers LaTeX (.tex): {output_dir / 'LaTeX'}\n"
        location_msg += f"Fichiers PDF: {output_dir / 'PDF'}\n"
        location_msg += f"Images: {output_dir / 'LaTeX' / 'images'}\n"
        location_msg += f"{'='*50}\n\n"
        
        self.log_message(location_msg, 'info')
        
        # Sur macOS, proposer d'ouvrir le dossier
        if sys.platform == 'darwin':
            self.log_message(" Astuce: Pour ouvrir ce dossier:", 'info')
            self.log_message(f"   1. Ouvrez le Finder", 'info')
            self.log_message(f"   2. Cmd+Shift+G puis collez: {output_dir}", 'info')
            self.log_message("", 'info')
    
    def setup_styles(self):
        """Configure les styles de l'interface."""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Couleurs personnalisées
        self.colors = {
            'bg': '#f0f0f0',
            'drop_zone': '#e8e8e8',
            'drop_zone_hover': '#d0d0d0',
            'button': '#4CAF50',
            'button_hover': '#45a049',
            'danger': '#f44336',
            'text_bg': '#ffffff',
            'text_fg': '#333333'
        }
        
        self.root.configure(bg=self.colors['bg'])
    
    def create_widgets(self):
        """Crée tous les widgets de l'interface."""
        
        # Frame principale
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        # Titre
        title_label = ttk.Label(
            main_frame, 
            text="Convertisseur de PV DocX → LaTeX/PDF",
            font=('Arial', 16, 'bold')
        )
        title_label.grid(row=0, column=0, pady=(0, 20))
        
        # Zone de drag and drop
        self.create_drop_zone(main_frame)
        
        # Liste des fichiers
        self.create_file_list(main_frame)
        
        # Boutons d'action
        self.create_action_buttons(main_frame)
        
        # Zone de logs
        self.create_log_area(main_frame)
        
        # Barre de progression
        self.create_progress_bar(main_frame)
    
    def create_drop_zone(self, parent):
        """Crée la zone de drag and drop."""
        drop_frame = tk.Frame(
            parent, 
            bg=self.colors['drop_zone'],
            relief=tk.RIDGE,
            bd=2,
            height=120
        )
        drop_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        drop_frame.grid_propagate(False)
        
        self.drop_label = tk.Label(
            drop_frame,
            text=" Glissez vos fichiers .docx ici\nou cliquez pour parcourir",
            font=('Arial', 12),
            bg=self.colors['drop_zone'],
            fg=self.colors['text_fg'],
            cursor="hand2"
        )
        self.drop_label.pack(expand=True, fill=tk.BOTH)
        
        # Bind click event
        self.drop_label.bind("<Button-1>", lambda e: self.browse_files())
        
        # Stockage de la frame pour le drag and drop
        self.drop_frame = drop_frame
    
    def create_file_list(self, parent):
        """Crée la liste des fichiers sélectionnés."""
        # Label
        files_label = ttk.Label(parent, text="Fichiers sélectionnés:", font=('Arial', 11, 'bold'))
        files_label.grid(row=2, column=0, sticky=tk.W, pady=(0, 5))
        
        # Frame pour la liste et scrollbar
        list_frame = ttk.Frame(parent)
        list_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        parent.rowconfigure(3, weight=1)
        
        # Listbox avec scrollbar
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.file_listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            height=6,
            bg=self.colors['text_bg'],
            font=('Consolas', 10),
            selectmode=tk.EXTENDED
        )
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)
    
    def create_action_buttons(self, parent):
        """Crée les boutons d'action."""
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=4, column=0, pady=(0, 15))
        
        # Bouton Parcourir
        self.browse_btn = tk.Button(
            button_frame,
            text="📂 Parcourir",
            command=self.browse_files,
            bg=self.colors['button'],
            fg='white',
            font=('Arial', 10, 'bold'),
            padx=20,
            pady=8,
            cursor="hand2"
        )
        self.browse_btn.pack(side=tk.LEFT, padx=5)
        
        # Bouton Convertir
        self.convert_btn = tk.Button(
            button_frame,
            text="▶ Convertir",
            command=self.start_conversion,
            bg=self.colors['button'],
            fg='white',
            font=('Arial', 10, 'bold'),
            padx=20,
            pady=8,
            cursor="hand2",
            state=tk.DISABLED
        )
        self.convert_btn.pack(side=tk.LEFT, padx=5)
        
        # Bouton Effacer la sélection
        self.clear_btn = tk.Button(
            button_frame,
            text="🗑 Effacer",
            command=self.clear_selection,
            bg=self.colors['danger'],
            fg='white',
            font=('Arial', 10, 'bold'),
            padx=20,
            pady=8,
            cursor="hand2"
        )
        self.clear_btn.pack(side=tk.LEFT, padx=5)
    
    def create_log_area(self, parent):
        """Crée la zone d'affichage des logs."""
        # Label
        log_label = ttk.Label(parent, text="Journal de conversion:", font=('Arial', 11, 'bold'))
        log_label.grid(row=5, column=0, sticky=tk.W, pady=(0, 5))
        
        # Zone de texte scrollable
        self.log_text = scrolledtext.ScrolledText(
            parent,
            height=12,
            bg=self.colors['text_bg'],
            fg=self.colors['text_fg'],
            font=('Consolas', 9),
            state='disabled',
            wrap=tk.WORD
        )
        self.log_text.grid(row=6, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        parent.rowconfigure(6, weight=2)
        
        # Tags pour les couleurs
        self.log_text.tag_config('info', foreground='black')
        self.log_text.tag_config('success', foreground='green')
        self.log_text.tag_config('warning', foreground='orange')
        self.log_text.tag_config('error', foreground='red')
    
    def create_progress_bar(self, parent):
        """Crée la barre de progression."""
        self.progress = ttk.Progressbar(
            parent,
            mode='determinate',
            length=200
        )
        self.progress.grid(row=7, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Label de statut
        self.status_label = ttk.Label(parent, text="Prêt", font=('Arial', 9))
        self.status_label.grid(row=8, column=0)
    
    def setup_drag_drop(self):
        """Configure le drag and drop."""
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind('<<Drop>>', self.on_drop)
        self.drop_frame.dnd_bind('<<DragEnter>>', self.on_drag_enter)
        self.drop_frame.dnd_bind('<<DragLeave>>', self.on_drag_leave)
    
    def on_drop(self, event):
        """Gère le drop de fichiers."""
        files = self.root.tk.splitlist(event.data)
        self.add_files(files)
        self.drop_frame.configure(bg=self.colors['drop_zone'])
    
    def on_drag_enter(self, event):
        """Change l'apparence lors du survol."""
        self.drop_frame.configure(bg=self.colors['drop_zone_hover'])
    
    def on_drag_leave(self, event):
        """Restaure l'apparence normale."""
        self.drop_frame.configure(bg=self.colors['drop_zone'])
    
    def browse_files(self):
        """Ouvre le dialogue de sélection de fichiers."""
        files = filedialog.askopenfilenames(
            title="Sélectionner des fichiers DocX",
            filetypes=[("Documents Word", "*.docx"), ("Tous les fichiers", "*.*")]
        )
        if files:
            self.add_files(files)
    
    def add_files(self, files):
        """Ajoute des fichiers à la liste."""
        for file_path in files:
            if file_path.endswith('.docx'):
                if file_path not in self.files_to_process:
                    self.files_to_process.append(file_path)
                    self.file_listbox.insert(tk.END, Path(file_path).name)
                    self.log_message(f"Fichier ajouté: {Path(file_path).name}", 'info')
            else:
                self.log_message(f"Fichier ignoré (pas un .docx): {Path(file_path).name}", 'warning')
        
        # Activer le bouton convertir si des fichiers sont présents
        if self.files_to_process:
            self.convert_btn.configure(state=tk.NORMAL)
    
    def clear_selection(self):
        """Efface la sélection de fichiers."""
        self.files_to_process.clear()
        self.file_listbox.delete(0, tk.END)
        self.convert_btn.configure(state=tk.DISABLED)
        self.log_message("Sélection effacée", 'info')
    
    def start_conversion(self):
        """Lance la conversion dans un thread séparé."""
        if not self.files_to_process or self.processing:
            return
        
        self.processing = True
        self.convert_btn.configure(state=tk.DISABLED)
        self.browse_btn.configure(state=tk.DISABLED)
        self.clear_btn.configure(state=tk.DISABLED)
        
        # Lancer dans un thread séparé
        thread = threading.Thread(target=self.process_files, daemon=True)
        thread.start()
    
    def process_files(self):
        """Traite tous les fichiers sélectionnés."""
        total_files = len(self.files_to_process)
        self.progress['maximum'] = total_files
        
        self.log_message(f"\n{'='*50}", 'info')
        self.log_message(f"Début de la conversion de {total_files} fichier(s)", 'info')
        self.log_message(f"{'='*50}\n", 'info')
        
        self.log_message(f" Les fichiers seront créés dans: {self.converter.output_base_dir}\n", 'info')


        success_count = 0
        error_count = 0
        
        for i, file_path in enumerate(self.files_to_process, 1):
            file_name = Path(file_path).stem
            self.update_status(f"Conversion {i}/{total_files}: {Path(file_path).name}")
            
            self.log_message(f"\n[{i}/{total_files}] Conversion de: {Path(file_path).name}", 'info')
            
            try:
                # Conversion
                latex_path = str(Path(file_path).with_suffix('.tex'))
                pdf_result = self.converter.convert(file_path, latex_path, compile_pdf=True)

                self.log_message(f"✓ Conversion réussie: {Path(latex_path).name}", 'success')

                if pdf_result and Path(pdf_result).exists():
                    self.log_message(f"✓ PDF généré: {Path(pdf_result).name}", 'success')
                else:
                    self.log_message("⚠ PDF non généré (LaTeX non installé?)", 'warning')

                
                success_count += 1
                
            except Exception as e:
                self.log_message(f"✗ Erreur lors de la conversion: {str(e)}", 'error')
                error_count += 1
            
            # Mise à jour de la progression
            self.progress['value'] = i
            self.root.update_idletasks()
        
        # Résumé final
        self.log_message(f"\n{'='*50}", 'info')
        self.log_message(f"Conversion terminée: {success_count} réussi(s), {error_count} erreur(s)", 
                        'success' if error_count == 0 else 'warning')
        self.log_message(f"{'='*50}\n", 'info')
        
        self.update_status("Conversion terminée")
        
        # Réactiver les boutons
        self.processing = False
        self.convert_btn.configure(state=tk.NORMAL)
        self.browse_btn.configure(state=tk.NORMAL)
        self.clear_btn.configure(state=tk.NORMAL)
        self.progress['value'] = 0
    
    def log_message(self, message, tag='info'):
        """Affiche un message dans la zone de log."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        
        self.log_text.configure(state='normal')
        self.log_text.insert(tk.END, formatted_message, tag)
        self.log_text.configure(state='disabled')
        self.log_text.see(tk.END)
    
    def update_status(self, message):
        """Met à jour le label de statut."""
        self.status_label.configure(text=message)
    
    def center_window(self):
        """Centre la fenêtre sur l'écran."""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def redirect_output(self):
        """Redirige stdout et stderr vers la zone de log."""
        sys.stdout = TextRedirector(self.log_text, "info")
        sys.stderr = TextRedirector(self.log_text, "error")
    
    def run(self):
        """Lance l'interface graphique."""
        # Afficher l'emplacement des fichiers au démarrage
        self.root.after(500, self.show_output_location)
        
        self.root.mainloop()


def main():
    """Point d'entrée principal."""
    app = ConverterGUI()
    app.run()


if __name__ == "__main__":
    main()